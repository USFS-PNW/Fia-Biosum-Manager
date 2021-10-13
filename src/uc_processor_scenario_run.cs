using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace FIA_Biosum_Manager
{
   
    public partial class uc_processor_scenario_run : UserControl
    {
        public int m_intError = 0;
        public string m_strError = "";

        private ado_data_access m_oAdo;

        //list view associated classes
        private ListViewEmbeddedControls.ListViewEx m_lvEx;
        private ProgressBarEx.ProgressBarEx m_oProgressBarEx1;
        
        private ListViewAlternateBackgroundColors m_oLvAlternateColors = new ListViewAlternateBackgroundColors();

        int m_intLvCheckedCount = 0;
        int m_intLvTotalCount=0;
        int m_intTotalSteps = 75;
        int m_intCurrentStep = 0;
        int m_intCurrentCount = 0;
        string m_strDateTimeCreated = "";
        string m_strOPCOSTBatchFile="";
        private string m_strDebugFile = frmMain.g_oEnv.strTempDir + "\\biosum_processor_debug.txt";
        private string m_strOPCOSTRefPath;
        //@ToDo: Remove these when removing old Processor code; These variable stand-in for checkboxes that are removed
        bool m_blnLowSlope = true;
        bool m_blnSteepSlope = true;
        
        Queries m_oQueries = new Queries();
        Tables m_oTables = new Tables();
        RxTools m_oRxTools = new RxTools();
        excel_latebinding.excel_latebinding m_oExcel=null;
        FIA_Biosum_Manager.RxPackageItem_Collection m_oRxPackageItem_Collection = null;
        FIA_Biosum_Manager.RxItem_Collection m_oRxItem_Collection = null;
        FIA_Biosum_Manager.RxPackageItem m_oRxPackageItem = null;
        string m_strRxCycleList = "";
        string[] m_strFVSTreeTableLinkNameArray = null;
        IList<string> m_lstFVSVariantRxPackage = null;
        private ListViewColumnSorter lvwColumnSorter;
        /* Indicates that inactive stands will be included in the analysis:
         * There is a PRE_FVS_COMPUTE table with a column named ACTIVITY  */
        private bool m_bIncludeInactiveStands = true;
        private string m_strTempSqliteDbFile = null;
        private SQLite.ADO.DataMgr m_oDataMgr = new SQLite.ADO.DataMgr();

        //reference variables
        private string _strScenarioId = "";
        private frmProcessorScenario _frmProcessorScenario = null;
        private ProgressBarEx.ProgressBarEx _oProgressBarEx = null;
        private System.Collections.Generic.IList<string> _lstErrorVariants = null;
       
        private const int COL_CHECKBOX = 0;
        private const int COL_VARIANT = 1;
        private const int COL_PACKAGE = 2;
        private const int COL_RUNSTATUS = 3;
        private const int COL_VOLVAL = 4;
        private const int COL_HVSTCOST = 5;
        private const int COL_OPCOSTDROP = 6;
        private const int COL_CUTCOUNT = 7;
        private const int COL_RXCYCLE1 = 8;
        private const int COL_RXCYCLE2 = 9;
        private const int COL_RXCYCLE3 = 10;
        private const int COL_RXCYCLE4 = 11;
        private const int COL_FVSTREEFILE = 12;
        private const int COL_FOUND = 13;
        private const int COL_FVSTREE_PROCESSINGDATETIME = 14;
        private const int COL_PROCESSOR_PROCESSINGDATETIME = 15;
       

        public frmProcessorScenario ReferenceProcessorScenarioForm
        {
            get { return this._frmProcessorScenario; }
            set { this._frmProcessorScenario = value; }
        }
        public string ScenarioId
        {
            get { return _strScenarioId; }
            set { _strScenarioId = value; }
        }
        private ProgressBarEx.ProgressBarEx ReferenceProgressBarEx
        {
            get { return this._oProgressBarEx; }
            set { this._oProgressBarEx = value; }
        }

        public uc_processor_scenario_run()
        {
            InitializeComponent();
           
           			
        }
        static public class ScenarioHarvestMethodVariables
        {
            static private bool _bUseDefaultRxHarvestMethod = false;
            static public bool UseDefaultRxHarvestMethod
            {
                get { return _bUseDefaultRxHarvestMethod; }
                set {_bUseDefaultRxHarvestMethod = value; }
            }
            static private string _strHarvestMethodLowSlope;
            static public string HarvestMethodLowSlope
            {
                get { return _strHarvestMethodLowSlope; }
                set { _strHarvestMethodLowSlope = value; }
            }
            static private string _strHarvestMethodSteepSlope;
            static public string HarvestMethodSteepSlope
            {
                get { return _strHarvestMethodSteepSlope; }
                set { _strHarvestMethodSteepSlope = value; }
            }
            static private int _intSteepSlope;
            static public int SteepSlope
            {
                get { return _intSteepSlope; }
                set { _intSteepSlope = value; }
            }
            static private int _intSteepSlopeRecordCount=0;
            static public int SteepSlopeRecordCount
            {
                get { return _intSteepSlopeRecordCount; }
                set { _intSteepSlopeRecordCount = value; }
            }
            static private int _intLowSlopeRecordCount = 0;
            static public int LowSlopeRecordCount
            {
                get { return _intLowSlopeRecordCount; }
                set { _intLowSlopeRecordCount = value; }
            }
            static private bool _bProcessLowSlope = true;
            static public bool ProcessLowSlope
            {
                get { return _bProcessLowSlope; }
                set { _bProcessLowSlope = value; }
            }
            static private bool _bProcessSteepSlope = true;
            static public bool ProcessSteepSlope
            {
                get { return _bProcessSteepSlope; }
                set { _bProcessSteepSlope = value; }
            }            
           
        }

        static private class DiameterVariables
        {
            static private double _dblDiameter = 0.9;
            static public double diameter
            {
                get { return _dblDiameter; }
                set { _dblDiameter = value; }
            }
            static private double _dblMaxDia;
            static public double maxdia
            {
                get { return _dblMaxDia; }
                set { _dblMaxDia = value; }
            }
            static private double _dblMinDiaChips;
            static public double MinDiaChips
            {
                get { return _dblMinDiaChips; }
                set { _dblMinDiaChips = value; }
            }
            static private double _dblMaxDiaChips;
            static public double MaxDiaChips
            {
                get { return _dblMaxDiaChips; }
                set { _dblMaxDiaChips = value; }
            }
            static private double _dblMinDiaSmLogs;
            static public double MinDiaSmLogs
            {
                get { return _dblMinDiaSmLogs; }
                set { _dblMinDiaSmLogs = value; }
            }
            static private double _dblMaxDiaSmLogs;
            static public double MaxDiaSmLogs
            {
                get { return _dblMaxDiaSmLogs; }
                set { _dblMaxDiaSmLogs = value; }
            }
            static private double _dblMinDiaLgLogs;
            static public double MinDiaLgLogs
            {
                get { return _dblMinDiaLgLogs; }
                set { _dblMinDiaLgLogs = value; }
            }
            static private double _dblMaxDiaLgLogs;
            static public double MaxDiaLgLogs
            {
                get { return _dblMaxDiaLgLogs; }
                set { _dblMaxDiaLgLogs = value; }
            }
            static private double _dblMinDiaSteepSlope;
            static public double MinDiaSteepSlope
            {
                get { return _dblMinDiaSteepSlope; }
                set { _dblMinDiaSteepSlope = value; }
            }
            static private string _strDiaClass_BC;
            static public string DiaClass_BC
            {
                get { return _strDiaClass_BC; }
                set { _strDiaClass_BC = value; }
            }
            static private string _strDiaClass_CHIPS;
            static public string DiaClass_CHIPS
            {
                get { return _strDiaClass_CHIPS; }
                set { _strDiaClass_CHIPS = value; }
            }
            static private string _strDiaClass_SMLOGS;
            static public string DiaClass_SMLOGS
            {
                get { return _strDiaClass_SMLOGS; }
                set { _strDiaClass_SMLOGS = value; }
            }
            static private string _strDiaClass_LGLOGS;
            static public string DiaClass_LGLOGS
            {
                get { return _strDiaClass_LGLOGS; }
                set { _strDiaClass_LGLOGS = value; }
            }
            static private string _strDiaClass_AllLOGS;
            static public string DiaClass_AllLOGS
            {
                get { return _strDiaClass_AllLOGS; }
                set { _strDiaClass_AllLOGS = value; }
            }
            static private double _dblDiaCount;
            static public double DiaCount
            {
                get { return _dblDiaCount; }
                set { _dblDiaCount = value; }
            }
           
        }
        public void loadvalues()
        {
            
            int x, y;
            string strTableName="";
            string strFVSTreeTableLinkNameList="";
            string strErrMsg="";
            bool bUpdate;
            string strDebugFile = frmMain.g_oEnv.strTempDir + "\\biosum_processor_scenario_run_loadvalues.txt";
            if (System.IO.File.Exists(strDebugFile))
                System.IO.File.Delete(strDebugFile);

            System.Threading.Thread.Sleep(2000);

            if (frmMain.g_bDebug)
                frmMain.g_oUtils.WriteText(strDebugFile, "*****START*****" + System.DateTime.Now.ToString() + "\r\n");

            // INITIALIZE OPCOST REF PATH WHEN WE LOAD THE SCENARIO 
            m_strOPCOSTRefPath = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + 
                "\\" + Tables.Reference.DefaultOpCostReferenceDbFile;
            
            if (m_oAdo != null && m_oAdo.m_OleDbConnection != null)
                if (m_oAdo.m_OleDbConnection.State == System.Data.ConnectionState.Open) m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            cmbFilter.Items.Clear();
            cmbFilter.Items.Add("All");
            cmbFilter.Text = "All";
            //
            //INSTANTIATE EXTENDED LISTVIEW OBJECT
            //
            m_lvEx = new ListViewEmbeddedControls.ListViewEx();
            m_lvEx.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(m_lvEx_ItemCheck);
            m_lvEx.MouseUp += new System.Windows.Forms.MouseEventHandler(m_lvEx_MouseUp);
            m_lvEx.SelectedIndexChanged += new System.EventHandler(m_lvEx_SelectedIndexChanged);
            m_lvEx.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(m_lvEx_ColumnClick);
            this.panel1.Controls.Add(m_lvEx);
            m_lvEx.Size = this.lstFvsOutput.Size;
            m_lvEx.Location = this.lstFvsOutput.Location;
            m_lvEx.CheckBoxes = true;
            m_lvEx.AllowColumnReorder = false;
            m_lvEx.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            m_lvEx.FullRowSelect = false;
            m_lvEx.MultiSelect = false;
            m_lvEx.GridLines = true;
            this.m_lvEx.HideSelection = false;
            m_lvEx.View = System.Windows.Forms.View.Details;
            this.lstFvsOutput.Hide();
            //
            //INITIALIZE LISTVIEW ALTERNATE ROW COLORS
            //
            System.Windows.Forms.ListViewItem entryListItem = null;
            this.m_oLvAlternateColors.InitializeRowCollection();
            this.m_oLvAlternateColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceListView = m_lvEx;
            this.m_oLvAlternateColors.CustomFullRowSelect = true;
            m_oLvAlternateColors.ColumnsToNotUpdate(COL_RUNSTATUS.ToString());
            if (frmMain.g_oGridViewFont != null) m_lvEx.Font = frmMain.g_oGridViewFont;
            //
            //ASSIGN LISTVIEW COLUMN LABELS
            //
            m_lvEx.Show();
            this.m_lvEx.Clear();
            this.m_lvEx.Columns.Add(" ", 10, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("FVS Variant", 100, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("RxPackage", 100, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("Run Status", 250, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("TreeVolValRecordCount", 100, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("HarvestCostRecordCount", 100, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("OpcostDropped", 100, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("TreeCutListRecordCount", 100, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("Cycle1Rx", 80, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("Cycle2Rx", 80, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("Cycle3Rx", 80, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("Cycle4Rx", 80, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("FVSTreeFile", 100, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("File Found", 80, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("FVStree_DateTimeCreated", 80, HorizontalAlignment.Left);
            this.m_lvEx.Columns.Add("Processor_DateTimeCreated", 80, HorizontalAlignment.Left);
            this.m_lvEx.Columns[COL_CHECKBOX].Width = -2;
            this.m_lvEx.Columns[COL_VOLVAL].Width = -2;
            this.m_lvEx.Columns[COL_HVSTCOST].Width = -2;
            this.m_lvEx.Columns[COL_FVSTREE_PROCESSINGDATETIME].Width = -2;
            this.m_lvEx.Columns[COL_PROCESSOR_PROCESSINGDATETIME].Width = -2;
            // Create an instance of a ListView column sorter and assign it 
            // to the ListView control.
            lvwColumnSorter = new ListViewColumnSorter();
            this.m_lvEx.ListViewItemSorter = lvwColumnSorter;
            //
            //SCENARIO ID
            //
            ScenarioId = this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
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
                "\\processor\\" + ScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsDbFile;
            if (ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                strScenarioMDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                strScenarioResultsMDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + ScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultSqliteResultsDbFile;
            }
            //
            //LOAD PROJECT DATATASOURCES INFO
            //
            m_oQueries.m_oFvs.LoadDatasource = true;
            m_oQueries.m_oReference.LoadDatasource = true;
            m_oQueries.m_oProcessor.LoadDatasource = true;
            // The following call creates the temp database that is used for running calculations
            m_oQueries.LoadDatasources(true, ReferenceProcessorScenarioForm.m_bUsingSqlite, "processor", ScenarioId);
            //
            //LOAD RX PACKAGE INFO
            //
            //load rxpackage properties
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "START: LoadAllRxPackageItemsFromTableIntoRxPackageCollection - " + System.DateTime.Now.ToString() + "\r\n");
            m_oRxPackageItem_Collection = new RxPackageItem_Collection();
            m_oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(m_oQueries.m_strTempDbFile, m_oQueries, this.m_oRxPackageItem_Collection);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "END: LoadAllRxPackageItemsFromTableIntoRxPackageCollection - " + System.DateTime.Now.ToString() + "\r\n");

            //
            //LOAD RX INFO
            //
            //load rx properties
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "START: LoadAllRxItemsFromTableIntoRxCollection - " + System.DateTime.Now.ToString() + "\r\n");
            m_oRxItem_Collection = new RxItem_Collection();
            m_oRxTools.LoadAllRxItemsFromTableIntoRxCollection(m_oQueries.m_strTempDbFile, m_oQueries, m_oRxItem_Collection);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "END: LoadAllRxItemsFromTableIntoRxCollection - " + System.DateTime.Now.ToString() + "\r\n");
            //
            //CREATE LINK IN TEMP MDB TO ALL PROCESSOR SCENARIO TABLES
            //
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "START: Create Links to the Scenario tables - " + System.DateTime.Now.ToString() + "\r\n");
            dao_data_access oDao = new dao_data_access();
            ODBCMgr odbcmgr = new ODBCMgr();
            //link to all the scenario rule definition tables
            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                    "scenario_cost_revenue_escalators",
                    strScenarioMDB, "scenario_cost_revenue_escalators", true);
                oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                    "scenario_additional_harvest_costs",
                    strScenarioMDB, "scenario_additional_harvest_costs", true);
                oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                   "scenario_harvest_cost_columns",
                   strScenarioMDB, "scenario_harvest_cost_columns", true);
                oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                  "scenario_harvest_method",
                  strScenarioMDB, "scenario_harvest_method", true);
                oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName,
                    strScenarioMDB,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName, true);
                oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                 "scenario_tree_species_diam_dollar_values",
                 strScenarioMDB, "scenario_tree_species_diam_dollar_values", true);
                oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName,
                    strScenarioMDB,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName, true);
                oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName,
                    strScenarioMDB,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName, true);
            }
            else
            {
                // Check to see if the input SQLite DSN exists and if so, delete so we can add
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName);
                }
                odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName,
                    this.ReferenceProcessorScenarioForm.LoadedQueries.m_strTempDbFile);
                string strScenarioDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                //oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, Tables.ProcessorScenarioRuleDefinitions.DefaultCostRevenueEscalatorsTableName,
                //    Tables.ProcessorScenarioRuleDefinitions.DefaultCostRevenueEscalatorsTableName, ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName,
                //    strScenarioDB);
                //oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName,
                //    Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName, ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName,
                //    strScenarioDB);
                //oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestCostColumnsTableName,
                //    Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestCostColumnsTableName, ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName,
                //    strScenarioDB);
                //oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName,
                //    Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName, ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName,
                //    strScenarioDB);
                //oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName,
                //    Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName, ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName,
                //    strScenarioDB);
                //oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName,
                //    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName, ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName,
                //    strScenarioDB);
                //oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName,
                //    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName, ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName,
                //    strScenarioDB);
                oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName, ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName,
                    strScenarioDB);
            }

            //link scenario results tables
            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                    Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName,
                    strScenarioResultsMDB,
                    Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, true);
                oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                    Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName,
                    strScenarioResultsMDB,
                    Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, true);
            }
            else
            {
                //Check to see if the input SQLite DSN exists and if so, delete so we can add
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorResultsDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorResultsDsnName);
                }
                odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.ProcessorResultsDsnName,
                    strScenarioResultsMDB);
                oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName,
                    Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, ODBCMgr.DSN_KEYS.ProcessorResultsDsnName,
                    strScenarioResultsMDB);
                oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName,
                    Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, ODBCMgr.DSN_KEYS.ProcessorResultsDsnName,
                    strScenarioResultsMDB);
            }

            // link to PRE_FVS_SUMMARY table
            oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                Tables.FVS.DefaultPreFVSSummaryTableName,
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.FVS.DefaultPreFVSSummaryDbFile,
                Tables.FVS.DefaultPreFVSSummaryTableName, true);
            // link to PRE_FVS_COMPUTE table
            string strComputeDbPath = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultPreFVSComputeDbFile;
            if (System.IO.File.Exists(strComputeDbPath) &&
                oDao.TableExists(strComputeDbPath, Tables.FVS.DefaultPreFVSComputeTableName))
            {
                oDao.CreateTableLink(m_oQueries.m_strTempDbFile,
                    Tables.FVS.DefaultPreFVSComputeTableName,
                    strComputeDbPath, Tables.FVS.DefaultPreFVSComputeTableName, true);
            }
            else
            {
                m_bIncludeInactiveStands = false;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(strDebugFile, "loadvalues(): PRE_FVS_COMPUTE table was not found. Stands with no " +
                        "activity will not be included in analysis! \r\n");
            }

            oDao.m_DaoDbEngine.Idle(1);
            oDao.m_DaoDbEngine.Idle(8);
            oDao.m_DaoWorkspace.Close();
            oDao.m_DaoDbEngine = null;
            oDao = null;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "END: Create Links to the Scenario tables - " + System.DateTime.Now.ToString() + "\r\n");

            //
            //CREATE LINK IN TEMP MDB TO ALL VARIANT CUTLIST TABLES
            //
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "START: CreateTableLinksToFVSOutTreeListTables - " + System.DateTime.Now.ToString() + "\r\n");
            m_oRxTools.CreateTableLinksToFVSOutTreeListTables(m_oQueries, m_oQueries.m_strTempDbFile);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "END: CreateTableLinksToFVSOutTreeListTables - " + System.DateTime.Now.ToString() + "\r\n");
            //
            //OPEN CONNECTION TO TEMP DB FILE
            //
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));
            if (m_oAdo.m_intError == 0)
            {
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "ProcessorVariantPackageDateTimeCreated_work_table"))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE ProcessorVariantPackageDateTimeCreated_work_table");
                //
                //create temp work table with datetimecreated
                //
                m_oAdo.m_strSQL = "SELECT DISTINCT biosum_cond_id,rxpackage,DateTimeCreated " +
                                 "INTO ProcessorVariantPackageDateTimeCreated_work_table " +
                                 "FROM " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oAdo.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.AddColumn(m_oAdo.m_OleDbConnection,
                                    "ProcessorVariantPackageDateTimeCreated_work_table",
                                    "fvs_variant", "text", "2");
                m_oAdo.AddIndex(m_oAdo.m_OleDbConnection, "ProcessorVariantPackageDateTimeCreated_work_table", "ProcessorVariantPackageDateTimeCreated_work_table_idx1", "biosum_cond_id");

                m_oAdo.m_strSQL = "UPDATE ProcessorVariantPackageDateTimeCreated_work_table a " +
                                  "INNER JOIN (" + m_oQueries.m_oFIAPlot.m_strPlotTable + " p " +
                                         "INNER JOIN " + m_oQueries.m_oFIAPlot.m_strCondTable + " c " +
                                         "ON p.biosum_plot_id=c.biosum_plot_id) " +
                                  "ON trim(a.biosum_cond_id)=c.biosum_cond_id SET a.fvs_variant = p.fvs_variant;";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oAdo.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");

                //
                //GET LIST OF VARIANTS
                //
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "START: GetListOfFVSVariantsInPlotTable - " + System.DateTime.Now.ToString() + "\r\n");
                string strVariantsList = m_oRxTools.GetListOfFVSVariantsInPlotTable(m_oAdo, m_oAdo.m_OleDbConnection, m_oQueries.m_oFIAPlot.m_strPlotTable);
                string[] strVariantsArray = frmMain.g_oUtils.ConvertListToArray(strVariantsList, ",");
                //find the variants that have tree cut list tables
                m_lstFVSVariantRxPackage = new List<string>();
                for (x = 0; x <= strVariantsArray.Length - 1; x++)
                {
                    cmbFilter.Items.Add(strVariantsArray[x].Trim());
                    for (y = 0; y <= m_oRxPackageItem_Collection.Count - 1; y++)
                    {
                        strTableName = "fvs_tree_IN_" + strVariantsArray[x].Trim() + "_P" + m_oRxPackageItem_Collection.Item(y).RxPackageId + "_TREE_CUTLIST";
                        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, strTableName))
                        {
                            strFVSTreeTableLinkNameList = strFVSTreeTableLinkNameList + strTableName + ",";
                            m_lstFVSVariantRxPackage.Add(strVariantsArray[x].Trim() + "," + m_oRxPackageItem_Collection.Item(y).RxPackageId);
                        }
                    }
                }
                //load the list into an array
                if (strFVSTreeTableLinkNameList.Trim().Length > 0)
                {
                    strFVSTreeTableLinkNameList=strFVSTreeTableLinkNameList.Substring(0,strFVSTreeTableLinkNameList.Length-1);
                    m_strFVSTreeTableLinkNameArray = frmMain.g_oUtils.ConvertListToArray(strFVSTreeTableLinkNameList,",");
                }
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "END: GetListOfFVSVariantsInPlotTable - " + System.DateTime.Now.ToString() + "\r\n");

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "START: Populate List " + System.DateTime.Now.ToString() + "\r\n");

                //Check for ACTIVITY column so we know if we should include inactive stands
                if (m_bIncludeInactiveStands == true)
                {
                    if (!m_oAdo.ColumnExist(m_oAdo.m_OleDbConnection, Tables.FVS.DefaultPreFVSComputeTableName, "ACTIVITY"))
                    {
                        m_bIncludeInactiveStands = false;
                    }
                }
    
                //populate the listview object
                for (x = 0; x <= m_strFVSTreeTableLinkNameArray.Length - 1; x++)
                {
                    string strRxPackage = "";
                    string strCount = "";
                    string strVariant = "";
                    if (frmMain.g_bSuppressProcessorScenarioTableRowCount == false)
                    {
                        m_oAdo.m_strSQL = "SELECT rxpackage, fvs_variant,COUNT(*) AS rxpackage_variant_count " +
                                          "FROM " + m_strFVSTreeTableLinkNameArray[x] + " " +
                                          "GROUP BY rxpackage,fvs_variant";
                    }
                    else
                    {
                        m_oAdo.m_strSQL = "SELECT DISTINCT rxpackage,fvs_variant, 1 AS rxpackage_variant_count FROM " + m_strFVSTreeTableLinkNameArray[x];
                    }
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oAdo.m_strSQL + " "  + System.DateTime.Now.ToString() + "\r\n");
                    m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");

                    /* Method-level variable indicating that this variant-rxPackage is an inactive combination:
                     * There are no rows in the fvs_cutlist table */
                    bool _bInactiveVarRxPackage = false;
                    if (m_oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (m_oAdo.m_OleDbDataReader.Read())
                        {
                            if (m_oAdo.m_OleDbDataReader["rxpackage"] != System.DBNull.Value)
                                strRxPackage = m_oAdo.m_OleDbDataReader["rxpackage"].ToString().Trim();
                            if (m_oAdo.m_OleDbDataReader["fvs_variant"] != System.DBNull.Value)
                                strVariant = m_oAdo.m_OleDbDataReader["fvs_variant"].ToString().Trim();
                            if (m_oAdo.m_OleDbDataReader["rxpackage_variant_count"] != System.DBNull.Value)
                                strCount = m_oAdo.m_OleDbDataReader["rxpackage_variant_count"].ToString().Trim();
                        }
                        m_oAdo.m_OleDbDataReader.Dispose();
                    }
                    else if (m_bIncludeInactiveStands == true)
                    {
                        string[] arrVariantRxPkg = frmMain.g_oUtils.ConvertListToArray(m_lstFVSVariantRxPackage[x],",");
                        int count = m_oAdo.getRecordCount(m_oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM " + Tables.FVS.DefaultPreFVSComputeTableName + " WHERE fvs_variant = '" +
                                        arrVariantRxPkg[0] + "' AND rxpackage = '" + arrVariantRxPkg[1] + "'", Tables.FVS.DefaultPreFVSComputeTableName);
                        if (count > 0)
                        {
                            strVariant = arrVariantRxPkg[0];
                            strRxPackage = arrVariantRxPkg[1];
                            strCount = "1";
                            _bInactiveVarRxPackage = true;
                        }
                    }

                    if (!String.IsNullOrEmpty(strVariant))
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(strDebugFile, "Add To List Variant:" + strVariant + " RxPackage:" + strRxPackage + " " + System.DateTime.Now.ToString() + "\r\n");

                        //find the package item
                        for (y = 0; y <= m_oRxPackageItem_Collection.Count - 1; y++)
                        {
                            if (m_oRxPackageItem_Collection.Item(y).RxPackageId.Trim() == strRxPackage.Trim())
                            {
                                break;
                            }
                        }
                        if (y <= m_oRxPackageItem_Collection.Count - 1)
                        {
                            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Loading Scenario Run Data (Variant:" + strVariant + " RxPackage:" + strRxPackage + ")...Stand By");
                            frmMain.g_oDelegate.ExecuteStatusBarPanelMethod(frmMain.g_sbpInfo.Parent, 1, "Refresh");
                            bUpdate = false;
                            //found package item
                            //create listview row
                            // Add a ListItem object to the ListView.
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(strDebugFile, "Checkpoint 1 " + System.DateTime.Now.ToString() + "\r\n");
                            entryListItem =
                                m_lvEx.Items.Add(" ");

                            entryListItem.UseItemStyleForSubItems = false;

                            this.m_oLvAlternateColors.AddRow();
                            this.m_oLvAlternateColors.AddColumns(m_lvEx.Items.Count - 1, m_lvEx.Columns.Count);
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(strDebugFile, "Checkpoint 2 " + System.DateTime.Now.ToString() + "\r\n");

                            //variant
                            entryListItem.SubItems.Add(strVariant);
                            //rxpackage
                            entryListItem.SubItems.Add(strRxPackage);
                            //progress bar
                            entryListItem.SubItems.Add(" ");
                            this.m_oProgressBarEx1 = new ProgressBarEx.ProgressBarEx(Color.Gold);
                            this.m_oProgressBarEx1.MarqueePercentage = 25;
                            this.m_oProgressBarEx1.MarqueeSpeed = 30;
                            this.m_oProgressBarEx1.MarqueeStep = 1;
                            this.m_oProgressBarEx1.Maximum = 100;
                            this.m_oProgressBarEx1.Minimum = 0;
                            this.m_oProgressBarEx1.Name = "m_oProgressBarEx1";
                            this.m_oProgressBarEx1.ProgressPadding = 0;
                            this.m_oProgressBarEx1.ProgressType = ProgressBarEx.ProgressType.Smooth;
                            this.m_oProgressBarEx1.ShowPercentage = true;
                            this.m_oProgressBarEx1.BackColor = Color.LawnGreen;



                            this.m_oProgressBarEx1.BackgroundColor = Color.Black;
                            this.m_oProgressBarEx1.TabIndex = 18;
                            this.m_oProgressBarEx1.Text = "0%";
                            this.m_oProgressBarEx1.Value = 0;
                            m_lvEx.AddEmbeddedControl(this.m_oProgressBarEx1, COL_RUNSTATUS, m_lvEx.Items.Count - 1, System.Windows.Forms.DockStyle.Fill);
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(strDebugFile, "Checkpoint 3" + System.DateTime.Now.ToString() + "\r\n");

                            //tree volval count
                            m_oAdo.m_strSQL = "SELECT COUNT(*) as rowcount " +
                                              "FROM " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " t," +
                                                        m_oQueries.m_oFIAPlot.m_strCondTable + " c," +
                                                        m_oQueries.m_oFIAPlot.m_strPlotTable + " p " +
                                              "WHERE t.rxpackage='" + strRxPackage + "' AND " +
                                                    "(t.biosum_cond_id=c.biosum_cond_id AND " +
                                                     "p.biosum_plot_id=c.biosum_plot_id AND " +
                                                     "p.fvs_variant='" + strVariant + "')";
                            if (frmMain.g_bSuppressProcessorScenarioTableRowCount == false && _bInactiveVarRxPackage == false)
                            {
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oAdo.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");

                                entryListItem.SubItems.Add(Convert.ToString(m_oAdo.getRecordCount(m_oAdo.m_OleDbConnection,
                                    m_oAdo.m_strSQL, "temp")));
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");

                            }
                            else
                                entryListItem.SubItems.Add(" ");
                            //tree harvest cost count
                            m_oAdo.m_strSQL = "SELECT COUNT(*) as rowcount " +
                                              "FROM " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " t," +
                                                        m_oQueries.m_oFIAPlot.m_strCondTable + " c," +
                                                        m_oQueries.m_oFIAPlot.m_strPlotTable + " p " +
                                              "WHERE t.rxpackage='" + strRxPackage + "' AND " +
                                                    "(t.biosum_cond_id=c.biosum_cond_id AND " +
                                                     "p.biosum_plot_id=c.biosum_plot_id AND " +
                                                     "p.fvs_variant='" + strVariant + "')";
                            if (frmMain.g_bSuppressProcessorScenarioTableRowCount == false && _bInactiveVarRxPackage == false)
                            {
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oAdo.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");

                                entryListItem.SubItems.Add(Convert.ToString(m_oAdo.getRecordCount(m_oAdo.m_OleDbConnection,
                                    m_oAdo.m_strSQL, "temp")));

                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(strDebugFile, "END SQL: " + System.DateTime.Now.ToString() + "\r\n");

                            }
                            else
                                entryListItem.SubItems.Add(" ");
                            //opcost dropped column count; Always empty
                            entryListItem.SubItems.Add(" ");
                            //tree cutlist count
                            if (frmMain.g_bSuppressProcessorScenarioTableRowCount == false && !_bInactiveVarRxPackage)
                                entryListItem.SubItems.Add(strCount);
                            else if (_bInactiveVarRxPackage)
                            {
                                entryListItem.SubItems.Add("0");
                            }
                            else
                                entryListItem.SubItems.Add("> 0");
                            //cycle1 rx
                            if (this.m_oRxPackageItem_Collection.Item(y).SimulationYear1Rx.Trim().Length > 0)
                                entryListItem.SubItems.Add(this.m_oRxPackageItem_Collection.Item(y).SimulationYear1Rx);
                            else
                                entryListItem.SubItems.Add("000");
                            //cycle2 rx
                            if (this.m_oRxPackageItem_Collection.Item(y).SimulationYear2Rx.Trim().Length > 0)
                                entryListItem.SubItems.Add(this.m_oRxPackageItem_Collection.Item(y).SimulationYear2Rx);
                            else
                                entryListItem.SubItems.Add("000");
                            //cycle3 rx
                            if (this.m_oRxPackageItem_Collection.Item(y).SimulationYear3Rx.Trim().Length > 0)
                                entryListItem.SubItems.Add(this.m_oRxPackageItem_Collection.Item(y).SimulationYear3Rx);
                            else
                                entryListItem.SubItems.Add("000");
                            //cycle4 rx
                            if (this.m_oRxPackageItem_Collection.Item(y).SimulationYear4Rx.Trim().Length > 0)
                                entryListItem.SubItems.Add(this.m_oRxPackageItem_Collection.Item(y).SimulationYear4Rx);
                            else
                                entryListItem.SubItems.Add("000");

                            entryListItem.SubItems.Add(strVariant + "_P" + strRxPackage + "_tree_cutlist.mdb");  //out mdb file
                            entryListItem.SubItems.Add("Found");  //file found
                            //fvstree processing date and time variant,rxpackage
                            string strFvsTreeDateCreated = " ";
                            if (_bInactiveVarRxPackage == false)
                            {
                                m_oAdo.m_strSQL = "SELECT DISTINCT DateTimeCreated " +
                                              "FROM " + m_strFVSTreeTableLinkNameArray[x] + " t " +
                                              "WHERE t.rxpackage='" + strRxPackage + "'";

                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oAdo.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");

                                strFvsTreeDateCreated = (string)m_oAdo.getSingleStringValueFromSQLQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, "temp");

                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(strDebugFile, "END SQL: " + System.DateTime.Now.ToString() + "\r\n");
                            }

                            entryListItem.SubItems.Add((strFvsTreeDateCreated.Trim())); //date and time created


                            m_oAdo.m_strSQL = "SELECT DateTimeCreated " +
                                              "FROM  ProcessorVariantPackageDateTimeCreated_work_table " +
                                              "WHERE rxpackage='" + strRxPackage + "' AND " +
                                                    "fvs_variant='" + strVariant + "'";

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oAdo.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");

                            m_oAdo.m_strSQL = (string)m_oAdo.getSingleStringValueFromSQLQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, "temp");

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(strDebugFile, "END SQL: " + System.DateTime.Now.ToString() + "\r\n");

                            //fvstree processing date and time variant,rxpackage 
                            if (m_oAdo.m_strSQL.Trim().Length > 0)
                            {
                                entryListItem.SubItems.Add(m_oAdo.m_strSQL); //date and time created
                            }
                            else
                            {
                                entryListItem.SubItems.Add(" ");
                            }

                            this.m_oLvAlternateColors.m_oRowCollection.Item(m_oLvAlternateColors.m_oRowCollection.Count - 1).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn = false;

                            if (entryListItem.SubItems[COL_CUTCOUNT].Text.Trim() == "> 0" ||
                                Convert.ToInt32(entryListItem.SubItems[COL_CUTCOUNT].Text.Trim()) > 0)
                            {
                                if (entryListItem.SubItems[COL_FVSTREE_PROCESSINGDATETIME].Text.Trim().Length == 0 ||
                                    entryListItem.SubItems[COL_PROCESSOR_PROCESSINGDATETIME].Text.Trim().Length == 0)
                                {
                                    bUpdate = true;
                                }
                                else
                                {
                                    if (Convert.ToDateTime(entryListItem.SubItems[COL_FVSTREE_PROCESSINGDATETIME].Text.Trim()) >
                                        Convert.ToDateTime(entryListItem.SubItems[COL_PROCESSOR_PROCESSINGDATETIME].Text.Trim()))
                                    {
                                        bUpdate = true;
                                    }
                                }
                            }


                            this.m_oLvAlternateColors.ListViewItem(m_lvEx.Items[m_lvEx.Items.Count - 1]);
                            if (bUpdate && _bInactiveVarRxPackage == false)
                            {
                                this.lblMsg.Text = "* = New FVS Tree records to process";
                                if (this.lblMsg.Visible == false) lblMsg.Show();
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(m_lvEx, m_lvEx.Items.Count - 1, 0, "Text", "*");
                            }

                        }
                        else
                        {
                            //did not find package item so display error
                            strErrMsg = strErrMsg + "Table " + m_strFVSTreeTableLinkNameArray[x] + " contains RXPACKAGE: " + strRxPackage + " but " + strRxPackage + " is not a defined package. \r\n";
                        }
                    }

                }
                if (strErrMsg.Trim().Length > 0)
                {
                    MessageBox.Show("Error loading Processor Scenario run list \r\n" + strErrMsg, "FIA Biosum",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "END: Populate List " + System.DateTime.Now.ToString() + "\r\n");



                m_oAdo.m_OleDbDataReader.Dispose();
            }
            this.panel1_Resize();
            if (frmMain.g_bDebug)
                frmMain.g_oUtils.WriteText(strDebugFile, "*****END*****" + System.DateTime.Now.ToString() + "\r\n");
        }
        private void m_lvEx_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
        {
            m_lvEx.Items[e.Index].Selected = true;
        }
        private void m_lvEx_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            int x;
            try
            {
                if (e.Button == MouseButtons.Left)
                {
                    int intRowHt = this.m_lvEx.Items[0].Bounds.Height;
                    double dblRow = (double)(e.Y / intRowHt);
                    this.m_lvEx.Items[m_lvEx.TopItem.Index + (int)dblRow - 1].Selected = true;
                    this.m_oLvAlternateColors.DelegateListViewItem(m_lvEx.Items[m_lvEx.TopItem.Index + (int)dblRow - 1]);
                    
                   
                }
            }
            catch
            {
            }
        }

        private void m_lvEx_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (m_lvEx.SelectedItems.Count > 0)
                m_oLvAlternateColors.DelegateListViewItem(m_lvEx.SelectedItems[0]);
        }
        private void m_lvEx_ColumnClick(object sender, System.Windows.Forms.ColumnClickEventArgs e)
        {
            int x, y;

            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.m_lvEx.Sort();
            //reinitialize the alternate row colors
            for (x = 0; x <= this.m_lvEx.Items.Count - 1; x++)
            {
                for (y = 0; y <= this.m_lvEx.Columns.Count - 1; y++)
                {
                   m_oLvAlternateColors.ListViewSubItem(this.m_lvEx.Items[x].Index, y, this.m_lvEx.Items[this.m_lvEx.Items[x].Index].SubItems[y], false);
                }
            }
        }

        private void panel1_Resize(object sender, EventArgs e)
        {
            panel1_Resize();
        }
        public void panel1_Resize()
        {
            try
            {
                this.pnlFileSizeMonitor.Top = this.panel1.Height - this.pnlFileSizeMonitor.Height - 2;
               
                this.btnChkAll.Top = pnlFileSizeMonitor.Top - this.btnChkAll.Height - 2;
                this.btnUncheckAll.Top = this.btnChkAll.Top;
                this.cmbFilter.Top = btnChkAll.Top;
                this.btnRunOC7.Top = this.btnChkAll.Top;
                this.btnRun.Top = this.btnChkAll.Top;
                this.lblMsg.Top = this.btnRun.Top - this.lblMsg.Height - 5;
                this.m_lvEx.Height = this.lblMsg.Top - this.m_lvEx.Top + 10;
                this.m_lvEx.Width = this.panel1.Width - (m_lvEx.Left * 2);
                this.btnRun.Left = this.m_lvEx.Width - (int)(m_lvEx.Width * .5) - (int)(btnRun.Width * .5);
                this.btnRunOC7.Left = (int) this.btnRun.Left + 150;
                this.lblMsg.Width = this.m_lvEx.Width;

                if (uc_filesize_monitor1.lblMaxSize.Left + uc_filesize_monitor1.lblMaxSize.Width > uc_filesize_monitor1.Width)
                {
                    for (; ; )
                    {
                        uc_filesize_monitor1.Width = uc_filesize_monitor1.Width + 1;
                        if (uc_filesize_monitor1.lblMaxSize.Left + uc_filesize_monitor1.lblMaxSize.Width < uc_filesize_monitor1.Width)
                            break;

                    }
                }
                if (uc_filesize_monitor2.lblMaxSize.Left + uc_filesize_monitor2.lblMaxSize.Width > uc_filesize_monitor2.Width)
                {
                    for (; ; )
                    {
                        uc_filesize_monitor2.Width = uc_filesize_monitor2.Width + 1;
                        if (uc_filesize_monitor2.lblMaxSize.Left + uc_filesize_monitor2.lblMaxSize.Width < uc_filesize_monitor2.Width)
                            break;

                    }
                }
                if (uc_filesize_monitor3.lblMaxSize.Left + uc_filesize_monitor3.lblMaxSize.Width > uc_filesize_monitor3.Width)
                {
                    for (; ; )
                    {
                        uc_filesize_monitor3.Width = uc_filesize_monitor3.Width + 1;
                        if (uc_filesize_monitor3.lblMaxSize.Left + uc_filesize_monitor3.lblMaxSize.Width < uc_filesize_monitor3.Width)
                            break;

                    }
                }






                this.uc_filesize_monitor2.Left = this.uc_filesize_monitor1.Left + uc_filesize_monitor2.Width + 2;
                this.uc_filesize_monitor3.Left = this.uc_filesize_monitor2.Left + uc_filesize_monitor3.Width + 2;
            }
            catch
            {
            }
        }

        private void btnChkAll_Click(object sender, EventArgs e)
        {
            for (int x = 0; x <= this.m_lvEx.Items.Count - 1; x++)
            {
                if (cmbFilter.Text.Trim().ToUpper()=="ALL" ||
                    cmbFilter.Text.Trim().ToUpper()== m_lvEx.Items[x].SubItems[COL_VARIANT].Text.Trim().ToUpper())
                        m_lvEx.Items[x].Checked = true;
            }
        }

        private void btnUncheckAll_Click(object sender, EventArgs e)
        {
            for (int x = 0; x <= this.m_lvEx.Items.Count - 1; x++)
            {
                if (cmbFilter.Text.Trim().ToUpper() == "ALL" ||
                    cmbFilter.Text.Trim().ToUpper() == m_lvEx.Items[x].SubItems[COL_VARIANT].Text.Trim().ToUpper())
                        m_lvEx.Items[x].Checked = false;
            }
        }

        private void btnRunOC7_Click(object sender, EventArgs e)
        {

            btnRun_Click(sender, e);
        }

        private void RunScenario_Start()
        {
            ReferenceProcessorScenarioForm.tlbScenario.Enabled = false;
            ReferenceProcessorScenarioForm.EnableTabPage(ReferenceProcessorScenarioForm.tabControlScenario, "tbDataSources", false);
            ReferenceProcessorScenarioForm.EnableTabPage(ReferenceProcessorScenarioForm.tabControlRules, "tbHarvestMethod,tbWoodValue,tbEscalators,tbAddHarvestCosts,tbFilterCond,tbMoveInCosts", false);
            string strPath = frmMain.g_oFrmMain.getProjectDirectory() + "\\OPCOST\\Input";
            if (!System.IO.Directory.Exists(strPath))
                System.IO.Directory.CreateDirectory(strPath);
             
            btnChkAll.Enabled = false;
            btnUncheckAll.Enabled = false;

            this.lblMsg.Text = "";
            this.lblMsg.Show();
            this.m_strDateTimeCreated = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt");
           
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(this.RunScenario_Main));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.m_oThread.Start();


        }
        private void RunScenario_Main()
        {
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
           
            m_oExcel = new excel_latebinding.excel_latebinding();
            int x,y,z;
            int intCount = 0;
            int intRowCount = 0;
            //@ToDo: We are only supporting OpCost now; Eventually should delete all FRCS code
            bool bFRCS = false;
            bool bOPCOST = true;
            string strInputPath = "";

            
            dao_data_access oDao = new dao_data_access();

           
            string strRx1, strRx2, strRx3, strRx4, strRxPackage, strVariant;
            

            frmMain.g_oDelegate.CurrentThreadProcessName = "main";
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;

            if (System.IO.File.Exists(m_strDebugFile))
                System.IO.File.Delete(m_strDebugFile);

            System.Threading.Thread.Sleep(2000);

            if (frmMain.g_bDebug)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "*****START*****" + System.DateTime.Now.ToString() + "\r\n");

            m_intLvCheckedCount=0;
            m_intLvTotalCount = this.m_lvEx.Items.Count;
            for (x = 0; x <= this.m_lvEx.Items.Count - 1; x++)
            {
                ReferenceProgressBarEx = (ProgressBarEx.ProgressBarEx)this.m_lvEx.GetEmbeddedControl(COL_RUNSTATUS, x);
                ReferenceProgressBarEx.backgroundpainter.Color = ReferenceProgressBarEx.backgroundpainter.DefaultColor;
                frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", 0);
                if ((bool)frmMain.g_oDelegate.GetListViewExItemPropertyValue(this.m_lvEx, x, "Checked", false))
                {
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Text", "0%");
                    m_intLvCheckedCount++;
                    
                }
                else
                {   
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Text", "");
                    
                }
                frmMain.g_oDelegate.ExecuteControlMethod(ReferenceProgressBarEx, "Refresh");
               
            }
            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Prepare for processing...Stand By");
           
            
            for (x = 0; x <= this.m_lvEx.Items.Count - 1; x++)
            {
                if ((bool)frmMain.g_oDelegate.GetListViewExItemPropertyValue(this.m_lvEx, x, "Checked", false))
                {
                   
                    if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.UserControl)uc_filesize_monitor1, "Visible", false) == false)
                    {
                        uc_filesize_monitor1.BeginMonitoringFile(
                            m_oQueries.m_strTempDbFile, 2000000000, "2GB");
                        uc_filesize_monitor1.Information = "Work table containing table links";
                        uc_filesize_monitor2.BeginMonitoringFile(
                             frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                            "\\processor\\" + ScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsDbFile
                            , 2000000000, "2GB");
                        uc_filesize_monitor2.Information = "Scenario results DB file containing Harvest Costs and Tree Volume and Value tables";
                    }
                    frmMain.g_oDelegate.EnsureListViewExItemVisible(this.m_lvEx, x);

                    frmMain.g_oDelegate.SetListViewExItemPropertyValue(this.m_lvEx, x, "Selected", true);
                    frmMain.g_oDelegate.SetListViewExItemPropertyValue(this.m_lvEx, x, "Focused", true);

                    this.m_intError = 0;
                    this.m_strError = "";
                    ReferenceProgressBarEx = (ProgressBarEx.ProgressBarEx)this.m_lvEx.GetEmbeddedControl(COL_RUNSTATUS, x);

                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Maximum", 100);
                        
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Minimum", 0);
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", 0);
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Text", "0%");

                    frmMain.g_oDelegate.SetStatusBarPanelTextValue(
                               frmMain.g_sbpInfo.Parent,
                               1,
                               "Processing " + Convert.ToString(intCount + 1) + " Of " + Convert.ToString(frmMain.g_oDelegate.GetListViewExCheckedItemsCount(m_lvEx, false)) + "...Stand By");

                    strVariant = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_VARIANT, "Text", false);
                    strVariant = strVariant.Trim();

                    //get the package and treatments
                    strRxPackage = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_PACKAGE, "Text", false);
                    strRxPackage = strRxPackage.Trim();

                    strRx1 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_RXCYCLE1, "Text", false);
                    strRx1 = strRx1.Trim();

                    strRx2 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_RXCYCLE2, "Text", false);
                    strRx2 = strRx2.Trim();

                    strRx3 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_RXCYCLE3, "Text", false);
                    strRx3 = strRx3.Trim();

                    strRx4 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_RXCYCLE4, "Text", false);
                    strRx4 = strRx4.Trim();

                    m_strOPCOSTBatchFile=frmMain.g_oEnv.strTempDir + "\\" + 
                        "OPCOST_Input_P" + strRxPackage + "_" + strRx1 + "_" + strRx2 + "_" + strRx3 + "_" + strRx4 + ".BAT";

                    if (System.IO.File.Exists(m_strOPCOSTBatchFile))
                        System.IO.File.Delete(m_strOPCOSTBatchFile);

                    //find the package item in the package collection
                    for (y = 0; y <= this.m_oRxPackageItem_Collection.Count - 1; y++)
                    {
                        if (this.m_oRxPackageItem_Collection.Item(y).SimulationYear1Rx.Trim() == strRx1.Trim() &&
                            this.m_oRxPackageItem_Collection.Item(y).SimulationYear2Rx.Trim() == strRx2.Trim() &&
                            this.m_oRxPackageItem_Collection.Item(y).SimulationYear3Rx.Trim() == strRx3.Trim() &&
                            this.m_oRxPackageItem_Collection.Item(y).SimulationYear4Rx.Trim() == strRx4.Trim() &&
                            this.m_oRxPackageItem_Collection.Item(y).RxPackageId.Trim() == strRxPackage.Trim())
                            break;


                    }
                    if (y <= m_oRxPackageItem_Collection.Count - 1)
                    {
                        this.m_oRxPackageItem = new RxPackageItem();
                        m_oRxPackageItem.CopyProperties(m_oRxPackageItem_Collection.Item(y), m_oRxPackageItem);
                    }
                    else
                    {
                        this.m_oRxPackageItem = null;
                    }

                    //get the list of treatment cycle year fields to reference for this package
                    this.m_strRxCycleList = "";
                    if (strRx1.Trim().Length > 0 && strRx1.Trim() != "000") this.m_strRxCycleList = "1,";
                    if (strRx2.Trim().Length > 0 && strRx2.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "2,";
                    if (strRx3.Trim().Length > 0 && strRx3.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "3,";
                    if (strRx4.Trim().Length > 0 && strRx4.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "4,";

                    RunScenario_LoadScenarioHarvestMethodVariables();
                    //load diameter max min variables
                    RunScenario_LoadDiameterVariables(strVariant,strRxPackage);

                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "TreeVolValLowSlope") == true)
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE TreeVolValLowSlope");
                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "TreeVolValSteepSlope") == true)
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE TreeVolValSteepSlope");
                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "HarvestCostsWorkTable") == true)
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE HarvestCostsWorkTable");

                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "frcs_input_steepslope") == true)
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE frcs_input_steepslope");

                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "opcost_input_steepslope") == true)
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE opcost_input_steepslope");


                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "frcs_input") == true)
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE frcs_input");

                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "opcost_input") == true)
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE opcost_input");

                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "opcost_output") == true)
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE opcost_output");

                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "opcost_err") == true)
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE opcost_err");

                    

                    //Here we set the maximum number of ticks on the progress bar
                    //y cannot exceed theis max number
                    if (m_blnLowSlope == true && m_blnSteepSlope == true)
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Maximum", 44);
                    else
                    {
                        if (m_blnLowSlope == true)
                           frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Maximum", 27);
                        else
                           frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Maximum", 28);

                    }
                    //
                    //Init EXCEL FRCS
                    //
                    m_oExcel.ExcelFileName=frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\FRCS";
                    m_oExcel.ExcelFileName=frmMain.g_oUtils.getRandomFile(m_oExcel.ExcelFileName, "xls");
                    string strSourceFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\PROCESSOR\\DB\\FRCS.XLS";
                    System.IO.File.Copy(strSourceFile, m_oExcel.ExcelFileName.Trim());
                    //
                    //Delete Variant + RXPackage records
                    //
                    
                    y = 0;
                    //
                    //LOW SLOPE
                    //
                    if (m_blnLowSlope == true)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Create Low Slope Diameters Table...Stand By");
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                        RunScenario_CreateDiametersTable(Tables.ProcessorScenarioRun.DefaultDiametersLowSlopeTableName,false);

                       
                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Assign Low Slope Diameter Groups...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_AssignDiameterGroups(Tables.ProcessorScenarioRun.DefaultDiametersLowSlopeTableName);
                        }

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Set Low Slope Diameter Flags...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_SetDiameterFlags(Tables.ProcessorScenarioRun.DefaultDiametersLowSlopeTableName, false);
                        }

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append Low Slope Tree Data...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_AppendTreeData(Tables.ProcessorScenarioRun.DefaultTreeDataInLowSlopeTableName,
                                                       Tables.ProcessorScenarioRun.DefaultDiametersLowSlopeTableName,
                                                       strVariant, strRxPackage, false);
                        }

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Route Low Slope Tree Data To BINS...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_InitBinsTables(Tables.ProcessorScenarioRun.DefaultTreeDataInLowSlopeTableName,
                                                       Tables.ProcessorScenarioRun.DefaultTreeBinLowSlopeTableName,
                                                       Tables.ProcessorScenarioRun.DefaultTreeHwdBinLowSlopeTableName);
                        }

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Sum Low Slope BINS by Plot/RxPackage/Rx/RxCyle/Species Group/Diameter Group...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_SumBinsByPlotRxSpcGrpDbhGrp(Tables.ProcessorScenarioRun.DefaultTreeBinLowSlopeTableName,
                                                                    Tables.ProcessorScenarioRun.DefaultTreeHwdBinLowSlopeTableName,
                                                                    "bin_sums_lowslope", "hwd_bin_sums_lowslope");
                        }

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Sum Low Slope BINS by Plot/RxPackage/Rx/RxCyle...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_SumBinsByPlotRx("bin_sums_lowslope", "hwd_bin_sums_lowslope",
                                                        "bin_totals_lowslope", "hwd_bin_totals_lowslope");
                        }

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Create and Append To Low Slope Output Table...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_InitOutputTable("bin_sums_lowslope", "bin_output_lowslope");
                        }
                        if (m_intError == 0)
                        {
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_UpdateOutputTable("bin_output_lowslope",
                                                          "bin_totals_lowslope",
                                                          "hwd_bin_totals_lowslope",
                                                          Tables.ProcessorScenarioRun.DefaultTreeDataInLowSlopeTableName,
                                                          Tables.ProcessorScenarioRun.DefaultDiametersLowSlopeTableName,
                                                          "sumBiomassByLogSizeLowSlope",
                                                          "sumBiomassByStandRxLowSlope",
                                                          strVariant,
                                                          strRxPackage);
                        }
                        if (m_intError == 0)
                        {
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_UpdateOutputTableWithHarvestingSystem("bin_output_lowslope", false);
                        }

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append To Low Slope Tree Volumes And Values Work Table...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_InitTreeVolVal("bin_sums_lowslope",
                                                       "hwd_bin_sums_lowslope",
                                                       strVariant,
                                                       "< " + ScenarioHarvestMethodVariables.SteepSlope.ToString().Trim()
                                                       );
                        }

                        if (m_intError == 0)
                        {
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_AppendToTreeVolVal("TreeVolValLowSlope");
                        }


                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Update Low Slope Tree Volumes And Values Table With Merch and Chip Market Values...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_UpdateTreeVolValWithMerchChipMarketValues("TreeVolValLowSlope");
                        }

                        if (m_intError == 0)
                        {
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            if (bFRCS)
                            {
                                frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append Low Slope Data To FRCS Input Table...Stand By");
                                RunScenario_InitFRCSInputTable("bin_output_lowslope", "frcs_input");
                            }
                            if (bOPCOST)
                  
                            {
                                frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append Low Slope Data To OPCOST Input Table...Stand By");
                                RunScenario_InitOPCOSTInputTable("bin_output_lowslope", "opcost_input");
                            }
                           
                        }
                      


                        if (m_intError == 0)
                        {
                            if (bFRCS)
                            m_oAdo.m_strSQL = "SELECT COUNT(*) AS reccount FROM frcs_input";
                            else
                            m_oAdo.m_strSQL = "SELECT COUNT(*) AS reccount FROM opcost_input";

                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            ScenarioHarvestMethodVariables.LowSlopeRecordCount =
                                (int)m_oAdo.getSingleDoubleValueFromSQLQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, "frcs_input");

                        }

                        
                    }
                    //
                    //STEEP SLOPE
                    //
                    if (m_blnSteepSlope == true)
                    {
                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Create Steep Slope Diameters Table...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_CreateDiametersTable(Tables.ProcessorScenarioRun.DefaultDiametersSteepSlopeTableName, true);
                        }

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Assign Steep Slope Diameter Groups...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_AssignDiameterGroups(Tables.ProcessorScenarioRun.DefaultDiametersSteepSlopeTableName);
                        }

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Set Steep Slope Diameter Flags...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_SetDiameterFlags(Tables.ProcessorScenarioRun.DefaultDiametersSteepSlopeTableName, true);
                        }

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append Steep Slope Tree Data...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_AppendTreeData(Tables.ProcessorScenarioRun.DefaultTreeDataInSteepSlopeTableName,
                                                       Tables.ProcessorScenarioRun.DefaultDiametersSteepSlopeTableName,
                                                       strVariant, strRxPackage, true);
                        }

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Route Steep Slope Tree Data To BINS...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_InitBinsTables(Tables.ProcessorScenarioRun.DefaultTreeDataInSteepSlopeTableName,
                                                       Tables.ProcessorScenarioRun.DefaultTreeBinSteepSlopeTableName,
                                                       Tables.ProcessorScenarioRun.DefaultTreeHwdBinSteepSlopeTableName);
                        }
                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Sum Steep Slope BINS by Plot/RxPackage/Rx/RxCyle/Species Group/Diameter Group...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_SumBinsByPlotRxSpcGrpDbhGrp(Tables.ProcessorScenarioRun.DefaultTreeBinSteepSlopeTableName,
                                                                    Tables.ProcessorScenarioRun.DefaultTreeHwdBinSteepSlopeTableName,
                                                                    "bin_sums_steepslope", "hwd_bin_sums_steepslope");
                        }
                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Sum Steep Slope BINS by Plot/RxPackage/Rx/RxCyle...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_SumBinsByPlotRx("bin_sums_steepslope", "hwd_bin_sums_steepslope",
                                                        "bin_totals_steepslope", "hwd_bin_totals_steepslope");
                        }
                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Create and Append To Steep Slope Output Table...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_InitOutputTable("bin_sums_steepslope", "bin_output_steepslope");
                        }
                        if (m_intError == 0)
                        {
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_UpdateOutputTable("bin_output_steepslope",
                                                          "bin_totals_steepslope",
                                                          "hwd_bin_totals_steepslope",
                                                          Tables.ProcessorScenarioRun.DefaultTreeDataInSteepSlopeTableName,
                                                          Tables.ProcessorScenarioRun.DefaultDiametersSteepSlopeTableName,
                                                          "sumBiomassByLogSizeSteepSlope",
                                                          "sumBiomassByStandRxSteepSlope",
                                                          strVariant,
                                                          strRxPackage);
                        }
                        if (m_intError == 0)
                        {
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_UpdateOutputTableWithHarvestingSystem("bin_output_steepslope", true);
                        }
                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append To Steep Slope Tree Volumes And Values Work Table...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_InitTreeVolVal("bin_sums_steepslope",
                                                       "hwd_bin_sums_steepslope",
                                                       strVariant,
                                                       ">= " + ScenarioHarvestMethodVariables.SteepSlope.ToString().Trim()
                                                       );
                        }
                        if (m_intError == 0)
                        {
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_AppendToTreeVolVal("TreeVolValSteepSlope");
                        }
                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Update Steep Slope Tree Volumes And Values Work Table With Merch and Chip Market Values...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            RunScenario_UpdateTreeVolValWithMerchChipMarketValues("TreeVolValSteepSlope");
                        }
                        if (m_intError == 0)
                        {
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                        }
                        if (m_intError == 0 && bFRCS)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append Steep Slope Data To FRCS Input Table...Stand By");
                            RunScenario_InitFRCSInputTable("bin_output_steepslope", "frcs_input_steepslope");
                        }
                         if (m_intError == 0 && bOPCOST)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append Steep Slope Data To OPCOST Input Table...Stand By");
                            RunScenario_InitOPCOSTInputTable("bin_output_steepslope", "opcost_input_steepslope");
                        }
                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append Steep Slope Data To FRCS Input Table...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            if (bFRCS)
                                m_oAdo.m_strSQL = "SELECT COUNT(*) AS reccount FROM frcs_input_steepslope";
                            else
                                m_oAdo.m_strSQL = "SELECT COUNT(*) AS reccount FROM opcost_input_steepslope";
                            ScenarioHarvestMethodVariables.SteepSlopeRecordCount =
                                (int)m_oAdo.getSingleDoubleValueFromSQLQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, "frcs_input_steepslope");
                        }
                        if (m_intError == 0)
                        {
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                        }
                        if (m_intError == 0 && bFRCS)
                        {
                            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "frcs_input") == true)
                                RunScenario_AppendToFRCSInputTable("frcs_input", "frcs_input_steepslope");
                            else
                                RunScenario_InitFRCSInputTable("frcs_input", "frcs_input_steepslope");
                        }
                        if (m_intError == 0 && bOPCOST)
                        {
                            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "opcost_input") == true)
                                RunScenario_AppendToOPCOSTInputTable("opcost_input", "opcost_input_steepslope");
                            else
                            {
                                RunScenario_InitOPCOSTInputTable("bin_output_steepslope","opcost_input");
                               
                            }
                        }



                    }


                    intCount++;

                    if (m_intError == 0)
                    {
                        y++;
                    }
                    if (m_intError == 0 && bFRCS)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Update FRCS Spreadsheet With Threshold Values...Stand By");
                        RunScenario_UpdateFRCSThresholds();
                    }
                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }
                    if (m_intError == 0 && bFRCS)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "FRCS Processing Batch Input...Stand By");
                        RunScenario_ProcessFRCS();
                    }
                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }
                    if (m_intError == 0 && bFRCS)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Retrieve FRCS Fatal Errors And Warnings...Stand By");
                        RunScenario_FRCSWarningMessages();
                    }
                     if (m_intError == 0 && bOPCOST)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "OPCOST Processing Batch Input...Stand By");
                        RunScenario_ProcessOPCOST(strVariant,strRxPackage);
                    }
                    if (m_intError == 0)
                     {
                         y++;
                         frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                     }
                    //if (m_intError == 0 && bFRCS)
                    //{
                    //    frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append FRCS Data To Harvest Costs Work Table...Stand By");
                    //    RunScenario_AppendToHarvestCosts("HarvestCostsWorkTable",true,false);
                    //}
                    //if (m_intError == 0 && bOPCOST)
                    //{
                    //    frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append OPCOST Data To Harvest Costs Work Table...Stand By");
                    //    RunScenario_AppendToHarvestCosts("HarvestCostsWorkTable", false,false);
                    //}
                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }

                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Update Harvest Costs Work Table With Additional Costs...Stand By");
                        RunScenario_UpdateHarvestCostsTableWithAdditionalCosts("HarvestCostsWorkTable","");
                    }
                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }

                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Delete Old Variant=" + strVariant + " and RxPackage=" + strRxPackage + " Records From Harvest Costs And Tree Vol Val Table...Stand By");
                        RunScenario_DeleteFromTreeVolValAndHarvestCostsTable(strVariant, strRxPackage);
                    }
                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }

                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append New Variant=" + strVariant + " and RxPackage=" + strRxPackage + " Records To Harvest Costs And Tree Vol Val Table...Stand By");
                        RunScenario_AppendToTreeVolValAndHarvestCostsTable();
                    }

                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }
        

                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append Placeholder Records For Variant=" + strVariant + " and RxPackage=" + strRxPackage + " To Tree Vol Val And Harvest Cost Tables...Stand By");
                        RunScenario_AppendPlaceholdersToTreeVolValAndHarvestCostsTables();
                    }

                    
                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Finalizing Processor Scenario Database Tables...Stand By");
                        //update counts
                        intRowCount = 0;
                        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "TreeVolValLowSlope"))
                            intRowCount = (int)m_oAdo.getRecordCount(m_oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM TreeVolValLowSlope", "temp");

                        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "TreeVolValSteepSlope"))
                            intRowCount = intRowCount + (int)m_oAdo.getRecordCount(m_oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM TreeVolValSteepSlope", "temp");

                        frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_VOLVAL, intRowCount.ToString());

                        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "HarvestCostsWorkTable"))
                            intRowCount = (int)m_oAdo.getRecordCount(m_oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM HarvestCostsWorkTable", "temp");
                        frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_HVSTCOST, intRowCount.ToString());
                    }

                    if (System.IO.File.Exists(m_strOPCOSTBatchFile))
                        System.IO.File.Delete(m_strOPCOSTBatchFile);

                    //compact mdb
                    if (m_intError == 0)
                    {

                        string strConn = m_oAdo.m_OleDbConnection.ConnectionString;
                        string strDb = m_oQueries.m_strTempDbFile;
                        m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
                        System.Threading.Thread.Sleep(5000);
                       
                        //check if file size greater than 70% of 2GB
                        if (uc_filesize_monitor1.CurrentPercent( m_oQueries.m_strTempDbFile, 2000000000) > 70)
                        {
                            oDao.m_DaoDbEngine.Idle(1);
                            oDao.m_DaoDbEngine.Idle(8);
                            oDao.DisplayErrors = false;
                            oDao.m_intErrorCount = 0;
                            oDao.m_intError = -1;
                            for (; ; )
                            {
                                oDao.CompactMDB(strDb);
                                if (oDao.m_intError == 0)
                                {
                                   break;
                                }
                                if (oDao.m_intErrorCount ==  5) break;
                                if (oDao.m_intErrorCount == 4)
                                {
                                    int count = oDao.m_intErrorCount;
                                    oDao.m_DaoDbEngine.Idle(1);
                                    oDao.m_DaoDbEngine.Idle(8);
                                    oDao.m_DaoWorkspace.Close();
                                    oDao.m_DaoDbEngine = null;
                                    oDao = null;
                                    oDao = new dao_data_access();
                                    oDao.m_intErrorCount = count;

                                }
                                
                                oDao.m_intErrorCount++;
                                System.Threading.Thread.Sleep(5000);


                            }
                            System.Threading.Thread.Sleep(5000);
                        }
                        oDao.DisplayErrors = true;
                        if (oDao.m_intError != 0)
                        {
                            MessageBox.Show("Failed to compact and repair file " + m_oQueries.m_strTempDbFile, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        string strInputFile = "";
                        if (bFRCS)
                        {
                           strInputPath = frmMain.g_oFrmMain.getProjectDirectory() + "\\FRCS\\Input";
                           strInputFile = "FRCS_Input_P" + strRxPackage + "_" + strRx1 + "_" + strRx2 + "_" + strRx3 + "_" + strRx4 + "_" + m_strDateTimeCreated + ".accdb";
                           strInputFile = strInputFile.Replace(":", "_");
                           strInputFile = strInputFile.Replace(" ", "_");
                           System.Threading.Thread.Sleep(5000);
                           System.IO.File.Copy(m_oQueries.m_strTempDbFile, strInputPath + "\\" + strInputFile, true);
                           //delete the work tables and any links
                           m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strInputPath + "\\" + strInputFile, "", ""),5);
                           if (m_oAdo.m_intError == 0)
                           {
                               string[] strTables = m_oAdo.getTableNames(m_oAdo.m_OleDbConnection);
                               if (strTables != null)
                               {
                                   for (z = 0; z <= strTables.Length - 1; z++)
                                   {
                                       if (strTables[z] != null)
                                       {
                                           switch (strTables[z].Trim().ToUpper())
                                           {
                                               case "FRCS_INPUT": break;
                                               case "FRCSVARIABLESLOWSLOPETABLE": break;
                                               case "FRCSVARIABLESSTEEPSLOPETABLE": break;
                                               case "FRCS_FATAL_ERROR_MESSAGES": break;
                                               case "FRCS_OUTPUT": break;
                                               case "FRCS_WARNING_MESSAGES": break;
                                               default:
                                                   m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + strTables[z].Trim());
                                                   break;

                                           }
                                       }
                                   }
                               }
                               m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
                               System.Threading.Thread.Sleep(5000);
                               if (uc_filesize_monitor1.CurrentPercent(strInputPath + "\\" + strInputFile, 2000000000) > 70)
                               {
                                   oDao.m_DaoDbEngine.Idle(1);
                                   oDao.m_DaoDbEngine.Idle(8);
                                   oDao.CompactMDB(strInputPath + "\\" + strInputFile);
                                   System.Threading.Thread.Sleep(5000);
                               }

                           }
                           
                        }
                        if (bOPCOST)
                        {
                            strInputPath = frmMain.g_oFrmMain.getProjectDirectory() + "\\OPCOST\\Input";
                            strInputFile = System.IO.Path.GetFileNameWithoutExtension(frmMain.g_strOPCOSTDirectory) + "_Input_" +
                                           strVariant + "_P" + strRxPackage + "_" + strRx1 + "_" + strRx2 + "_" + strRx3 + "_" + strRx4 + "_" + m_strDateTimeCreated + ".accdb";
                            strInputFile = strInputFile.Replace(":", "_");
                            strInputFile = strInputFile.Replace(" ", "_");
                            System.IO.File.Copy(m_oQueries.m_strTempDbFile, strInputPath + "\\" + strInputFile, true);
                            System.Threading.Thread.Sleep(5000);
                            //delete the work tables and any links
                            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strInputPath + "\\" + strInputFile, "", ""),5);
                            if (m_oAdo.m_intError == 0)
                            {
                                string[] strTables = m_oAdo.getTableNames(m_oAdo.m_OleDbConnection);
                                if (strTables != null)
                                {
                                    for (z = 0; z <= strTables.Length - 1; z++)
                                    {
                                        if (strTables[z] != null)
                                        {
                                            switch (strTables[z].Trim().ToUpper())
                                            {
                                                case "OPCOST_ERRORS": break;
                                                //case "OPCOST_IDEAL_ERRORS": break;
                                                case "OPCOST_INPUT": break;
                                                case "OPCOST_OUTPUT": break;
                                                //case "OPCOST_IDEAL_OUTPUT": break;
                                                default:
                                                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + strTables[z].Trim());
                                                    break;

                                            }
                                        }
                                    }
                                }
                                
                                m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
                                System.Threading.Thread.Sleep(5000);
                                if (uc_filesize_monitor1.CurrentPercent(strInputPath + "\\" + strInputFile, 2000000000) > 70)
                                {
                                    oDao.m_DaoDbEngine.Idle(1);
                                    oDao.m_DaoDbEngine.Idle(8);
                                    oDao.CompactMDB(strInputPath + "\\" + strInputFile);
                                    System.Threading.Thread.Sleep(5000);
                                }

                            }
                        }
                        m_intError = oDao.m_intError;
                        m_oAdo.OpenConnection(strConn,5);
                    }
                   
                    





                   


                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                        frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_CHECKBOX, " ");
                        frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_PROCESSOR_PROCESSINGDATETIME, m_strDateTimeCreated);
                    }
                    else
                    {
                        ReferenceProgressBarEx.backgroundpainter.Color = Color.Red;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", 0);
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Text", "!!Error!!");
                    }

                    System.Threading.Thread.Sleep(2000);
                    


                }

                
                
            }

            oDao.m_DaoDbEngine.Idle(1);
            oDao.m_DaoDbEngine.Idle(8);
            oDao.m_DaoWorkspace.Close();
            oDao.m_DaoDbEngine = null;
            oDao = null;

            MessageBox.Show("Done","FIA Biosum");
            if (frmMain.g_bDebug)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "*****END*****" + System.DateTime.Now.ToString() + "\r\n");
           	
            RunScenario_Finished();
            
			
            frmMain.g_oDelegate.CurrentThreadProcessDone = true;
            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        }
        private void RunScenario_DeleteAll()
        {
            m_oAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName;
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL);

            m_oAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName;
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL);
        }

        private void RunScenario_LoadScenarioHarvestMethodVariables()
        {
            int x;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel>1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_LoadScenarioHarvestMethodVariables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            m_oAdo.m_strSQL = "SELECT userxdefaultharvestmethodyn FROM scenario_harvest_method WHERE TRIM(UCASE(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'";
            if ((string)m_oAdo.getSingleStringValueFromSQLQuery(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL,"temp").Trim()=="Y")
            {
                ScenarioHarvestMethodVariables.UseDefaultRxHarvestMethod=true;
            }
            else
            {
                ScenarioHarvestMethodVariables.UseDefaultRxHarvestMethod=false;
            }

            ScenarioHarvestMethodVariables.HarvestMethodLowSlope = (string)m_oAdo.getSingleStringValueFromSQLQuery(m_oAdo.m_OleDbConnection,"SELECT harvestmethodlowslope FROM scenario_harvest_method WHERE TRIM(UCASE(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'","temp"); 
            ScenarioHarvestMethodVariables.HarvestMethodSteepSlope = (string)m_oAdo.getSingleStringValueFromSQLQuery(m_oAdo.m_OleDbConnection,"SELECT harvestmethodsteepslope FROM scenario_harvest_method WHERE TRIM(UCASE(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'","temp"); 
            ScenarioHarvestMethodVariables.SteepSlope = (int)m_oAdo.getSingleDoubleValueFromSQLQuery(m_oAdo.m_OleDbConnection,"SELECT steepslope FROM scenario_harvest_method WHERE TRIM(UCASE(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'","temp");                
        }
        private void RunScenario_LoadDiameterVariables(string p_strVariant,string p_strRxPackage)
        {
            int x;
            string strTableName;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_LoadDiameterVariables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            strTableName = "fvs_tree_IN_" + p_strVariant + "_P" + p_strRxPackage + "_TREE_CUTLIST";

            m_oAdo.m_strSQL = "SELECT MAX(DBH) AS maxdbh FROM " + strTableName;
            DiameterVariables.maxdia = (double)m_oAdo.getSingleDoubleValueFromSQLQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL,strTableName);

            m_oAdo.m_strSQL = "SELECT min_chip_dbh FROM scenario_harvest_method WHERE TRIM(UCASE(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'";
            DiameterVariables.MinDiaChips = (double)m_oAdo.getSingleDoubleValueFromSQLQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, strTableName);

            m_oAdo.m_strSQL = "SELECT min_sm_log_dbh FROM scenario_harvest_method WHERE TRIM(UCASE(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'";
            DiameterVariables.MinDiaSmLogs = (double)m_oAdo.getSingleDoubleValueFromSQLQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, strTableName);
                
            DiameterVariables.MaxDiaChips = (double)(DiameterVariables.MinDiaSmLogs - (double)0.1);

            m_oAdo.m_strSQL = "SELECT min_lg_log_dbh FROM scenario_harvest_method WHERE TRIM(UCASE(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'";
            DiameterVariables.MinDiaLgLogs = (double)m_oAdo.getSingleDoubleValueFromSQLQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, strTableName);

            DiameterVariables.MaxDiaSmLogs = (double)(DiameterVariables.MinDiaLgLogs - (double)0.1);

            DiameterVariables.MaxDiaLgLogs = DiameterVariables.maxdia;

            m_oAdo.m_strSQL = "SELECT min_dbh_steep_slope FROM scenario_harvest_method WHERE TRIM(UCASE(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'";
            DiameterVariables.MinDiaSteepSlope = (double)m_oAdo.getSingleDoubleValueFromSQLQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, strTableName);

            DiameterVariables.DiaClass_BC = "1-" + Convert.ToString(DiameterVariables.MinDiaChips - 0.1).Trim();

            DiameterVariables.DiaClass_CHIPS = Convert.ToString(DiameterVariables.MinDiaChips).Trim() + "-" + Convert.ToString(DiameterVariables.MaxDiaChips).Trim();

            DiameterVariables.DiaClass_SMLOGS = Convert.ToString(DiameterVariables.MaxDiaSmLogs).Trim() + "-" + Convert.ToString(DiameterVariables.MaxDiaSmLogs).Trim();

            DiameterVariables.DiaClass_LGLOGS = Convert.ToString(DiameterVariables.MaxDiaLgLogs).Trim() + "+";

            DiameterVariables.DiaClass_AllLOGS = Convert.ToString(DiameterVariables.MinDiaSteepSlope).Trim() + "+";

            DiameterVariables.DiaCount = DiameterVariables.maxdia * 10;


        }
        private void RunScenario_CreateDiametersTable(string p_strDiametersTableName,bool p_bSteepSlope)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_CreateDiametersTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            int x;
            int y;
            string strTempTable = "";
            
            double diameter;
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection,p_strDiametersTableName.Trim())==true)
                  m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection,"DROP TABLE " + p_strDiametersTableName);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2 && m_oAdo.m_intError==0)
                frmMain.g_oUtils.WriteText(m_strDebugFile, Tables.ProcessorScenarioRun.CreateDiametersTableSQL(p_strDiametersTableName) + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            if (m_oAdo.m_intError==0) 
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.CreateDiametersTableSQL(p_strDiametersTableName));
            //This query should include the scenario if it's ever used again
            m_oAdo.m_strSQL = "SELECT species_group " +
                              "FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName + " " +
                              "ORDER BY species_group";

            if (m_oAdo.m_intError == 0)
            {
                string[] strSpcGrpArray = frmMain.g_oUtils.ConvertListToArray(m_oAdo.CreateCommaDelimitedList(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, ""), ",");
                if (m_oAdo.m_intError == 0)
                {
                    for (y = 0; y <= strSpcGrpArray.Length - 1; y++)
                    {
                        strTempTable = p_strDiametersTableName.Trim() + strSpcGrpArray[y].Trim();
                        diameter = 0.9;

                        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, strTempTable) == true)
                            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + strTempTable);
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, Tables.ProcessorScenarioRun.CreateDiametersTableSQL(strTempTable) + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.CreateDiametersTableSQL(strTempTable));
                        if (m_oAdo.m_intError == 0)
                        {
                            for (x = 1; x <= DiameterVariables.DiaCount; x++)
                            {
                                diameter = diameter + 0.1;
                                m_oAdo.m_strSQL = "INSERT INTO " + strTempTable + " " +
                                                  "(DBH) SELECT ROUND(" + diameter + ",1) AS dbh";
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                                if (m_oAdo.m_intError != 0) break;
                            }
                        }
                        if (m_oAdo.m_intError == 0)
                        {
                            if (p_bSteepSlope == false)
                            {
                                m_oAdo.m_strSQL = "UPDATE " + strTempTable + " d " +
                                                      "SET d.class = IIF(d.DBH >= " + DiameterVariables.MinDiaLgLogs.ToString() + ",'LGLOGS'," +
                                                                    "IIF(d.DBH >= " + DiameterVariables.MinDiaSmLogs.ToString() + " AND " +
                                                                        "d.DBH < " + DiameterVariables.MinDiaLgLogs.ToString() + ",'SMLOGS'," +
                                                                    "IIF(d.DBH >= " + DiameterVariables.MinDiaChips.ToString() + " AND " +
                                                                        "d.DBH < " + DiameterVariables.MinDiaSmLogs.ToString() + ",'CHIPS'," +
                                                                    "IIF(d.DBH < " + DiameterVariables.MinDiaChips.ToString() + ",'BC',null))))";
                            }
                            else
                            {
                                m_oAdo.m_strSQL = "UPDATE " + strTempTable + " d " +
                                                  "SET d.class = IIF(d.DBH >= " + DiameterVariables.MinDiaLgLogs.ToString() + " AND " +
                                                                    "d.DBH >= " + DiameterVariables.MinDiaSteepSlope.ToString() + ",'LGLOGS'," +
                                                                "IIF(d.DBH >= " + DiameterVariables.MinDiaSmLogs.ToString() + " AND " +
                                                                    "d.DBH >= " + DiameterVariables.MinDiaSteepSlope.ToString() + " AND " +
                                                                    "d.DBH < " + DiameterVariables.MinDiaLgLogs.ToString() + ",'SMLOGS'," +
                                                                "IIF(d.DBH >= " + DiameterVariables.MinDiaChips.ToString() + " AND " +
                                                                    "d.DBH >= " + DiameterVariables.MinDiaSteepSlope.ToString() + " AND " +
                                                                    "d.DBH < " + DiameterVariables.MinDiaSmLogs.ToString() + ",'CHIPS'," +
                                                                "IIF(d.DBH<" + DiameterVariables.MinDiaChips.ToString() + " " +
                                                                  ",'BC',Null))))";

                            }
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                        }
                        if (m_oAdo.m_intError == 0)
                        {
                            m_oAdo.m_strSQL = "UPDATE " + strTempTable + " " +
                                         "SET  Util_Logs_Bk = 'N', Util_Chips_Bk = 'N'," +
                                               "Res_Stmp_Bk = 'N', Res_Lnd_Bk = 'N', Util_Logs_Br = 'N'," +
                                               "Util_Chips_Br = 'N', Res_Stmp_Br = 'N', Res_Lnd_Br = 'N'," +
                                               "Util_Logs_Bl = 'N', Util_Chips_Bl = 'N', Res_Stmp_Bl = 'N'," +
                                               "Res_Lnd_Bl = 'N'," + 
                                               "NonUtil_Logs_Bl='N',NonUtil_Chips_Bl='N'," + 
                                               "Species_Group=" + strSpcGrpArray[y].Trim();
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

                            m_oAdo.m_strSQL = "INSERT INTO " + p_strDiametersTableName + " " +
                                              "SELECT * FROM " + strTempTable;
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

                            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + strTempTable);
                        }
                        if (m_oAdo.m_intError != 0) break;
                    }
                }
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;

        }
        private void RunScenario_AssignDiameterGroups(string p_strDiametersTableName)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_AssignDiameterGroups\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            int x;
           
            string strDiamGroupList = "";
            string strMinDiaList = "";
            string strSqlList = "";
            double dblDBH;

            m_oAdo.m_strSQL = "SELECT diam_group, min_diam FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName +
                " WHERE TRIM(scenario_id)='" + this.ScenarioId.Trim() + "' ";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            if (m_oAdo.m_intError == 0)
            {
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        strDiamGroupList = strDiamGroupList + m_oAdo.m_OleDbDataReader["diam_group"].ToString().Trim() + ",";
                        strMinDiaList = strMinDiaList + m_oAdo.m_OleDbDataReader["min_diam"].ToString().Trim() + ",";

                    }
                    m_oAdo.m_OleDbDataReader.Close();

                    strDiamGroupList = strDiamGroupList.Substring(0, strDiamGroupList.Length - 1);
                    strMinDiaList = strMinDiaList.Substring(0, strMinDiaList.Length - 1);

                    string[] strDiamGroupArray = frmMain.g_oUtils.ConvertListToArray(strDiamGroupList, ",");
                    string[] strMinDiaArray = frmMain.g_oUtils.ConvertListToArray(strMinDiaList, ",");

                    m_oAdo.m_strSQL = "SELECT DBH FROM " + p_strDiametersTableName;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    if (m_oAdo.m_intError == 0)
                    {
                        if (m_oAdo.m_OleDbDataReader.HasRows)
                        {
                            while (m_oAdo.m_OleDbDataReader.Read())
                            {
                                dblDBH = System.Math.Round(Convert.ToDouble(m_oAdo.m_OleDbDataReader["dbh"]), 1);
                                for (x = 0; x <= strDiamGroupArray.Length - 1; x++)
                                {
                                    if (x < strDiamGroupArray.Length - 1)
                                    {
                                        if (dblDBH >=
                                            System.Math.Round(Convert.ToDouble(strMinDiaArray[x]), 1) &&
                                            dblDBH < System.Math.Round(Convert.ToDouble(strMinDiaArray[x + 1]), 1))
                                        {
                                            strSqlList = strSqlList + "UPDATE " + p_strDiametersTableName + " " +
                                                                      "SET diam_group = " + strDiamGroupArray[x] + " " +
                                                                      "WHERE DBH=" + Convert.ToString(dblDBH) + ",";
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        if (dblDBH >= Convert.ToDouble(strMinDiaArray[x]))
                                        {
                                            strSqlList = strSqlList + "UPDATE " + p_strDiametersTableName + " " +
                                                                      "SET diam_group = " + strDiamGroupArray[x] + " " +
                                                                      "WHERE DBH=" + Convert.ToString(dblDBH) + ",";
                                            break;
                                        }
                                    }

                                }



                            }
                            m_oAdo.m_OleDbDataReader.Close();
                            strSqlList = strSqlList.Substring(0, strSqlList.Length - 1);
                            string[] strSqlArray = frmMain.g_oUtils.ConvertListToArray(strSqlList, ",");
                            for (x = 0; x <= strSqlArray.Length - 1; x++)
                            {
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, strSqlArray[x] + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, strSqlArray[x]);
                                if (m_oAdo.m_intError != 0) break;
                            }

                        }
                        else m_oAdo.m_OleDbDataReader.Close();
                    }
                }
                else m_oAdo.m_OleDbDataReader.Close();
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;


        }
        private void RunScenario_SetDiameterFlags(string p_strDiametersTableName,bool p_bSteepSlope)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_SetDiameterFlags\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            string strWoodBin = "";
            string strSpcGrp = "";
            string strDbhGrp = "";
            int x = 0;


            string[] strSql = new string[(int)m_oAdo.getSingleDoubleValueFromSQLQuery(m_oAdo.m_OleDbConnection, "SELECT COUNT(*) as reccount FROM scenario_tree_species_diam_dollar_values WHERE TRIM(UCASE(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'", "temp") * 4];
            m_oAdo.m_strSQL = "SELECT species_group,diam_group,wood_bin FROM scenario_tree_species_diam_dollar_values WHERE TRIM(UCASE(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

            if (m_oAdo.m_intError == 0)
            {
                while (m_oAdo.m_OleDbDataReader.Read())
                {
                    strWoodBin = Convert.ToString(m_oAdo.m_OleDbDataReader["wood_bin"]).Trim();
                    strSpcGrp = Convert.ToString(m_oAdo.m_OleDbDataReader["species_group"]).Trim();
                    strDbhGrp = Convert.ToString(m_oAdo.m_OleDbDataReader["diam_group"]).Trim();

                    /******************************************************************************
                     **BRUSH CUT: Set the NonUtil_Chips_Bl Y/N flag depending if the diameter is 
                     **below minimum chips diameter
                     ******************************************************************************/
                    strSql[x] = "UPDATE " + p_strDiametersTableName.Trim() + " " +
                                "SET NonUtil_Chips_Bl = IIf(DBH <" + DiameterVariables.MinDiaChips.ToString().Trim() + ",'Y','N') " +
                                "WHERE species_group = " + strSpcGrp + " AND " +
                                      "diam_group = " + strDbhGrp;
                    x++;
                    if (p_bSteepSlope == false)
                    {
                       
                         
                        /******************************************************************************
                         **CHIPS: Set the chips Y/N flag depending if the diameters.dbh value falls 
                         **between the user entered MinChipDBHLT40 and MinLogsDBH value.
                         **Uupdate value to "Y" in diameters table for Chips if DBH within min/max
                         **for chip trees
                         ******************************************************************************/
                        strSql[x] = "UPDATE " + p_strDiametersTableName.Trim() + " " +
                                    "SET Util_Chips_Bl = IIf(DBH=" + DiameterVariables.MinDiaChips.ToString().Trim() + " OR " +
                                                            "DBH > " + DiameterVariables.MinDiaChips.ToString().Trim() + " AND " +
                                                            "DBH < " + DiameterVariables.MinDiaSmLogs.ToString().Trim() + ",'Y','N') " +
                                    "WHERE species_group = " + strSpcGrp + " AND " +
                                          "diam_group = " + strDbhGrp;
                        x++;

                        if (strWoodBin == "M")
                        {


                            /*******************************************************************************
                             **SMALL LOGS: Set the Logs Y/N flag depending if the diameters.dbh value falls
                             **between the user entered MinSmLogsDBH and MinLgLogsDBH
                             **update value to "Y" in diameters table for Logs if DBH within min/max 
                             **for small log trees
                             *******************************************************************************/
                            strSql[x] = "UPDATE " + p_strDiametersTableName.Trim() + " " +
                                        "SET Util_Logs_Bl = IIf(DBH = " + DiameterVariables.MinDiaSmLogs.ToString().Trim() + " OR " +
                                                               "DBH > " + DiameterVariables.MinDiaSmLogs.ToString().Trim() + " AND " +
                                                               "DBH < " + DiameterVariables.MinDiaLgLogs.ToString().Trim() + ",'Y',Util_Logs_Bl) " +
                                        "WHERE species_group = " + strSpcGrp + " AND " +
                                              "diam_group = " + strDbhGrp;
                            x++;

                            /*******************************************************************************
                             **LARGE LOGS: Set the Logs Y/N flag depending if the diameters.dbh value is 
                             **greater or equal to the user entered MinLgLogsDBH 
                             **update value to "Y" in diameters table for Logs if DBH within min/max 
                             **for large log trees
                             *******************************************************************************/
                            strSql[x] = "UPDATE " + p_strDiametersTableName.Trim() + " " +
                                        "SET Util_Logs_Bl = IIf(DBH = " + DiameterVariables.MinDiaLgLogs.ToString().Trim() + " OR " +
                                                               "DBH > " + DiameterVariables.MinDiaLgLogs.ToString().Trim() + ",'Y',Util_Logs_Bl) " +
                                        "WHERE species_group = " + strSpcGrp + " AND " +
                                              "diam_group = " + strDbhGrp;
                            x++;
                        }
                        else
                        {
                            //EVERYTHING GETS PROCESSED AS CHIPS
                            /*******************************************************************************
                             **SMALL LOGS: Set the Chips Y/N flag depending if the diameters.dbh value falls
                             **between the user entered MinSmLogsDBH and MinLgLogsDBH
                             **update value to "Y" in diameters table for Logs if DBH within min/max 
                             **for small log trees
                             *******************************************************************************/
                            strSql[x] = "UPDATE " + p_strDiametersTableName.Trim() + " " +
                                        "SET Util_Chips_Bl = IIf(DBH = " + DiameterVariables.MinDiaSmLogs.ToString().Trim() + " OR " +
                                                                "DBH > " + DiameterVariables.MinDiaSmLogs.ToString().Trim() + " AND " +
                                                                "DBH < " + DiameterVariables.MinDiaLgLogs.ToString().Trim() + ",'Y',Util_Chips_Bl) " +
                                        "WHERE species_group = " + strSpcGrp + " AND " +
                                              "diam_group = " + strDbhGrp;
                            x++;

                            /*******************************************************************************
                             **LARGE LOGS: Set the Chips Y/N flag depending if the diameters.dbh value is 
                             **greater or equal to the user entered MinLgLogsDBH 
                             **update value to "Y" in diameters table for Logs if DBH within min/max 
                             **for large log trees
                             *******************************************************************************/
                            strSql[x] = "UPDATE " + p_strDiametersTableName.Trim() + " " +
                                        "SET Util_Chips_Bl = IIf(DBH = " + DiameterVariables.MinDiaLgLogs.ToString().Trim() + " OR " +
                                                                "DBH > " + DiameterVariables.MinDiaLgLogs.ToString().Trim() + ",'Y',Util_Chips_Bl) " +
                                        "WHERE species_group = " + strSpcGrp + " AND " +
                                              "diam_group = " + strDbhGrp;
                            x++;
                        }
                    }
                    else
                    {

                        /******************************************************************************
                        **CHIPS: Set the chips Y/N flag depending if the diameters.dbh value falls 
                        **between the user entered MinChipDBHLT40 and MinLogsDBH value and the diameters.dbh 
                        **exceeds the minimum diameter for steep slopes 
                        **Update value to "Y" in diameters table for Chips if DBH within min/max
                        **for chip trees
                        ******************************************************************************/
                        strSql[x] = "UPDATE " + p_strDiametersTableName.Trim() + " " +
                                    "SET Util_Chips_Bl = IIf(DBH=" + DiameterVariables.MinDiaChips.ToString().Trim() + " OR " +
                                                            "DBH > " + DiameterVariables.MinDiaChips.ToString().Trim() + " AND " +
                                                            "DBH > " + DiameterVariables.MinDiaSteepSlope.ToString().Trim() + " AND " +
                                                            "DBH < " + DiameterVariables.MinDiaSmLogs.ToString().Trim() + ",'Y','N') " +
                                    "WHERE species_group = " + strSpcGrp + " AND " +
                                          "diam_group = " + strDbhGrp;
                        x++;

                        
                        if (strWoodBin == "M")
                        {
                            /*******************************************************************************
                             **SMALL LOGS: Set the Logs Y/N flag depending if the diameters.dbh value falls
                             **between the user entered MinSmLogsDBH and MinLgLogsDBH
                             **update value to "Y" in diameters table for Logs if DBH within min/max 
                             **for small log trees
                             *******************************************************************************/
                            strSql[x] = "UPDATE " + p_strDiametersTableName.Trim() + " " +
                                        "SET Util_Logs_Bl = IIf(DBH >= " + DiameterVariables.MinDiaSmLogs.ToString().Trim() + " AND " +
                                                               "DBH >= " + DiameterVariables.MinDiaSteepSlope.ToString().Trim() + " AND " +
                                                               "DBH < " + DiameterVariables.MinDiaLgLogs.ToString().Trim() + ",'Y',Util_Logs_Bl) " +
                                        "WHERE species_group = " + strSpcGrp + " AND " +
                                              "diam_group = " + strDbhGrp;
                            x++;

                            /*******************************************************************************
                             **LARGE LOGS: Set the Logs Y/N flag depending if the diameters.dbh value is 
                             **greater or equal to the user entered MinLgLogsDBH 
                             **update value to "Y" in diameters table for Logs if DBH within min/max 
                             **for large log trees
                             *******************************************************************************/
                            strSql[x] = "UPDATE " + p_strDiametersTableName.Trim() + " " +
                                       "SET Util_Logs_Bl = IIf(DBH >= " + DiameterVariables.MinDiaLgLogs.ToString().Trim() + " AND " +
                                                              "DBH >= " + DiameterVariables.MinDiaSteepSlope.ToString().Trim() +
                                                              ",'Y',Util_Logs_Bl) " +
                                       "WHERE species_group = " + strSpcGrp + " AND " +
                                             "diam_group = " + strDbhGrp;
                            x++;
                        }
                        else
                        {
                            //EVERYTHING GETS PROCESSED AS CHIPS
                            /*******************************************************************************
                             **SMALL LOGS: Set the Chips Y/N flag depending if the diameters.dbh value falls
                             **between the user entered MinSmLogsDBH and MinLgLogsDBH
                             **update value to "Y" in diameters table for Logs if DBH within min/max 
                             **for small log trees
                             *******************************************************************************/
                            strSql[x] = "UPDATE " + p_strDiametersTableName.Trim() + " " +
                                        "SET Util_Chips_Bl = IIf(DBH >= " + DiameterVariables.MinDiaSmLogs.ToString().Trim() + " AND " +
                                                                "DBH >= " + DiameterVariables.MinDiaSteepSlope.ToString().Trim() + " AND " +
                                                                "DBH < " + DiameterVariables.MinDiaLgLogs + ",'Y',Util_Chips_Bl) " +
                                        "WHERE species_group = " + strSpcGrp + " AND " +
                                              "diam_group = " + strDbhGrp;
                            x++;

                            /*******************************************************************************
                             **LARGE LOGS: Set the Chips Y/N flag depending if the diameters.dbh value is 
                             **greater or equal to the user entered MinLgLogsDBH 
                             **update value to "Y" in diameters table for Logs if DBH within min/max 
                             **for large log trees
                             *******************************************************************************/
                            strSql[x] = "UPDATE " + p_strDiametersTableName.Trim() + " " +
                                       "SET Util_Chips_Bl = IIf(DBH >= " + DiameterVariables.MinDiaLgLogs.ToString().Trim() + " AND " +
                                                              "DBH >= " + DiameterVariables.MinDiaSteepSlope.ToString().Trim() +
                                                              ",'Y',Util_Chips_Bl) " +
                                       "WHERE species_group = " + strSpcGrp + " AND " +
                                             "diam_group = " + strDbhGrp;
                            x++;
                        }
                    }

                }
                m_oAdo.m_OleDbDataReader.Close();
                for (x = 0; x <= strSql.Length - 1; x++)
                {
                    if (strSql[x] == null) break;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, strSql[x] + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, strSql[x]);
                    if (m_oAdo.m_intError != 0) break;
                }
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;

        }
        private void RunScenario_AppendTreeData(string p_strTreeDataInTableName,string p_strDiametersTableName,string p_strVariant, string p_strRxPackage, bool p_bSteepSlope)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_AppendTreeData\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            string strDestColumns = "fvs_tree_id,biosum_cond_id,rxpackage,rx,rxcycle,spcd,dia,tpa,merch_vol_cf,merch_wt_gt,chip_vol_cf,chip_wt_gt,residue_fraction,slope,species_group";
            string strSourceColumns = "t.fvs_tree_id,t.biosum_cond_id,t.rxpackage,t.rx,t.rxcycle," + 
                                      "z.spcd,t.dbh AS dia,t.tpa," + 
                                      "t.volcfnet * t.tpa AS merch_vol_cf," + 
                                      "(t.drybiom * t.tpa / s.Dry_to_Green)/2000 AS merch_wt_gt," + 
                                      "(t.drybiot * t.tpa) / (s.Od_Wgt) AS chip_vol_cf," + 
                                      "(t.drybiot * t.tpa / s.Dry_to_Green)/2000  AS chip_wt_gt," + 
                                      "IIF(t.drybiom > 0, Round(((t.drybiot * t.tpa) - (t.drybiom * t.tpa)) / (t.drybiom * t.tpa) * 100,0),0) AS residue_fraction," + 
                                      "z.slope,s.user_spc_group AS species_group";

            string strTableName = "fvs_tree_IN_" + p_strVariant + "_P" + p_strRxPackage + "_TREE_CUTLIST";
            //
            //CREATE TEMP TREE TABLE WITH ONLY COLUMNS NEEDED
            //
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "treetemp"))
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE treetemp");
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "treetemp2"))
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE treetemp2");
           

            m_oAdo.m_strSQL = "SELECT t.fvs_tree_id,t.spcd, IIF(c.slope IS NULL,0,c.slope) AS slope " +
                              "INTO treetemp " +
                              "FROM " + m_oQueries.m_oFIAPlot.m_strTreeTable + " t," +
                                         m_oQueries.m_oFIAPlot.m_strCondTable + " c " +
                              "WHERE t.biosum_cond_id=c.biosum_cond_id AND mid(t.fvs_tree_id,1,2)='" + p_strVariant + "'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            if (m_oAdo.m_intError == 0)
            {
                //
                //CREATE TEMP TREE TABLE WITH ONLY COLUMNS NEEDED FOR FVS CREATED TREES
                //
                m_oAdo.m_strSQL = "SELECT z.fvs_tree_id,CINT(z.fvs_species) AS spcd, IIF(c.slope IS NULL,0,c.slope) AS slope " +
                                  "INTO treetemp2 " +
                                  "FROM " + strTableName + " z, " +
                                            m_oQueries.m_oFIAPlot.m_strCondTable + " c " +
                                  "WHERE z.rxpackage='" + p_strRxPackage + "' AND " +
                                        "z.biosum_cond_id=c.biosum_cond_id AND " +
                                        "z.fvscreatedtree_YN='Y' AND " +
                                        "mid(z.fvs_tree_id,1,2)='" + p_strVariant + "'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }



            if (m_oAdo.m_intError == 0)
            {
                //
                //SET NULL VALUES TO 0
                //
                m_oAdo.m_strSQL = "UPDATE " + strTableName + " SET volcfnet = IIF(volcfnet IS NULL,0,volcfnet)," +
                                                                  "volcfgrs = IIF(volcfgrs IS NULL,0,volcfgrs)," +
                                                                  "volcsgrs = IIF(volcsgrs IS NULL,0,volcsgrs)," +
                                                                  "drybiom  = IIF(drybiom IS NULL,0,drybiom)," +
                                                                  "drybiot  = IIF(drybiot IS NULL,0,drybiot)," + 
                                                                  "voltsgrs = IIF(voltsgrs IS NULL,0,voltsgrs)";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            if (m_oAdo.m_intError == 0)
            {
                //
                //CREATE TREEDATAIN TABLE
                //
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strTreeDataInTableName))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + p_strTreeDataInTableName);

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, Tables.ProcessorScenarioRun.CreateTreeBinTableSQL(p_strTreeDataInTableName) + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                frmMain.g_oTables.m_oProcessorScenarioRun.CreateTreeDataInTable(m_oAdo, m_oAdo.m_OleDbConnection, p_strTreeDataInTableName);
            }

            if (m_oAdo.m_intError == 0)
            {
                if (p_bSteepSlope == false)
                {
                    m_oAdo.m_strSQL = "SELECT  DISTINCT " + strSourceColumns + " " +
                                     "FROM " + strTableName + " t," +
                                      m_oQueries.m_oFvs.m_strTreeSpcTable + " s," +
                                     "treetemp z  " +
                                     "WHERE t.rxpackage='" + p_strRxPackage + "' AND " +
                                           "t.fvscreatedtree_YN='N' AND " +
                                           "t.fvs_tree_id=z.fvs_tree_id AND " +
                                           "(z.spcd = s.spcd AND s.fvs_variant='" + p_strVariant + "' AND " +
                                            "z.slope < " + ScenarioHarvestMethodVariables.SteepSlope.ToString() + ")";
                }
                else
                {
                    m_oAdo.m_strSQL = "SELECT  DISTINCT " + strSourceColumns + " " +
                                    "FROM " + strTableName + " t," +
                                     m_oQueries.m_oFvs.m_strTreeSpcTable + " s," +
                                    "treetemp z  " +
                                    "WHERE t.rxpackage='" + p_strRxPackage + "' AND " +
                                          "t.fvscreatedtree_YN='N' AND " +
                                          "t.fvs_tree_id=z.fvs_tree_id AND " +
                                          "(z.spcd = s.spcd AND s.fvs_variant='" + p_strVariant + "' AND " +
                                           "z.slope >= " + ScenarioHarvestMethodVariables.SteepSlope.ToString() + ")";
                }

                m_oAdo.m_strSQL = "INSERT INTO " + p_strTreeDataInTableName + " (" + strDestColumns + ") " +
                                  m_oAdo.m_strSQL;


                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");


                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }

            if (m_oAdo.m_intError == 0)
            {

                //fvs created trees
                if (p_bSteepSlope == false)
                {
                    m_oAdo.m_strSQL = "SELECT DISTINCT " + strSourceColumns + " " +
                                    "FROM " + strTableName + " t," +
                                     m_oQueries.m_oFvs.m_strTreeSpcTable + " s," +
                                    "treetemp2 z  " +
                                    "WHERE t.rxpackage='" + p_strRxPackage + "' AND " +
                                          "t.fvscreatedtree_YN='Y' AND " +
                                          "t.fvs_tree_id=z.fvs_tree_id AND " +
                                          "(z.spcd = s.spcd AND s.fvs_variant='" + p_strVariant + "' AND " +
                                           "z.slope < " + ScenarioHarvestMethodVariables.SteepSlope.ToString() + ")";
                }
                else
                {
                    m_oAdo.m_strSQL = "SELECT DISTINCT " + strSourceColumns + " " +
                                   "FROM " + strTableName + " t," +
                                    m_oQueries.m_oFvs.m_strTreeSpcTable + " s," +
                                   "treetemp2 z  " +
                                   "WHERE t.rxpackage='" + p_strRxPackage + "' AND " +
                                         "t.fvscreatedtree_YN='Y' AND " +
                                         "t.fvs_tree_id=z.fvs_tree_id AND " +
                                         "(z.spcd = s.spcd AND s.fvs_variant='" + p_strVariant + "' AND " +
                                          "z.slope >= " + ScenarioHarvestMethodVariables.SteepSlope.ToString() + ")";
                }

                m_oAdo.m_strSQL = "INSERT INTO " + p_strTreeDataInTableName + " (" + strDestColumns + ") " +
                                  m_oAdo.m_strSQL;


                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");


                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            if (m_oAdo.m_intError == 0)
            {
                //
                //UPDATE THE TREE DATA IN TABLE WITH THE DIAMETER GROUP
                //
                m_oAdo.m_strSQL = "UPDATE " + p_strTreeDataInTableName + " t " +
                                  "INNER JOIN " + p_strDiametersTableName + " d " +
                                  "ON t.DIA = d.DBH AND t.species_group = d.species_group " +
                                  "SET t.diam_group = d.diam_group";
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                //
                //REMOVE MERCH FROM CHIPS UNLESS MERCH IS PROCESSED AS CHIPS
                //
                m_oAdo.m_strSQL = "UPDATE " + p_strTreeDataInTableName + " t " +
                                  "INNER JOIN scenario_tree_species_diam_dollar_values s " +
                                  "ON t.species_group = s.species_group AND t.diam_group = s.diam_group " +
                                  "SET chip_vol_cf = IIF(s.wood_bin='M',chip_vol_cf - merch_vol_cf, chip_vol_cf)," +
                                      "chip_wt_gt =  IIF(s.wood_bin='M',chip_wt_gt - merch_wt_gt, chip_wt_gt)," +
                                      "merch_vol_cf = IIF(s.wood_bin='C',0,merch_vol_cf)," +
                                      "merch_wt_gt  = IIF(s.wood_bin='C',0,merch_wt_gt)," +
                                      "merch_to_chipbin_YN = IIF(s.wood_bin='C','Y','N') " +
                                  "WHERE TRIM(UCASE(s.scenario_id))='" + ScenarioId.Trim().ToUpper() + "'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
        }
        private void RunScenario_InitBinsTables(string p_strTreeDataInTableName,
                                                string p_strBinTableName,
                                                string p_strHwdBinTableName)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_InitBinsTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            //
            //CREATE THE TREE BINS TABLE
            //
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strBinTableName))
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + p_strBinTableName);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, Tables.ProcessorScenarioRun.CreateTreeBinTableSQL(p_strBinTableName) + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.CreateTreeBinTableSQL(p_strBinTableName));
            if (m_oAdo.m_intError == 0)
            {
                //
                //INSERT KEY VALUES
                //
                m_oAdo.m_strSQL = "INSERT INTO " + p_strBinTableName + " " +
                                  "(fvs_tree_id,biosum_cond_id,rxpackage,rx,rxcycle,species,species_group,diam_group,dbh) " +
                                     "SELECT fvs_tree_id,biosum_cond_id,rxpackage,rx,rxcycle," +
                                            "spcd AS species,species_group,diam_group,dia AS dbh " +
                                     "FROM " + p_strTreeDataInTableName;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            if (m_oAdo.m_intError == 0)
            {
                //
                //INITIALIZE VALUES
                //
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.InitializeBinsTable(p_strBinTableName, "BC");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.InitializeBinsTable(p_strBinTableName, "CHIPS");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.InitializeBinsTable(p_strBinTableName, "SMLOGS");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.InitializeBinsTable(p_strBinTableName, "LGLOGS");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            if (m_oAdo.m_intError == 0)
            {
                //
                //CREATE THE TREE HARDWOOD BINS TABLE
                //
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strHwdBinTableName))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + p_strHwdBinTableName);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, Tables.ProcessorScenarioRun.CreateTreeHwdBinTableSQL(p_strHwdBinTableName.Trim()) + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.CreateTreeHwdBinTableSQL(p_strHwdBinTableName.Trim()));
            }
            if (m_oAdo.m_intError == 0)
            {
                //
                //INSERT KEY VALUES
                //
                m_oAdo.m_strSQL = "INSERT INTO " + p_strHwdBinTableName + " " +
                                  "(fvs_tree_id,biosum_cond_id,rxpackage,rx,rxcycle,species,species_group,diam_group,dbh) " +
                                     "SELECT fvs_tree_id,biosum_cond_id,rxpackage,rx,rxcycle," +
                                            "spcd AS species,species_group,diam_group,dia AS dbh " +
                                     "FROM " + p_strTreeDataInTableName;
               
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            if (m_oAdo.m_intError == 0)
            {
                //
                //INITIALIZE VALUES
                //
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.InitializeHwdBinsTable(p_strHwdBinTableName, "BC");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.InitializeHwdBinsTable(p_strHwdBinTableName, "CHIPS");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.InitializeHwdBinsTable(p_strHwdBinTableName, "SMLOGS");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.InitializeHwdBinsTable(p_strHwdBinTableName, "LGLOGS");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
           


        }

        private void RunScenario_SumBinsByPlotRxSpcGrpDbhGrp(string p_strBinTableName,
                                                             string p_strHwdBinTableName,
                                                             string p_strBinSumTableName,
                                                             string p_strHwdBinSumTableName)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_SumBinsByPlotRxSpcGrpDbhGrp\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            //
            //DELETE bin_sums table
            //
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection,p_strBinSumTableName))
                 m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection,"DROP TABLE " + p_strBinSumTableName);
            //
            //CREATE bin_sums table
            //
            m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.SumBinByPlotRxSpcGrpDbhGrp(p_strBinTableName,p_strBinSumTableName);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            if (m_oAdo.m_intError == 0)
            {
                //
                //DELETE hwd_bin_sums table
                //
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strHwdBinSumTableName))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + p_strHwdBinSumTableName);
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.SumHwdBinByPlotRxSpcGrpDbhGrp(p_strHwdBinTableName, p_strHwdBinSumTableName);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
        }
        private void RunScenario_SumBinsByPlotRx(string p_strBinSumTableName,
                                                 string p_strHwdBinSumTableName,
                                                 string p_strBinTotalsTableName,
                                                 string p_strHwdBinTotalsTableName)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_SumBinsByPlotRx\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strBinTotalsTableName))
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + p_strBinTotalsTableName);
            m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.SumBinByPlotRx(p_strBinSumTableName, p_strBinTotalsTableName);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            if (m_oAdo.m_intError == 0)
            {

                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strHwdBinTotalsTableName))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + p_strHwdBinTotalsTableName);
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.SumHwdBinByPlotRx(p_strHwdBinSumTableName, p_strHwdBinTotalsTableName);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
        }
        private void RunScenario_InitOutputTable(string p_strBinSumTableName,string p_strBinOutputTableName)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_InitOutputTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            //
            //DELETE OUTPUT TABLE
            //
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strBinOutputTableName))
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + p_strBinOutputTableName);
            //
            //CREATE OUTPUT TABLE
            //
            m_oAdo.m_strSQL = Tables.ProcessorScenarioRun.CreateOutputTableSQL(p_strBinOutputTableName);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL);
            //
            //INSERT KEY VALUES
            //
            m_oAdo.m_strSQL = "INSERT INTO " + p_strBinOutputTableName + " " + 
                    "(BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE) " + 
                       "SELECT BIOSUM_COND_ID, RXPACKAGE, RX, RXCYCLE " + 
                       "FROM " + p_strBinSumTableName + " " + 
                       "GROUP BY BIOSUM_COND_ID, RXPACKAGE, RX, RXCYCLE";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL);
            if (m_oAdo.m_intError ==0) 
                frmMain.g_oTables.m_oProcessorScenarioRun.CreateOutputTableIndexes(m_oAdo, m_oAdo.m_OleDbConnection, p_strBinOutputTableName);
            if (m_oAdo.m_intError == 0)
            {
                //
                //INITIALIZE VALUES
                //
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.InitializeOutputTableValues(p_strBinOutputTableName);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;


        }
        private void RunScenario_UpdateOutputTable(string p_strBinOutputTableName,
                                                   string p_strBinTotalsTableName,
                                                   string p_strHwdBinTotalsTableName,
                                                   string p_strTreeDataInTableName,
                                                   string p_strDiametersTableName,
                                                   string p_strBiomassByLogSizeTableName,
                                                   string p_strBiomassByStandRxTableName,
                                                   string p_strVariant,
                                                   string p_strRxPackage)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_UpdateOutputTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            string strTableName = "fvs_tree_IN_" + p_strVariant + "_P" + p_strRxPackage + "_TREE_CUTLIST";
            //
            //Update yarding distance, slope  and elevation from PLOT,COND table on output
            //
            m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " o " +
                              "INNER JOIN (" + m_oQueries.m_oFIAPlot.m_strPlotTable + " p " +
                              "INNER JOIN " + m_oQueries.m_oFIAPlot.m_strCondTable + " c " +
                              "ON p.biosum_plot_id = c.biosum_plot_id) " +
                              "ON o.biosum_cond_id = c.biosum_cond_id " +
                              "SET o.gis_yard_dist_ft = IIF(p.gis_yard_dist_ft IS NULL,1,p.gis_yard_dist_ft)," +
                                  "o.elev = p.elev," + 
                                  "o.slope = IIF(c.slope IS NULL,0,c.slope)";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            if (m_oAdo.m_intError == 0)
            {

                //
                //Update output table with bin table values
                //
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.UpdateOutputTableFromBinTables(p_strBinOutputTableName, p_strBinTotalsTableName, p_strHwdBinTotalsTableName);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            if (m_oAdo.m_intError == 0)
            {

                //
                //DELETE SUMBIOMASSBYLOGSIZE TABLE
                //
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strBiomassByLogSizeTableName))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + p_strBiomassByLogSizeTableName);
                //
                //Sum Biomass By Log Size into SUMBIOMASSBYLOGSIZE table
                //
                m_oAdo.m_strSQL = "SELECT t.biosum_cond_id, t.rxpackage, t.rx, t.rxcycle,d.class," +
                                          "SUM(ROUND(f.drybiot,0)) AS TreeBiomass," +
                                          "SUM(ROUND(f.drybiom,0)) AS BoleBiomass " +
                                  "INTO " + p_strBiomassByLogSizeTableName + " " +
                                  "FROM (" + p_strTreeDataInTableName + " t " +
                                  "INNER JOIN " + p_strDiametersTableName + " d " +
                                  "ON t.dia = d.DBH AND t.species_group=d.species_group) " +
                                  "INNER JOIN " + strTableName + " f " +
                                  "ON (t.fvs_tree_id = f.fvs_tree_id) " +
                                  "GROUP BY t.biosum_cond_id, t.rxpackage,t.rx,t.rxcycle, d.class " +
                                  "HAVING (((TRIM(d.class))='SMLOGS' OR (TRIM(d.class))='LGLOGS'))";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            if (m_oAdo.m_intError == 0)
            {

                //
                //Update OUTPUT table with  residue fraction for small logs from the SUMBIOMASSBYLOGSIZE table
                //
                m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " o " +
                                  "INNER JOIN " + p_strBiomassByLogSizeTableName + " s " +
                                  "ON  o.BIOSUM_COND_ID = s.biosum_cond_id AND " +
                                      "o.rxpackage = s.rxpackage AND " +
                                      "o.rx = s.rx AND " +
                                      "o.rxcycle = s.rxcycle  " +
                                  "SET o.[SMLOGS Chip Fraction] = IIF(s.BoleBiomass=0,0,ROUND((s.TreeBiomass-s.BoleBiomass)/s.BoleBiomass*100,0)) " +
                                  "WHERE (((TRIM(s.class))='SMLOGS'))";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            if (m_oAdo.m_intError == 0)
            {


                //
                //Update OUTPUT table with  residue fraction for large logs from the SUMBIOMASSBYLOGSIZE table
                //
                m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " o " +
                                  "INNER JOIN " + p_strBiomassByLogSizeTableName + " s " +
                                  "ON o.biosum_cond_id = s.biosum_cond_id AND " +
                                     "o.rxpackage = s.rxpackage AND " +
                                     "o.rx = s.rx AND " +
                                     "o.rxcycle = s.rxcycle " +
                                  "SET o.[LGLOGS Chip Fraction] = IIF(s.BoleBiomass=0,0,ROUND((s.TreeBiomass-s.BoleBiomass)/s.BoleBiomass*100,0)) " +
                                  "WHERE (((TRIM(s.class))='LGLOGS'))";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            if (m_oAdo.m_intError == 0)
            {

                //
                //DELETE SUMBIOMASSBYSTANDRX TABLE
                //
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strBiomassByStandRxTableName))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + p_strBiomassByStandRxTableName);
                //
                //Sum Biomass By Stand+Rx into SUMBIOMASSBYSTANDRX table
                //
                m_oAdo.m_strSQL = "SELECT t.biosum_cond_id, t.rxpackage, t.rx, t.rxcycle," +
                                          "SUM(ROUND(f.drybiot,0)) AS TreeBiomass," +
                                          "SUM(ROUND(f.drybiom,0)) AS BoleBiomass " +
                                  "INTO " + p_strBiomassByStandRxTableName + " " +
                                  "FROM (" + p_strTreeDataInTableName + " t " +
                                  "INNER JOIN " + p_strDiametersTableName + " d " +
                                  "ON t.dia = d.DBH AND t.species_group=d.species_group) " +
                                  "INNER JOIN " + strTableName + " f " +
                                  "ON (t.fvs_tree_id = f.fvs_tree_id) " +
                                  "GROUP BY t.biosum_cond_id, t.rxpackage,t.rx,t.rxcycle, d.class";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            if (m_oAdo.m_intError == 0)
            {

                //
                //Update OUTPUT table with TOTAL residue fraction for a STAND and TREATMENT from the SUMBIOMASSBYSTANDRX table
                //
                m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " o " +
                                  "INNER JOIN " + p_strBiomassByStandRxTableName + " s " +
                                  "ON  o.BIOSUM_COND_ID = s.biosum_cond_id AND " +
                                      "o.rxpackage = s.rxpackage AND " +
                                      "o.rx = s.rx AND " +
                                      "o.rxcycle = s.rxcycle  " +
                                  "SET o.[TOTAL Chip Fraction] = IIF(s.BoleBiomass=0,0,ROUND((s.TreeBiomass-s.BoleBiomass)/s.BoleBiomass*100,0))";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            if (m_oAdo.m_intError == 0)
            {

                //
                //Update OUTPUT table with  utilized chips for small and large logs from the SUMBIOMASSBYLOGSIZE table
                //
                m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " o " +
                                  "SET o.[SMLOGS utilized chips (tons)] = " +
                                          "IIF(o.[SMLOGS utilized chips (tons)] IS NOT NULL," +
                                              "o.[SMLOGS utilized chips (tons)]*(o.[SMLOGS Chip Fraction]/100),0)," +
                                      "o.[LGLOGS utilized chips (tons)] = " +
                                          "IIF(o.[LGLOGS utilized chips (tons)] IS NOT NULL," +
                                              "o.[LGLOGS utilized chips (tons)]*(o.[LGLOGS Chip Fraction]/100),0)," +
                                      "o.[TOTAL utilized chips (tons)] = " +
                                          "IIF(o.[BRUSH CUT utilized chips (tons)] IS NOT NULL," +
                                              "o.[BRUSH CUT utilized chips (tons)]*(o.[TOTAL Chip Fraction]/100),0)";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            if (m_oAdo.m_intError == 0)
            {

                //
                //Update OUTPUT table null values with zero
                //
                m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " o " +
                                   "SET o.[CHIPS Average Vol (ft3)] = " +
                                            "IIf(o.[CHIPS Average Vol (ft3)] Is Null,0,o.[CHIPS Average Vol (ft3)])," +
                                       "o.[CHIPS Average Weight (tons)] = " +
                                            "IIf(o.[CHIPS Average Weight (tons)] Is Null,0,o.[CHIPS Average Weight (tons)])," +
                                       "o.[CHIPS Average Density (lbs/ft3)] = " +
                                            "IIf(o.[CHIPS Average Density (lbs/ft3)] Is Null,0,o.[CHIPS Average Density (lbs/ft3)])," +
                                       "o.[CHIPS Hwd Proportion] = " +
                                            "IIf(o.[CHIPS Hwd Proportion] Is Null,0,o.[CHIPS Hwd Proportion])," +
                                       "o.[SMLOGS Average Vol (ft3)] = " +
                                            "IIf(o.[SMLOGS Average Vol (ft3)] Is Null,0,o.[SMLOGS Average Vol (ft3)])," +
                                       "o.[SMLOGS Average Weight (tons)] = " +
                                            "IIf(o.[SMLOGS Average Weight (tons)] Is Null,0,o.[SMLOGS Average Weight (tons)])," +
                                       "o.[SMLOGS Average Density (lbs/ft3)] = " +
                                            "IIf(o.[SMLOGS Average Density (lbs/ft3)] Is Null,0,o.[SMLOGS Average Density (lbs/ft3)])," +
                                       "o.[SMLOGS Hwd Proportion] = " +
                                            "IIf(o.[SMLOGS Hwd Proportion] Is Null,0,o.[SMLOGS Hwd Proportion])," +
                                       "o.[LGLOGS Average Vol (ft3)] = " +
                                            "IIf(o.[LGLOGS Average Vol (ft3)] Is Null,0,o.[LGLOGS Average Vol (ft3)])," +
                                       "o.[LGLOGS Average Weight (tons)] = " +
                                            "IIf(o.[LGLOGS Average Weight (tons)] Is Null,0,o.[LGLOGS Average Weight (tons)])," +
                                       "o.[LGLOGS Average Density (lbs/ft3)] = " +
                                            "IIf(o.[LGLOGS Average Density (lbs/ft3)] Is Null,0,o.[LGLOGS Average Density (lbs/ft3)])," +
                                       "o.[LGLOGS Hwd Proportion] = " +
                                            "IIf(o.[LGLOGS Hwd Proportion] Is Null,0,o.[LGLOGS Hwd Proportion])," +
                                       "o.[TOTAL Average Vol (ft3)] = " +
                                            "IIf(o.[TOTAL Average Vol (ft3)] Is Null,0,o.[TOTAL Average Vol (ft3)])," +
                                       "o.[TOTAL Average Weight (tons)] = " +
                                            "IIf(o.[TOTAL Average Weight (tons)] Is Null,0,o.[TOTAL Average Weight (tons)])," +
                                       "o.[TOTAL Average Density (lbs/ft3)] = " +
                                            "IIf(o.[TOTAL Average Density (lbs/ft3)] Is Null,0,o.[TOTAL Average Density (lbs/ft3)])," +
                                       "o.[TOTAL Hwd Proportion] = " +
                                            "IIf(o.[TOTAL Hwd Proportion] Is Null,0,o.[TOTAL Hwd Proportion])";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;


        }
        private void RunScenario_UpdateOutputTableWithHarvestingSystem(string p_strBinOutputTableName,
                                                                       bool  p_bSteepSlope)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_UpdateOutputTableWithHarvestingSystem\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            int x;
            if (ScenarioHarvestMethodVariables.UseDefaultRxHarvestMethod == false)
            {
                if (p_bSteepSlope == false)
                {
                    m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " SET [Harvesting system] = '" + ScenarioHarvestMethodVariables.HarvestMethodLowSlope.Trim() + "'";
                }
                else
                {
                    m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " SET [Harvesting system] = '" + ScenarioHarvestMethodVariables.HarvestMethodSteepSlope.Trim() + "'";
                }
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            else
            {
                //find the default harvesting method for each treatment
                if (this.m_oRxPackageItem.SimulationYear1Rx != "000")
                {
                    for (x = 0; x <= m_oRxItem_Collection.Count - 1; x++)
                    {
                        if (m_oRxItem_Collection.Item(x).RxId.Trim() == this.m_oRxPackageItem.SimulationYear1Rx)
                        {
                            if (p_bSteepSlope == false)
                            {
                                m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " " +
                                                  "SET [Harvesting system] = '" +
                                                     m_oRxItem_Collection.Item(x).HarvestMethodLowSlope.Trim() + "' " +
                                                  "WHERE rxpackage='" + m_oRxPackageItem.RxPackageId.Trim() + "' AND " +
                                                        "rx='" + m_oRxPackageItem.SimulationYear1Rx.Trim() + "' AND " +
                                                        "rxcycle='1'";
                            }
                            else
                            {
                                m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " " +
                                                 "SET [Harvesting system] = '" +
                                                    m_oRxItem_Collection.Item(x).HarvestMethodSteepSlope.Trim() + "' " +
                                                 "WHERE rxpackage='" + m_oRxPackageItem.RxPackageId.Trim() + "' AND " +
                                                       "rx='" + m_oRxPackageItem.SimulationYear1Rx.Trim() + "' AND " +
                                                       "rxcycle='1'";
                            }
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                            break;

                        }
                    }
                }
                if (m_oAdo.m_intError == 0)
                {

                    if (this.m_oRxPackageItem.SimulationYear2Rx != "000")
                    {
                        for (x = 0; x <= m_oRxItem_Collection.Count - 1; x++)
                        {
                            if (m_oRxItem_Collection.Item(x).RxId.Trim() == this.m_oRxPackageItem.SimulationYear2Rx)
                            {
                                if (p_bSteepSlope == false)
                                {
                                    m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " " +
                                                      "SET [Harvesting system] = '" +
                                                         m_oRxItem_Collection.Item(x).HarvestMethodLowSlope.Trim() + "' " +
                                                      "WHERE rxpackage='" + m_oRxPackageItem.RxPackageId.Trim() + "' AND " +
                                                            "rx='" + m_oRxPackageItem.SimulationYear2Rx.Trim() + "' AND " +
                                                            "rxcycle='2'";
                                }
                                else
                                {
                                    m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " " +
                                                      "SET [Harvesting system] = '" +
                                                         m_oRxItem_Collection.Item(x).HarvestMethodSteepSlope.Trim() + "' " +
                                                      "WHERE rxpackage='" + m_oRxPackageItem.RxPackageId.Trim() + "' AND " +
                                                            "rx='" + m_oRxPackageItem.SimulationYear2Rx.Trim() + "' AND " +
                                                            "rxcycle='2'";
                                }
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                                break;

                            }
                        }
                    }
                }
                if (m_oAdo.m_intError == 0)
                {

                    if (this.m_oRxPackageItem.SimulationYear3Rx != "000")
                    {
                        for (x = 0; x <= m_oRxItem_Collection.Count - 1; x++)
                        {
                            if (m_oRxItem_Collection.Item(x).RxId.Trim() == this.m_oRxPackageItem.SimulationYear3Rx)
                            {
                                if (p_bSteepSlope == false)
                                {
                                    m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " " +
                                                      "SET [Harvesting system] = '" +
                                                         m_oRxItem_Collection.Item(x).HarvestMethodLowSlope.Trim() + "' " +
                                                      "WHERE rxpackage='" + m_oRxPackageItem.RxPackageId.Trim() + "' AND " +
                                                            "rx='" + m_oRxPackageItem.SimulationYear3Rx.Trim() + "' AND " +
                                                            "rxcycle='3'";
                                }
                                else
                                {
                                    m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " " +
                                                     "SET [Harvesting system] = '" +
                                                        m_oRxItem_Collection.Item(x).HarvestMethodSteepSlope.Trim() + "' " +
                                                     "WHERE rxpackage='" + m_oRxPackageItem.RxPackageId.Trim() + "' AND " +
                                                           "rx='" + m_oRxPackageItem.SimulationYear3Rx.Trim() + "' AND " +
                                                           "rxcycle='3'";
                                }
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                                break;

                            }
                        }
                    }
                }
                if (m_oAdo.m_intError == 0)
                {

                    if (this.m_oRxPackageItem.SimulationYear4Rx != "000")
                    {
                        for (x = 0; x <= m_oRxItem_Collection.Count - 1; x++)
                        {
                            if (m_oRxItem_Collection.Item(x).RxId.Trim() == this.m_oRxPackageItem.SimulationYear4Rx)
                            {
                                if (p_bSteepSlope == false)
                                {
                                    m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " " +
                                                      "SET [Harvesting system] = '" +
                                                         m_oRxItem_Collection.Item(x).HarvestMethodLowSlope.Trim() + "' " +
                                                      "WHERE rxpackage='" + m_oRxPackageItem.RxPackageId.Trim() + "' AND " +
                                                            "rx='" + m_oRxPackageItem.SimulationYear4Rx.Trim() + "' AND " +
                                                            "rxcycle='4'";
                                }
                                else
                                {
                                    m_oAdo.m_strSQL = "UPDATE " + p_strBinOutputTableName + " " +
                                                      "SET [Harvesting system] = '" +
                                                         m_oRxItem_Collection.Item(x).HarvestMethodSteepSlope.Trim() + "' " +
                                                      "WHERE rxpackage='" + m_oRxPackageItem.RxPackageId.Trim() + "' AND " +
                                                            "rx='" + m_oRxPackageItem.SimulationYear4Rx.Trim() + "' AND " +
                                                            "rxcycle='4'";
                                }
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                                break;

                            }
                        }
                    }
                }
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
        }
        
        private void RunScenario_InitTreeVolVal(string p_strBinSumTableName,
                                                string p_strHwdBinSumTableName,
                                                string p_strVariant, 
                                                string p_strSlopeExpr)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_InitTreeVolVal\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            
            //
            //DELETE WORK TABLE
            //
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection,"tree_vol_by_spc_grp_dbh_grp"))
                 m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection,"DROP TABLE tree_vol_by_spc_grp_dbh_grp");
            //
            //POPULATE WORK TABLE
            //
            m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.SumBinSumTableByPlotRxSpcGrpDbhGrp
                                ("tree_vol_by_spc_grp_dbh_grp", p_strBinSumTableName, p_strHwdBinSumTableName);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL);
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
        }
        private void RunScenario_AppendToTreeVolVal(string p_strTreeVolValTableName)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_AppendDataToTreeVolVal\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            //
            //Populate the TREE_VOL_VAL_BY_SPECIES_DIAM_GROUPS WORK table with TREE_VOL_BY_SPC_GRP_DBH_GRP table.
            //
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strTreeVolValTableName) == false)
                frmMain.g_oTables.m_oProcessor.CreateTreeVolValSpeciesDiamGroupsWorkTable(
                      m_oAdo, m_oAdo.m_OleDbConnection, p_strTreeVolValTableName);

            m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.AppendToTreeVolVal("tree_vol_by_spc_grp_dbh_grp", p_strTreeVolValTableName, m_strDateTimeCreated);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
        }
        
        private void RunScenario_UpdateTreeVolValWithMerchChipMarketValues(string p_strTreeVolValTableName)
        {
            //
            //Update the TREE_VOL_VAL_BY_SPECIES_DIAM_GROUPS table with Chip and Merch market values by species groups and diameter groups.
            //
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_UpdateTreeVolValWithMerchChipMarketValues\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.UpdateTreeVolValWithMerchChipMarketValues
                 ("scenario_tree_species_diam_dollar_values",
                  "scenario_cost_revenue_escalators",
                  ScenarioId,
                  p_strTreeVolValTableName);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            //
            //remove any nulls
            //
            m_oAdo.m_strSQL = "UPDATE " + p_strTreeVolValTableName + " " + 
                             "SET chip_val_dpa = IIF(chip_val_dpa IS NULL,0,chip_val_dpa)," +
                                 "merch_val_dpa = IIF(merch_val_dpa IS NULL,0,merch_val_dpa)";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
        }
        private void RunScenario_InitFRCSInputTable(string p_strBinOutputTableName,string p_strFRCSInputTableName)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_InitFRCSInputTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.PopulateFRCSInputTable(p_strFRCSInputTableName, p_strBinOutputTableName);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            if (m_oAdo.m_intError == 0)
            {
                m_oAdo.m_strSQL = "UPDATE " + p_strFRCSInputTableName + " SET [One-way Yarding Distance] = 1 WHERE [One-way Yarding Distance] < 1";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
        }
        private void RunScenario_InitOPCOSTInputTable(string p_strBinOutputTableName, string p_strOPCOSTInputTableName)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_InitOPCOSTInputTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.PopulateOPCOSTInputTable(p_strOPCOSTInputTableName, p_strBinOutputTableName);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            if (m_oAdo.m_intError == 0)
            {
                m_oAdo.m_strSQL = "UPDATE " + p_strOPCOSTInputTableName + " SET [One-way Yarding Distance] = 1 WHERE [One-way Yarding Distance] < 1";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
        }
        private void RunScenario_AppendToFRCSInputTable(string p_strDestTableName, string p_strSourceTableName)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_AppendToFRCSInputTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.AppendToFRCSInputTable(p_strDestTableName, p_strSourceTableName);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            if (m_oAdo.m_intError == 0)
            {
                m_oAdo.m_strSQL = "UPDATE " + p_strDestTableName + " SET [One-way Yarding Distance] = 1 WHERE [One-way Yarding Distance] < 1";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
        }
        private void RunScenario_AppendToOPCOSTInputTable(string p_strDestTableName, string p_strSourceTableName)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_AppendToOPCOSTInputTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.AppendToOPCOSTInputTable(p_strDestTableName, p_strSourceTableName);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            if (m_oAdo.m_intError == 0)
            {
                m_oAdo.m_strSQL = "UPDATE " + p_strDestTableName + " SET [One-way Yarding Distance] = 1 WHERE [One-way Yarding Distance] < 1";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
        }
 
        private void RunScenario_UpdateFRCSThresholds()
        {
            int x;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//UpdateFRCSThresholds\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strTempDir + "\\FIA_BIOSUM_FRCS_UPDATETHRESHOLDS.TXT", "get FRCS batch output from DataForAccess Sheet");
            //
            //START FRCS
            //
            m_oExcel.DisplayAlerts = false;
            m_oExcel.StartExcelApplication();
            //open the workbook found in value m_oExcel.ExcelFileName
            if (m_oExcel.m_intError == 0)
            {
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Open Excel Workbook: " + m_oExcel.ExcelFileName + "\r\n");
                m_oExcel.OpenWorkBook();
            }
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Error code: " + m_oExcel.m_intError.ToString().Trim() + "\r\n");
            if (m_oExcel.m_intError == 0)
            {
                m_oExcel.Hide();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Assign Current Worksheet <Outputs>\r\n");
                m_oExcel.AssignCurrentSheet("Outputs");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Error code: " + m_oExcel.m_intError.ToString().Trim() + "\r\n");
            }
            if (m_oExcel.m_intError == 0)
            {

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Unprotect <Outputs> worksheet\r\n");
                m_oExcel.UnprotectSheet();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Error code: " + m_oExcel.m_intError.ToString().Trim() + "\r\n");
            }
            if (m_oExcel.m_intError == 0)
            {
                //Delete all this code; Obsolete
            }
            if (m_oExcel.m_intError == 0)
            {
                m_oExcel.DisplayAlerts = false;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Save And Exit Workbook\r\n");
                m_oExcel.SaveAndExit();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Error code: " + m_oExcel.m_intError.ToString().Trim() + "\r\n");
            }
            System.IO.File.Delete(frmMain.g_oEnv.strTempDir + "\\FIA_BIOSUM_FRCS_UPDATETHRESHOLDS.TXT");
            System.Threading.Thread.Sleep(10000);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Release Excel Com Objects\r\n");
            m_oExcel.ReleaseComObjects();
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Error code: " + m_oExcel.m_intError.ToString().Trim() + "\r\n");
            m_intError = m_oExcel.m_intError;
            m_strError = m_oExcel.m_strError;

        }
        private void RunScenario_ProcessOPCOST(string p_strVariant, string p_strRxPackage)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_ProcessOPCOST\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            string strFile = "";
            //add the year
            //ALTER TABLE YEAR COLUMN
            m_oAdo.m_strSQL = "ALTER TABLE OPCOST_INPUT ALTER  COLUMN [YearCostCalc] INTEGER;";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

            string strTable = "fvs_tree_IN_" + p_strVariant + "_P" + p_strRxPackage + "_TREE_CUTLIST";
            m_oAdo.m_strSQL = "DROP TABLE temp_year";
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "temp_year"))
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            m_oAdo.m_strSQL = "SELECT DISTINCT biosum_cond_id+rxpackage+rx+rxcycle AS STAND,RXYEAR " +
                              "INTO temp_year " +
                              "FROM " + strTable;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            m_oAdo.m_strSQL = "UPDATE opcost_input a INNER JOIN temp_year b ON a.STAND=b.STAND SET a.YearCostCalc=CINT(b.RXYEAR)";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            m_oAdo.m_strSQL = "DROP TABLE temp_year";
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

            string strOPCOSTErrorFilePath = frmMain.g_oFrmMain.getProjectDirectory() + "\\OPCOST\\Input\\" +
                                             p_strVariant + "_" + p_strRxPackage + "_opcost_error_log.txt";
            bool bOPCOSTWindow = RunScenario_CreateOPCOSTBatchFile(strOPCOSTErrorFilePath);
            RunScenario_ExecuteOPCOST(bOPCOSTWindow, strOPCOSTErrorFilePath);

            ado_data_access oAdo = new ado_data_access();
            using (var oConn = new System.Data.OleDb.OleDbConnection(m_oAdo.m_OleDbConnection.ConnectionString))
            {
                oConn.Open();
                if (oAdo.TableExist(oConn.ConnectionString, Tables.ProcessorScenarioRun.DefaultOpcostErrorsTableName))
                {
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Querying " + Tables.ProcessorScenarioRun.DefaultOpcostErrorsTableName + "table \r\n");
                    int intCount = Convert.ToInt32(oAdo.getRecordCount(oConn, "select count(*) from " + Tables.ProcessorScenarioRun.DefaultOpcostErrorsTableName,
                        Tables.ProcessorScenarioRun.DefaultOpcostErrorsTableName));

                    if (intCount > 0)
                    {
                        // This is the first error so we pop a message
                        if (_lstErrorVariants.Count == 0)
                        {
                            string strMessage = "Costs could not be estimated for " + intCount +
                                                " stands in Variant " + p_strVariant + " Sequence " + p_strRxPackage +
                                                ". Do you wish to continue ? ";
                            DialogResult res = MessageBox.Show(strMessage, "FIA Biosum", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (res != DialogResult.Yes)
                            {
                                m_intError = -1;
                                return;
                            }
                        }
                            _lstErrorVariants.Add(p_strVariant + "|" + p_strRxPackage);
                    }
                 }
                 oAdo = null;
            }
        }
        
        private void RunScenario_ExecuteOPCOST(bool bOPCOSTWindow, string strOPCOSTErrorFilePath)
        {
            //close the open connection
            string strConn = m_oAdo.m_OleDbConnection.ConnectionString;
            string strDb = m_oQueries.m_strTempDbFile;
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            System.Threading.Thread.Sleep(5000);

            System.Diagnostics.Process proc = new System.Diagnostics.Process();

            proc.StartInfo.RedirectStandardOutput = false;
            proc.StartInfo.RedirectStandardInput = false;
            proc.StartInfo.RedirectStandardError = false;
            if (! bOPCOSTWindow)
            {
                //suppress opCost window
                proc.StartInfo.CreateNoWindow = true;

            }
            else
            {
                proc.StartInfo.CreateNoWindow = false;
                proc.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
            }
            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.WorkingDirectory = frmMain.g_oEnv.strTempDir;
            proc.StartInfo.ErrorDialog = false;
            proc.EnableRaisingEvents = false;

            proc.StartInfo.FileName = "cmd.exe";
            proc.StartInfo.Arguments = "/c START /B /WAIT " + m_strOPCOSTBatchFile;

            try
            {
                proc.Start();

                proc.WaitForExit();
            }
            catch (Exception err)
            {
                m_intError = -1;
                m_strError = err.Message;
                MessageBox.Show("OPCOST Processing Error", "FIA Biosum");
            }

            m_oAdo.OpenConnection(strConn);

            if (!m_oAdo.TableExist(m_oAdo.m_OleDbConnection,"OPCOST_OUTPUT"))
            {
                m_intError=-1;
                m_strError="!!OPCOST processing did not complete successfully. Check the error log at " +
                            strOPCOSTErrorFilePath + " for details!!";

                MessageBox.Show(m_strError,"FIA Biosum",MessageBoxButtons.OK,MessageBoxIcon.Error);


            }


        }
        private bool RunScenario_CreateOPCOSTBatchFile(string strOPCOSTErrorFilePath)
        {
            
            //create a batch file containing the command
            System.IO.FileStream oTextFileStream;
            System.IO.StreamWriter oTextStreamWriter;
            oTextFileStream = new System.IO.FileStream(m_strOPCOSTBatchFile, System.IO.FileMode.Create,
                    System.IO.FileAccess.Write);
            oTextStreamWriter = new System.IO.StreamWriter(oTextFileStream);

            oTextStreamWriter.Write("@ECHO OFF\r\n");
            oTextStreamWriter.Write("ECHO ****************************************************\r\n");
            oTextStreamWriter.Write("ECHO *BIOSUM OPCOST\r\n");
            oTextStreamWriter.Write("ECHO ****************************************************\r\n\r\n");
            oTextStreamWriter.Write("TITLE=BIOSUM OPCOST\r\n\r\n");
            oTextStreamWriter.Write("SET RFILE=" + frmMain.g_strRDirectory + "\r\n");
            oTextStreamWriter.Write("SET OPCOSTRFILE=" + frmMain.g_strOPCOSTDirectory + "\r\n");
            oTextStreamWriter.Write("SET INPUTFILE=" + m_oQueries.m_strTempDbFile + "\r\n");
            oTextStreamWriter.Write("SET CONFIGFILE=" + m_strOPCOSTRefPath + "\r\n");
            oTextStreamWriter.Write("SET ERRORFILE=" + strOPCOSTErrorFilePath +  "\r\n");
            oTextStreamWriter.Write("SET PATH=" + frmMain.g_oUtils.getDirectory(frmMain.g_strRDirectory).Trim() + ";%PATH%\r\n\r\n");
            string strRedirect = " 2> " + "\"" + "%ERRORFILE%" + "\"";
            // Suppress OpCost window if debugging is turned off OR debug level < 3
            bool bOPCOSTWindow = frmMain.g_bDebug;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel < 3)
                bOPCOSTWindow = false;
            if (! bOPCOSTWindow)
            {
                //OpCost window is suppressed so we write standard out to log
                strRedirect = "> \"" + "%ERRORFILE%" + "\"" + " 2>&1";
            }
            oTextStreamWriter.Write("\"" + "%RFILE%" + "\"" + " " + "\"" + "%OPCOSTRFILE%" + "\"" + " " + "\"" + "%INPUTFILE%" + "\"" + " \"" + "%CONFIGFILE%" + "\"" + strRedirect + "\r\n\r\n");
            oTextStreamWriter.Write("EXIT\r\n");
            oTextStreamWriter.Close();
            oTextFileStream.Close();
            return bOPCOSTWindow;
        
        }
        private void RunScenario_ProcessFRCS()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_ProcessFRCS\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            //
            //POPULATE FRCS INPUT SHEET 
            //
            m_oAdo.m_strSQL = "SELECT * INTO [Excel 12.0;DATABASE=" + m_oExcel.ExcelFileName + ";HDR=YES;].[input] from [frcs_input]";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            if (m_oAdo.m_intError == 0)
            {
                //
                //START FRCS
                //
                m_oExcel.DisplayAlerts = false;
                m_oExcel.StartExcelApplication();
                if (m_oExcel.m_intError == 0)
                {
                    //open the workbook found in value m_oExcel.ExcelFileName
                    m_oExcel.OpenWorkBook();
                }
                if (m_oExcel.m_intError == 0)
                {
                    m_oExcel.Hide();
                    //process the batch input
                    m_oExcel.RunMacro("ProcessBatchInputSheet");
                }
                if (m_oExcel.m_intError == 0)
                {
                    m_oExcel.SaveAndExit();
                }
                m_oExcel.ReleaseComObjects();
            }
            if (m_oAdo.m_intError == 0 && m_oExcel.m_intError == 0)
            {
                //
                //GET THE FRCS OUTPUT
                //
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "FRCS_output"))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE FRCS_output");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.CreateFRCSOutputTableSQL("FRCS_output"));
                if (m_oAdo.m_intError == 0)
                {
                    frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strTempDir + "\\FIA_BIOSUM_FRCS_GETBATCHOUTPUT.TXT", "get FRCS batch output from DataForAccess Sheet");
                    System.Threading.Thread.Sleep(10000);
                    //FRCS batch output is found on sheet DataForAccess
                    m_oAdo.m_strSQL = @"INSERT INTO FRCS_output SELECT * FROM [Excel 12.0;HDR=Yes;IMEX=1;DATABASE=" + this.m_oExcel.ExcelFileName + "].[DataForAccess$] ;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    System.IO.File.Delete(frmMain.g_oEnv.strTempDir + "\\FIA_BIOSUM_FRCS_GETBATCHOUTPUT.TXT");
                }
            }
            if (m_oAdo.m_intError != 0)
            {
                m_intError = m_oAdo.m_intError;
                m_strError = m_oAdo.m_strError;
            }
            else if (m_oExcel.m_intError != 0)
            {
                m_intError = m_oExcel.m_intError;
                m_strError = m_oExcel.m_strError;
            }
            
        }
        private void RunScenario_FRCSWarningMessages()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_FRCSWarningMessages\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            string strHarvestMethodsList="";
            string[] strHarvestMethodsArray=null;
            int x, y;

            //
            //FATAL ERROR MESSAGES
            //
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "FRCS_fatal_error_messages_work_table"))
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE FRCS_fatal_error_messages_work_table");
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "FRCS_fatal_error_messages"))
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE FRCS_fatal_error_messages");

            m_oAdo.m_strSQL = "SELECT MID(STAND,1,25) AS biosum_cond_id," +
                                     "MID(rxpackage_rx_rxcycle,1,3) AS RXPackage," +
                                     "MID(rxpackage_rx_rxcycle,4,3) AS RX," +
                                     "MID(rxpackage_rx_rxcycle,7,1) AS RXCycle," +
                                     "[$/Acre] " +
                              "INTO FRCS_fatal_error_messages_work_table " +
                              "FROM FRCS_output " +
                              "WHERE LEN(TRIM(STAND)) > 30 AND " +
                                    "[$/Acre] IS NOT NULL AND " +
                                    "LEN(TRIM([$/Acre])) > 0";



            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            if (m_oAdo.m_intError == 0)
            {
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "FRCS_fatal_error_messages_work_table"))
                {

                    m_oAdo.m_strSQL = "ALTER TABLE FRCS_fatal_error_messages_work_table ALTER COLUMN biosum_cond_id CHAR(25)";
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    m_oAdo.m_strSQL = "ALTER TABLE FRCS_fatal_error_messages_work_table ALTER COLUMN RXPackage CHAR(3)";
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    m_oAdo.m_strSQL = "ALTER TABLE FRCS_fatal_error_messages_work_table ALTER COLUMN RX CHAR(3)";
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    m_oAdo.m_strSQL = "ALTER TABLE FRCS_fatal_error_messages_work_table ALTER COLUMN RXCycle CHAR(1)";
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    frmMain.g_oTables.m_oProcessorScenarioRun.CreateFRCSWarningTableIndexes(m_oAdo, m_oAdo.m_OleDbConnection, "FRCS_fatal_error_messages_work_table");

                    m_oAdo.m_strSQL = "SELECT biosum_cond_id,RXPackage,RX,RXCycle," +
                                             "[$/Acre] AS FRCS_Fatal_Error_Message " +
                                      "INTO FRCS_fatal_error_messages " +
                                      "FROM FRCS_fatal_error_messages_work_table " +
                                      "WHERE ASC([$/Acre]) > 64 AND " +
                                            "ASC([$/Acre]) < 122 OR " +
                                            "ASC([$/Acre]) = 32 OR " +
                                            "ASC([$/Acre]) = 36 ";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "FRCS_fatal_error_messages"))
                        frmMain.g_oTables.m_oProcessorScenarioRun.CreateFRCSWarningTableIndexes(m_oAdo, m_oAdo.m_OleDbConnection, "FRCS_fatal_error_messages");
                }
            }
            if (m_oAdo.m_intError == 0)
            {
                //
                //WARNING MESSAGES
                //
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "FRCS_warning_messages"))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE FRCS_warning_messages");
                frmMain.g_oTables.m_oProcessorScenarioRun.CreateFRCSWarningTable(m_oAdo, m_oAdo.m_OleDbConnection, "FRCS_warning_messages");
                m_oAdo.m_strSQL = "INSERT INTO FRCS_warning_messages (biosum_cond_id,rxpackage,rx,rxcycle) " +
                                     "SELECT MID(STAND,1,25) AS biosum_cond_id," +
                                            "MID(rxpackage_rx_rxcycle,1,3) AS RXPackage," +
                                            "MID(rxpackage_rx_rxcycle,4,3) AS RX," +
                                            "MID(rxpackage_rx_rxcycle,7,1) AS RXCycle " +
                                     "FROM FRCS_output ";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }

            if (m_oAdo.m_intError == 0)
            {
                //
                //LOW SLOPE
                //
                if (ScenarioHarvestMethodVariables.ProcessLowSlope)
                {
                    //get the low slope harvest methods used for this batch
                    m_oAdo.m_strSQL = "SELECT DISTINCT i.[Harvesting System] " +
                                        "FROM frcs_input i WHERE i.[Percent Slope] <= " +
                                        ScenarioHarvestMethodVariables.SteepSlope.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    strHarvestMethodsList = m_oAdo.CreateCommaDelimitedList(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, "");
                    if (strHarvestMethodsList.Trim().Length > 0)
                    {
                        strHarvestMethodsArray = frmMain.g_oUtils.ConvertListToArray(strHarvestMethodsList, ",");
                        // Delete code; obsolete; FRCS no longer used
                    }
                }
            }
            if (m_oAdo.m_intError == 0)
            {
                //
                //STEEP SLOPE
                //
                if (ScenarioHarvestMethodVariables.ProcessSteepSlope)
                {

                    //get the steep slope harvest methods used for this batch
                    m_oAdo.m_strSQL = "SELECT DISTINCT i.[Harvesting System] " +
                                        "FROM frcs_input i WHERE i.[Percent Slope] > " +
                                        ScenarioHarvestMethodVariables.SteepSlope.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    strHarvestMethodsList = m_oAdo.CreateCommaDelimitedList(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, "");
                    if (strHarvestMethodsList.Trim().Length > 0)
                    {
                        strHarvestMethodsArray = frmMain.g_oUtils.ConvertListToArray(strHarvestMethodsList, ",");
                        // Delete code; No longer using FRCS
                    }
                }
            }
            //
            //update the WARNINGS column
            //
            m_oAdo.m_strSQL = "UPDATE FRCS_warning_messages SET warning = " +
                "IIF(lglogpa IS NOT NULL AND LEN(TRIM(lglogpa)) > 0,lglogpa + ' ','') + " +
                "IIF(lglogpatoallpa IS NOT NULL AND LEN(TRIM(lglogpatoallpa)) > 0,lglogpatoallpa + ' ','') + " +
                "IIF(chipvol IS NOT NULL AND LEN(TRIM(chipvol)) > 0,chipvol + ' ','') + " +
                "IIF(lglogvol IS NOT NULL AND LEN(TRIM(lglogvol)) > 0,lglogvol + ' ','') + " +
                "IIF(smlogvol IS NOT NULL AND LEN(TRIM(smlogvol)) > 0,smlogvol + ' ','') + " +
                "IIF(alllogvol IS NOT NULL AND LEN(TRIM(alllogvol)) > 0,alllogvol + ' ','') + " +
                "IIF(allvol IS NOT NULL AND LEN(TRIM(allvol)) > 0,allvol + ' ','') + " +
                "IIF(slope IS NOT NULL AND LEN(TRIM(slope)) > 0,slope + ' ','')";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;

            

        }
        private void RunScenario_AppendToHarvestCosts(string p_strHarvestCostTableName, bool bInactiveVarRxPackage)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_AppendToHarvestCosts\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                frmMain.g_oTables.m_oProcessor.CreateHarvestCostsTable(m_oAdo, m_oAdo.m_OleDbConnection, p_strHarvestCostTableName);
            }
            else
            {
                frmMain.g_oTables.m_oProcessor.CreateSqliteHarvestCostsTable(m_oDataMgr, m_oDataMgr.m_Connection, p_strHarvestCostTableName);
                if (!m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strHarvestCostTableName))
                {
                    dao_data_access oDao = new dao_data_access();
                    oDao.CreateSQLiteTableLink(m_oAdo.m_OleDbConnection.DataSource, p_strHarvestCostTableName,
                        p_strHarvestCostTableName, ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName,
                        m_oDataMgr.m_Connection.FileName);
                    oDao.m_DaoWorkspace.Close();
                    oDao = null;
                    do
                    {
                        // break out of loop if it runs too long
                        int i = 0;
                        if (i > 20)
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "An error occurred while trying to link to the " + p_strHarvestCostTableName + " table! \r\n");
                            }
                            break;
                        }
                        System.Threading.Thread.Sleep(1000);
                        i++;
                    }
                    while (!m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strHarvestCostTableName));
                }
            }

            if (bInactiveVarRxPackage)
            {
                // Don't append anything; There is nothing to append
                m_oAdo.m_strSQL = "";
            }
            else
            {
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.AppendToOPCOSTHarvestCostsTable(
                    "OpCost_Output", "OpCost_Input", p_strHarvestCostTableName, m_strDateTimeCreated);
                
            }

            if (m_oAdo.m_strSQL.Length > 0)
            {
                if (m_oAdo.m_intError == 0 && frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                if (m_oAdo.m_intError == 0) m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
        }
        private void RunScenario_UpdateHarvestCostsTableWithAdditionalCosts(string p_strHarvestCostsTableName, string p_strInactiveStandsTableName)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_UpdateHarvestCostsWithAdditionalCosts\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            int x, y;
            string strSum = "";

            string[] strRXArray = null;
            //create work table to hold total additional costs
            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "HarvestCostsTotalAdditionalWorkTable"))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE HarvestCostsTotalAdditionalWorkTable");
                frmMain.g_oTables.m_oProcessorScenarioRun.CreateTotalAdditionalHarvestCostsTable(
                    m_oAdo, m_oAdo.m_OleDbConnection, "HarvestCostsTotalAdditionalWorkTable");
                if (m_oAdo.m_intError == 0)
                {
                    //insert plot+rx records for the current scenario
                    m_oAdo.m_strSQL = "INSERT INTO HarvestCostsTotalAdditionalWorkTable " +
                                      "(biosum_cond_id,rx) SELECT biosum_cond_id,rx " +
                                                          "FROM scenario_additional_harvest_costs " +
                                                          "WHERE TRIM(UCASE(scenario_id)) = " +
                                                          "'" + ScenarioId.ToUpper().Trim() + "'";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                }

            if (m_oAdo.m_intError == 0)
            {
                //
                //GET THE ADDITIONAL HARVEST COST COLUMNS AND THEIR ASSOCIATED TREATMENT
                //
                m_oAdo.m_strSQL = "SELECT columnname,rx " +
                    "FROM scenario_harvest_cost_columns " +
                    "WHERE trim(ucase(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

                if (m_oAdo.m_intError == 0)
                {
                    string strScenarioColumnNameList = "";
                    string strScenarioRxList = "";
                    string[] strScenarioColumnNameArray = null;
                    string[] strScenarioRxArray = null;
                    string strCol = "";
                    //strScenarioColumnNameList = "harvest_cpa,";
                    if (m_oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (m_oAdo.m_OleDbDataReader.Read())
                        {
                            strCol = "";
                            //make sure the row is not null values
                            if (m_oAdo.m_OleDbDataReader["ColumnName"] != System.DBNull.Value &&
                                m_oAdo.m_OleDbDataReader["ColumnName"].ToString().Trim().Length > 0)
                            {

                                strCol = m_oAdo.m_OleDbDataReader["ColumnName"].ToString().Trim();
                                strScenarioColumnNameList = strScenarioColumnNameList + strCol + ",";

                                if (m_oAdo.m_OleDbDataReader["rx"] != System.DBNull.Value &&
                                    m_oAdo.m_OleDbDataReader["rx"].ToString().Trim().Length > 0)
                                {
                                    strCol = m_oAdo.m_OleDbDataReader["rx"].ToString();
                                }
                                else
                                {
                                    strCol = "*";
                                }
                                strScenarioRxList = strScenarioRxList + strCol + ",";

                            }
                        }
                    }
                    m_oAdo.m_OleDbDataReader.Close();
                    if (strScenarioColumnNameList.Trim().Length > 0)
                    {
                        /*****************************************************************
                         *the key is that the strScenarioColumnNameArray 
                         *and the strScenarioRxArray match in length
                         *and values as found in the scenario_harvest_cost_columns table
                         *****************************************************************/

                        strScenarioColumnNameList = strScenarioColumnNameList.Substring(0, strScenarioColumnNameList.Length - 1);
                        strScenarioColumnNameArray = frmMain.g_oUtils.ConvertListToArray(strScenarioColumnNameList, ",");
                        strScenarioRxList = strScenarioRxList.Substring(0, strScenarioRxList.Length - 1);
                        strScenarioRxArray = frmMain.g_oUtils.ConvertListToArray(strScenarioRxList, ",");

                        //update by rx that have both unique and global costs
                        m_oAdo.m_strSQL = "SELECT DISTINCT rx " +
                           "FROM scenario_harvest_cost_columns " +
                           "WHERE trim(ucase(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        strRXArray = frmMain.g_oUtils.ConvertListToArray(
                             (string)m_oAdo.CreateCommaDelimitedList(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, ""), ",");
                    }
                    //
                    //WITH TREATMENTS:
                    //PROCESS SCENARIO_HARVEST_COST_COLUMNS THAT HAVE ASSOCIATED TREATMENT
                    //
                    if (strRXArray != null)
                    {
                        for (x = 0; x <= strRXArray.Length - 1; x++)
                        {
                            strSum = "";
                            for (y = 0; y <= strScenarioRxArray.Length - 1; y++)
                            {
                                if (strScenarioRxArray[y].Trim().Length > 0 && strScenarioRxArray[y] != "*")
                                {
                                    if (strRXArray[x] == strScenarioRxArray[y])
                                    {
                                        //rx harvest cost
                                        strSum = strSum + "IIF(b." + strScenarioColumnNameArray[y].Trim() + " IS NOT NULL,b." + strScenarioColumnNameArray[y].Trim() + ",0) +";

                                    }
                                }
                                else
                                {
                                    //scenario harvest cost
                                    strSum = strSum + "IIF(b." + strScenarioColumnNameArray[y].Trim() + " IS NOT NULL,b." + strScenarioColumnNameArray[y].Trim() + ",0) +";
                                }
                            }
                            strSum = strSum.Substring(0, strSum.Length - 1);
                            m_oAdo.m_strSQL = "UPDATE HarvestCostsTotalAdditionalWorkTable a " +
                                              "INNER JOIN scenario_additional_harvest_costs b " +
                                              "ON a.biosum_cond_id=b.biosum_cond_id AND " +
                                                 "a.RX = b.RX " +
                                              "SET a.complete_additional_cpa = " + strSum + " " +
                                              "WHERE TRIM(UCASE(b.scenario_id))='" + ScenarioId.ToUpper().Trim() + "' AND " +
                                                     "b.RX = '" + strRXArray[x] + "'";

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

                        }
                    }
                    if (m_oAdo.m_intError == 0)
                    {
                        //
                        //WITHOUT TREATMENTS:
                        //PROCESS TREATMENTS THAT ARE IN THE SCENARIO_ADDITIONAL_HARVEST_COSTS table
                        //BUT ARE NOT IN THE SCENARIO_ADDITIONAL_HARVEST_COST_COLUMNS table
                        //
                        m_oAdo.m_strSQL = "SELECT DISTINCT  a.rx " +
                                          "FROM scenario_additional_harvest_costs a " +
                                          "WHERE NOT EXISTS (SELECT b.rx " +
                                                            "FROM scenario_harvest_cost_columns b " +
                                                            "WHERE b.rx=a.rx AND TRIM(UCASE(scenario_id))='" + ScenarioId.Trim().ToUpper() + "')";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        strRXArray = frmMain.g_oUtils.ConvertListToArray(
                                   (string)m_oAdo.CreateCommaDelimitedList(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, ""), ",");

                        if (strRXArray != null && strScenarioRxArray != null)
                        {

                            for (x = 0; x <= strRXArray.Length - 1; x++)
                            {
                                strSum = "";

                                for (y = 0; y <= strScenarioRxArray.Length - 1; y++)
                                {
                                    //check 
                                    if (strScenarioRxArray[y].Trim().Length > 0 && strScenarioRxArray[y] != "*")
                                    {

                                    }
                                    else
                                    {
                                        //scenario harvest cost
                                        strSum = strSum + "IIF(b." + strScenarioColumnNameArray[y].Trim() + " IS NOT NULL,b." + strScenarioColumnNameArray[y].Trim() + ",0) +";
                                    }
                                }


                                if (strSum.Trim().Length > 0)
                                {
                                    strSum = strSum.Substring(0, strSum.Length - 1);
                                    m_oAdo.m_strSQL = "UPDATE HarvestCostsTotalAdditionalWorkTable a " +
                                                      "INNER JOIN scenario_additional_harvest_costs b " +
                                                      "ON a.biosum_cond_id=b.biosum_cond_id AND " +
                                                         "a.RX = b.RX " +
                                                      "SET a.complete_additional_cpa = " + strSum + " " +
                                                      "WHERE TRIM(UCASE(b.scenario_id))='" + ScenarioId.ToUpper().Trim() + "' AND " +
                                                             "b.RX = '" + strRXArray[x] + "'";

                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                                    if (m_oAdo.m_intError != 0) break;
                                }

                            }
                        }
                    }
                    if (m_oAdo.m_intError == 0)
                    {
                        m_oAdo.m_strSQL = "UPDATE HarvestCostsTotalAdditionalWorkTable " +
                                          "SET complete_additional_cpa = 0 " +
                                          "WHERE complete_additional_cpa IS NULL";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    }
                    if (m_oAdo.m_intError == 0)
                    {
                        //update the harvest costs table complete costs per acre column
                        m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.UpdateHarvestCostsTableWithCompleteCostsPerAcre(
                                             "scenario_cost_revenue_escalators",
                                             "HarvestCostsTotalAdditionalWorkTable",
                                             p_strHarvestCostsTableName, ScenarioId, false, false);

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    }
                    if (m_oAdo.m_intError == 0 && m_bIncludeInactiveStands == true)
                    {                        
                        //update the inactive stands table complete costs per acre column
                        m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.UpdateHarvestCostsTableWithCompleteCostsPerAcre(
                                             "scenario_cost_revenue_escalators",
                                             "HarvestCostsTotalAdditionalWorkTable",
                                             p_strInactiveStandsTableName, ScenarioId, true, false);

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    }
                    if (m_oAdo.m_intError == 0 && m_bIncludeInactiveStands == true)
                    {
                        //append the inactive stands to the harvest cost table where complete_cpa > 0
                        m_oAdo.m_strSQL = "INSERT INTO " + p_strHarvestCostsTableName +
                                          " SELECT biosum_cond_id, rxpackage, rx, rxcycle, harvest_cpa, complete_cpa, '" +
                                          m_strDateTimeCreated + "' as DateTimeCreated, 0 as chip_cpa, 0 as assumed_movein_cpa, " +
                                          "0 as ideal_complete_cpa, 0 as ideal_harvest_cpa, 0 as ideal_chip_cpa, 0 as ideal_assumed_movein_cpa" +
                                          " FROM " + p_strInactiveStandsTableName +
                                          " WHERE complete_cpa > 0";

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    }
                    //if (m_oAdo.m_intError == 0)
                    //{
                    //    //update the harvest costs table ideal complete costs per acre column
                    //    m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.UpdateHarvestCostsTableWithIdealCompleteCostsPerAcre(
                    //                         "scenario_cost_revenue_escalators",
                    //                         "HarvestCostsTotalAdditionalWorkTable",
                    //                         p_strHarvestCostsTableName, ScenarioId);
                    //    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    //        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    //    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    //}
                }
            }

            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
            }
            else
            {
                RunScenario_UpdateHarvestCostsTableWithAdditionalCostsSqlite(p_strHarvestCostsTableName, p_strInactiveStandsTableName);
            }
        }

        private void RunScenario_UpdateHarvestCostsTableWithAdditionalCostsSqlite(string p_strHarvestCostsTableName, 
            string p_strInactiveStandsTableName)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_UpdateHarvestCostsWithAdditionalCostsSqlite\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            int x, y;
            string strSum = "";

            //create work table to hold total additional costs
            if (m_oDataMgr.TableExist(m_oDataMgr.m_Connection, "HarvestCostsTotalAdditionalWorkTable"))
                m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, "DROP TABLE HarvestCostsTotalAdditionalWorkTable");
            frmMain.g_oTables.m_oProcessorScenarioRun.CreateSqliteTotalAdditionalHarvestCostsTable(
                m_oDataMgr, m_oDataMgr.m_Connection, "HarvestCostsTotalAdditionalWorkTable");
            if (m_oDataMgr.m_intError == 0)
            {
                //insert plot+rx records for the current scenario
                m_oDataMgr.m_strSQL = "INSERT INTO HarvestCostsTotalAdditionalWorkTable " +
                                      "(biosum_cond_id,rx) SELECT biosum_cond_id,rx " +
                                                      "FROM definitions.scenario_additional_harvest_costs " +
                                                      "WHERE TRIM(scenario_id) = " +
                                                      "'" + ScenarioId.Trim() + "'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
            }

            if (m_oDataMgr.m_intError == 0)
            {
                //
                //GET THE ADDITIONAL HARVEST COST COLUMNS AND THEIR ASSOCIATED TREATMENT
                //
                m_oDataMgr.m_strSQL = "SELECT columnname,rx " +
                    "FROM definitions.scenario_harvest_cost_columns " +
                    "WHERE trim(scenario_id)='" + ScenarioId.Trim() + "'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlQueryReader(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);

                if (m_oDataMgr.m_intError == 0)
                {
                    string strScenarioColumnNameList = "";
                    string strScenarioRxList = "";
                    string[] strScenarioColumnNameArray = null;
                    string[] strScenarioRxArray = null;
                    string strCol = "";
                    if (m_oDataMgr.m_DataReader.HasRows)
                    {
                        while (m_oDataMgr.m_DataReader.Read())
                        {
                            strCol = "";
                            //make sure the row is not null values
                            if (m_oDataMgr.m_DataReader["ColumnName"] != System.DBNull.Value &&
                                m_oDataMgr.m_DataReader["ColumnName"].ToString().Trim().Length > 0)
                            {

                                strCol = m_oDataMgr.m_DataReader["ColumnName"].ToString().Trim();
                                strScenarioColumnNameList = strScenarioColumnNameList + strCol + ",";

                                if (m_oDataMgr.m_DataReader["rx"] != System.DBNull.Value &&
                                    m_oDataMgr.m_DataReader["rx"].ToString().Trim().Length > 0)
                                {
                                    strCol = m_oDataMgr.m_DataReader["rx"].ToString();
                                }
                                else
                                {
                                    strCol = "*";
                                }
                                strScenarioRxList = strScenarioRxList + strCol + ",";

                            }
                        }
                    }
                    m_oDataMgr.m_DataReader.Close();
                    IList<string> lstRx = null;
                    if (strScenarioColumnNameList.Trim().Length > 0)
                    {
                        /*****************************************************************
                         *the key is that the strScenarioColumnNameArray 
                         *and the strScenarioRxArray match in length
                         *and values as found in the scenario_harvest_cost_columns table
                         *****************************************************************/

                        strScenarioColumnNameList = strScenarioColumnNameList.Substring(0, strScenarioColumnNameList.Length - 1);
                        strScenarioColumnNameArray = frmMain.g_oUtils.ConvertListToArray(strScenarioColumnNameList, ",");
                        strScenarioRxList = strScenarioRxList.Substring(0, strScenarioRxList.Length - 1);
                        strScenarioRxArray = frmMain.g_oUtils.ConvertListToArray(strScenarioRxList, ",");

                        //update by rx that have both unique and global costs
                        m_oDataMgr.m_strSQL = "SELECT DISTINCT rx " +
                           "FROM definitions.scenario_harvest_cost_columns " +
                           "WHERE trim(scenario_id)='" + ScenarioId.Trim() + "'";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        lstRx = m_oDataMgr.getStringList(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                    }
                    //
                    //WITH TREATMENTS:
                    //PROCESS SCENARIO_HARVEST_COST_COLUMNS THAT HAVE ASSOCIATED TREATMENT
                    //
                    if (lstRx != null)
                    {
                        for (x = 0; x <= lstRx.Count - 1; x++)
                        {
                            strSum = "";
                            for (y = 0; y <= strScenarioRxArray.Length - 1; y++)
                            {
                                if (strScenarioRxArray[y].Trim().Length > 0 && strScenarioRxArray[y] != "*")
                                {
                                    if (lstRx[x] == strScenarioRxArray[y])
                                    {
                                        //rx harvest cost
                                        strSum = strSum + "COALESCE(b." + strScenarioColumnNameArray[y].Trim() + ", 0) + ";
                                    }
                                }
                                else
                                {
                                    //scenario harvest cost
                                    strSum = strSum + "COALESCE(b." + strScenarioColumnNameArray[y].Trim() + ", 0) + ";
                                }
                            }
                            strSum = strSum.Substring(0, strSum.Length - 2);
                            m_oDataMgr.m_strSQL = "UPDATE HarvestCostsTotalAdditionalWorkTable AS a " +
                                              "SET complete_additional_cpa = (SELECT " + strSum + " " +
                                              "FROM definitions.scenario_additional_harvest_costs as b " +
                                              "WHERE a.biosum_cond_id = b.biosum_cond_id " +
                                              "AND a.RX = b.RX " +
                                              "AND TRIM(b.scenario_id) = '" + ScenarioId.Trim() + "' " + 
                                              "AND b.RX = '" + lstRx[x] + "') " +
                                              "WHERE EXISTS(" +
                                              "SELECT * " +
                                              "FROM definitions.scenario_additional_harvest_costs AS b " +
                                              "WHERE a.biosum_cond_id = b.biosum_cond_id " +
                                              "AND a.RX = b.RX AND " +
                                              "TRIM(b.scenario_id)='" + ScenarioId.Trim() + "' AND " +
                                              "b.RX = '" + lstRx[x] + "')";

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);

                        }
                    }
                    if (m_oDataMgr.m_intError == 0)
                    {
                        //
                        //WITHOUT TREATMENTS:
                        //PROCESS TREATMENTS THAT ARE IN THE SCENARIO_ADDITIONAL_HARVEST_COSTS table
                        //BUT ARE NOT IN THE SCENARIO_ADDITIONAL_HARVEST_COST_COLUMNS table
                        //
                        m_oDataMgr.m_strSQL = "SELECT DISTINCT  a.rx " +
                                          "FROM scenario_additional_harvest_costs a " +
                                          "WHERE NOT EXISTS (SELECT b.rx " +
                                                            "FROM scenario_harvest_cost_columns b " +
                                                            "WHERE b.rx=a.rx AND TRIM(scenario_id)='" + ScenarioId.Trim() + "')";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        lstRx = m_oDataMgr.getStringList(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);


                        if (lstRx != null && strScenarioRxArray != null)
                        {

                            for (x = 0; x <= lstRx.Count - 1; x++)
                            {
                                strSum = "";

                                for (y = 0; y <= strScenarioRxArray.Length - 1; y++)
                                {
                                    //check 
                                    if (strScenarioRxArray[y].Trim().Length > 0 && strScenarioRxArray[y] != "*")
                                    {

                                    }
                                    else
                                    {
                                        //scenario harvest cost
                                        strSum = strSum + "COALESCE(b." + strScenarioColumnNameArray[y].Trim() + ", 0) + ";
                                    }
                                }


                                if (strSum.Trim().Length > 0)
                                {
                                    strSum = strSum.Substring(0, strSum.Length - 2);
                                    m_oDataMgr.m_strSQL = "UPDATE HarvestCostsTotalAdditionalWorkTable AS a " +
                                                      "SET complete_additional_cpa = (SELECT " + strSum + " " +
                                                      "FROM definitions.scenario_additional_harvest_costs as b " +
                                                      "WHERE a.biosum_cond_id = b.biosum_cond_id " +
                                                      "AND a.RX = b.RX " +
                                                      "AND TRIM(b.scenario_id) = '" + ScenarioId.Trim() + "' " +
                                                      "AND b.RX = '" + lstRx[x] + "') " +
                                                      "WHERE EXISTS(" +
                                                      "SELECT * " +
                                                      "FROM definitions.scenario_additional_harvest_costs AS b " +
                                                      "WHERE a.biosum_cond_id = b.biosum_cond_id " +
                                                      "AND a.RX = b.RX AND " +
                                                      "TRIM(b.scenario_id)='" + ScenarioId.Trim() + "' AND " +
                                                      "b.RX = '" + lstRx[x] + "')";

                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                    m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                                    if (m_oDataMgr.m_intError != 0) break;
                                }

                            }
                        }
                    }
                    if (m_oDataMgr.m_intError == 0)
                    {
                        m_oDataMgr.m_strSQL = "UPDATE HarvestCostsTotalAdditionalWorkTable " +
                                          "SET complete_additional_cpa = 0 " +
                                          "WHERE complete_additional_cpa IS NULL";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                    }
                    if (m_oDataMgr.m_intError == 0)
                    {
                        //update the harvest costs table complete costs per acre column
                        m_oDataMgr.m_strSQL = Queries.ProcessorScenarioRun.UpdateHarvestCostsTableWithCompleteCostsPerAcre(
                                             "scenario_cost_revenue_escalators",
                                             "HarvestCostsTotalAdditionalWorkTable",
                                             p_strHarvestCostsTableName, ScenarioId, false, true);

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                    }
                    if (m_oDataMgr.m_intError == 0 && m_bIncludeInactiveStands == true)
                    {
                        //update the inactive stands table complete costs per acre column
                        m_oDataMgr.m_strSQL = Queries.ProcessorScenarioRun.UpdateHarvestCostsTableWithCompleteCostsPerAcre(
                                             "scenario_cost_revenue_escalators",
                                             "HarvestCostsTotalAdditionalWorkTable",
                                             p_strInactiveStandsTableName, ScenarioId, true, true);

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                    }
                    if (m_oDataMgr.m_intError == 0 && m_bIncludeInactiveStands == true)
                    {
                        //append the inactive stands to the harvest cost table where complete_cpa > 0
                        m_oDataMgr.m_strSQL = "INSERT INTO " + p_strHarvestCostsTableName +
                                          " (biosum_cond_id, rxpackage, rx, rxcycle, harvest_cpa, complete_cpa,  DateTimeCreated, " +
                                          "chip_cpa, assumed_movein_cpa, ideal_complete_cpa, ideal_harvest_cpa,ideal_chip_cpa,ideal_assumed_movein_cpa) " +
                                          " SELECT biosum_cond_id, rxpackage, rx, rxcycle, harvest_cpa, complete_cpa, '" +
                                          m_strDateTimeCreated + "' as DateTimeCreated, 0 as chip_cpa, 0 as assumed_movein_cpa, " +
                                          "0 as ideal_complete_cpa, 0 as ideal_harvest_cpa, 0 as ideal_chip_cpa, 0 as ideal_assumed_movein_cpa" +
                                          " FROM " + p_strInactiveStandsTableName +
                                          " WHERE complete_cpa > 0";

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                    }
                }
            }

            m_intError = m_oDataMgr.m_intError;
            m_strError = m_oDataMgr.m_strError;
        }

        private void RunScenario_CreateInactiveStandsHarvestCostsWorkTable(string p_strHarvestCostsTableName, string p_strInactiveStandsTableName,
            string p_strVariant, string p_strRxPackage)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_CreateInactiveStandsHarvestCostsWorkTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            if (m_oAdo.ColumnExist(m_oAdo.m_OleDbConnection, Tables.FVS.DefaultPreFVSComputeTableName, "ACTIVITY"))
            {
                m_oAdo.m_strSQL = "select count (*) from " + Tables.FVS.DefaultPreFVSComputeTableName +
                    " where ACTIVITY = 1";
                int intInactiveCount = m_oAdo.getRecordCount(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, Tables.FVS.DefaultPreFVSComputeTableName);
                if (intInactiveCount < 1)
                {
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "//No inactive stands marked in PRE_FVS_COMPUTE table! " +
                            "Nothing to append \r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                    }
                    m_bIncludeInactiveStands = false;
                    return;
                }
            }
            else
            {
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//ACTIVITY_1  column is missing from PRE_FVS_COMPUTE table. " +
                        "Stands with no activity will not be included in analysis! \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                }
                m_bIncludeInactiveStands = false;
                return;
            }

            string strFvsComputeTable = Tables.FVS.DefaultPreFVSComputeTableName;
            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                m_oAdo.m_strSQL = "SELECT " + strFvsComputeTable + ".biosum_cond_id, " + strFvsComputeTable + ".rxpackage, " +
                  strFvsComputeTable + ".rx," + strFvsComputeTable + ".rxcycle, " + strFvsComputeTable +
                  ".fvs_variant, 0 as harvest_cpa, 0 as complete_cpa" +
                  " INTO " + p_strInactiveStandsTableName +
                  " from " + strFvsComputeTable +
                  " left outer join " + p_strHarvestCostsTableName + " on " +
                  p_strHarvestCostsTableName + ".biosum_cond_id=" + strFvsComputeTable + ".biosum_cond_id and " +
                  p_strHarvestCostsTableName + ".rxpackage=" + strFvsComputeTable + ".rxpackage and " +
                  p_strHarvestCostsTableName + ".rx=" + strFvsComputeTable + ".rx and " +
                  p_strHarvestCostsTableName + ".rxcycle=" + strFvsComputeTable + ".rxcycle" +
                  " where " + p_strHarvestCostsTableName + ".biosum_cond_id is null" +
                  " and " + strFvsComputeTable + ".ACTIVITY  = 1 AND " +
                  strFvsComputeTable + ".FVS_VARIANT = '" + p_strVariant + "' AND " +
                  strFvsComputeTable + ".RXPACKAGE = '" + p_strRxPackage + "'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + m_oAdo.m_strSQL + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            else
            {
                string strSql = "CREATE TABLE " + p_strInactiveStandsTableName + " (" +
                    "biosum_cond_id TEXT," +
                    "rxpackage TEXT," +
                    "rx TEXT," +
                    "rxcycle TEXT," +
                    "fvs_variant TEXT," +
                    "complete_cpa REAL DEFAULT 0," +
                    "harvest_cpa REAL DEFAULT 0)";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strSql + "\r\n");
                m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, strSql);

                if (!m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strInactiveStandsTableName))
                {
                    dao_data_access oDao = new dao_data_access();
                    oDao.CreateSQLiteTableLink(m_oAdo.m_OleDbConnection.DataSource, p_strInactiveStandsTableName,
                        p_strInactiveStandsTableName, ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName,
                        m_oDataMgr.m_Connection.FileName);
                    oDao.m_DaoWorkspace.Close();
                    oDao = null;
                    do
                    {
                        // break out of loop if it runs too long
                        int i = 0;
                        if (i > 20)
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "An error occurred while trying to link to the " + p_strInactiveStandsTableName + " table! \r\n");
                            }
                            break;
                        }
                        System.Threading.Thread.Sleep(1000);
                        i++;
                    }
                    while (!m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strInactiveStandsTableName));
                }

                strSql = "INSERT INTO " + p_strInactiveStandsTableName +
                         " SELECT " + strFvsComputeTable + ".biosum_cond_id, " + strFvsComputeTable + ".rxpackage, " +
                         strFvsComputeTable + ".rx," + strFvsComputeTable + ".rxcycle, " + strFvsComputeTable +
                         ".fvs_variant, 0 as harvest_cpa, 0 as complete_cpa" +
                         " FROM " + strFvsComputeTable +
                         " left outer join " + p_strHarvestCostsTableName +
                         " on trim(" + p_strHarvestCostsTableName + ".biosum_cond_id) = " + strFvsComputeTable + ".biosum_cond_id and" +
                         " trim(" + p_strHarvestCostsTableName + ".rxpackage) = " + strFvsComputeTable + ".rxpackage and" +
                         " trim(" + p_strHarvestCostsTableName + ".rx) = " + strFvsComputeTable + ".rx and" +
                         " trim(" + p_strHarvestCostsTableName + ".rxcycle) = " + strFvsComputeTable + ".rxcycle" +
                         " where " + p_strHarvestCostsTableName + ".biosum_cond_id is null" +
                         " and " + strFvsComputeTable + ".ACTIVITY = 1 AND " +
                         strFvsComputeTable + ".FVS_VARIANT = '" + p_strVariant + "' AND " +
                         strFvsComputeTable + ".RXPACKAGE = '" + p_strRxPackage + "'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strSql + "\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, strSql);

            }

            if (m_oAdo.m_intError != 0)
            {
                m_bIncludeInactiveStands = false;
            }

            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
        }


        private void RunScenario_DeleteFromTreeVolValAndHarvestCostsTable(string p_strVariant, string p_strRxPackage)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_DeleteFromTreeVolValAndHarvestCostsTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            //
            //DELETE THE VARIANT+RXPACKAGE FROM THE TREEVOLVAL TABLE
            //
            //@ToDo: This has to be done on the Access side because of references to cond and plot table
            m_oAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " t " +
                              "WHERE EXISTS (SELECT c.biosum_cond_id,p.fvs_variant " +
                                            "FROM " + this.m_oQueries.m_oFIAPlot.m_strCondTable + " c," +
                                                      this.m_oQueries.m_oFIAPlot.m_strPlotTable + " p " +
                                            "WHERE t.rxpackage='" + p_strRxPackage + "' AND " +
                                                  "t.biosum_cond_id=c.biosum_cond_id AND " +
                                                  "p.biosum_plot_id=c.biosum_plot_id AND " +
                                                  "p.fvs_variant='" + p_strVariant + "')";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            if (m_oAdo.m_intError == 0)
            {                                  
                //
                //DELETE THE VARIANT+RXPACKAGE FROM THE HARVEST COSTS TABLE
                //
                m_oAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " t " +
                                  "WHERE EXISTS (SELECT c.biosum_cond_id,p.fvs_variant " +
                                                "FROM " + this.m_oQueries.m_oFIAPlot.m_strCondTable + " c," +
                                                          this.m_oQueries.m_oFIAPlot.m_strPlotTable + " p " +
                                                "WHERE trim(t.rxpackage)='" + p_strRxPackage + "' AND " +
                                                     "(trim(t.biosum_cond_id)=c.biosum_cond_id AND " +
                                                      "p.biosum_plot_id=c.biosum_plot_id AND " +
                                                      "p.fvs_variant='" + p_strVariant + "'))";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;

        }
        private void RunScenario_AppendToTreeVolValAndHarvestCostsTable()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_AppendDataToTreeVolVal\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "TreeVolValLowSlope"))
                {
                    //
                    //APPEND TO SCENARIO TREE VOL VAL TABLE
                    //
                    m_oAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " " +
                                      "SELECT * FROM TreeVolValLowSlope";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                }
                if (m_oAdo.m_intError == 0 && m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "HarvestCostsWorkTable"))
                {
                    //
                    //APPEND TO SCENARIO HARVEST COST TABLE
                    //
                    m_oAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " " +
                                      "SELECT * FROM HarvestCostsWorkTable";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                }
                m_intError = m_oAdo.m_intError;
                m_strError = m_oAdo.m_strError;
            }
            else
            {
                if (m_oDataMgr.TableExist(m_oDataMgr.m_Connection, "TreeVolValLowSlope"))
                {
                    //
                    //APPEND TO SCENARIO TREE VOL VAL TABLE
                    //
                    m_oDataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " (biosum_cond_id, rxpackage,"+
                                           "rx, rxcycle, species_group, diam_group, chip_vol_cf, chip_wt_gt, " +
                                           "chip_val_dpa, chip_mkt_val_pgt, merch_vol_cf, merch_wt_gt, merch_val_dpa, " +
                                           "merch_to_chipbin_YN, bc_vol_cf, bc_wt_gt, stand_residue_wt_gt, place_holder, DateTimeCreated) " +
                                           "SELECT biosum_cond_id, rxpackage, " +
                                           "rx, rxcycle, species_group, diam_group, chip_vol_cf, chip_wt_gt, " +
                                           "chip_val_dpa, chip_mkt_val_pgt, merch_vol_cf, merch_wt_gt, merch_val_dpa, " +
                                           "merch_to_chipbin_YN, bc_vol_cf, bc_wt_gt, stand_residue_wt_gt, place_holder, DateTimeCreated " +
                                           "FROM TreeVolValLowSlope";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                }
                if (m_oDataMgr.m_intError == 0 && m_oDataMgr.TableExist(m_oDataMgr.m_Connection, "HarvestCostsWorkTable"))
                {
                    //
                    //APPEND TO SCENARIO HARVEST COST TABLE
                    //
                    m_oDataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " " +
                                      "SELECT * FROM HarvestCostsWorkTable";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                }
                m_intError = m_oDataMgr.m_intError;
                m_strError = m_oDataMgr.m_strError;
            }
        }

        private void RunScenario_Finished()
        {
            //if (m_oExcel.ExcelFileName.Trim().Length > 0) System.IO.File.Delete(m_oExcel.ExcelFileName);
            if (m_oExcel != null)
            {
                m_oExcel.ReleaseComObjects();
                m_oExcel = null;
            }
            uc_filesize_monitor1.EndMonitoringFile();
            uc_filesize_monitor2.EndMonitoringFile();
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceProcessorScenarioForm.tlbScenario, "Enabled", true);
            ReferenceProcessorScenarioForm.EnableTabPage(ReferenceProcessorScenarioForm.tabControlScenario, "tbDataSources", true);
            ReferenceProcessorScenarioForm.EnableTabPage(ReferenceProcessorScenarioForm.tabControlScenario, "tbTreeGroupings", true);
            ReferenceProcessorScenarioForm.EnableTabPage(ReferenceProcessorScenarioForm.tabControlRules, "tbHarvestMethod,tbWoodValue,tbEscalators,tbAddHarvestCosts,tbFilterCond,tbMoveInCosts", true);
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceProcessorScenarioForm.tlbScenario,"Enabled",true);
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceProcessorScenarioForm.tabControlRules, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceProcessorScenarioForm.tabControlScenario, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)btnChkAll, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)btnUncheckAll, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)btnRun, "Text", "Run");
            frmMain.g_oDelegate.SetStatusBarPanelTextValue((System.Windows.Forms.StatusBar)frmMain.g_sbpInfo.Parent, 1, "Ready");
            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "* = New FVS Tree records to process");
            //frmMain.g_oDelegate.ExecuteControlMethod(lblMsg, "Hide");
            frmMain.g_oDelegate.ExecuteControlMethod(this, "Refresh");
            frmMain.g_oDelegate.CurrentThreadProcessIdle = true;
            

        }

        private void chkLowSlope_CheckedChanged(object sender, EventArgs e)
        {
            if (ReferenceProcessorScenarioForm.m_bDataSourceFirstTime == false) ReferenceProcessorScenarioForm.m_bSave = true;
            ScenarioHarvestMethodVariables.ProcessLowSlope = m_blnLowSlope;
        }

        private void chkSteepSlope_CheckedChanged(object sender, EventArgs e)
        {
            if (ReferenceProcessorScenarioForm.m_bDataSourceFirstTime == false) ReferenceProcessorScenarioForm.m_bSave = true;
            ScenarioHarvestMethodVariables.ProcessSteepSlope = m_blnSteepSlope;
        }

        private void RunScenario_AppendPlaceholdersToTreeVolValAndHarvestCostsTables()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_AppendPlaceholdersToTreeVolValTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                // TREE VOL VAL table
                if (m_oAdo.m_intError == 0)
                {
                    // Query the conditions/rxpackage that have records in cycles 2,3, and 4 but not in cycle 1
                    m_oAdo.m_strSQL = "SELECT t.biosum_cond_id, t.rxpackage, t.rx " +
                        "FROM " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " t " +
                        "WHERE t.rxcycle in ('2','3','4') " +
                        "AND NOT EXISTS (" +
                        "SELECT t1.biosum_cond_id, t1.rxpackage " +
                        "FROM " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " t1 " +
                        "WHERE t1.rxcycle = '1' " +
                        "AND t.biosum_cond_id = t1.biosum_cond_id " +
                        "AND t.rxpackage = t1.rxpackage) " +
                        "GROUP BY t.biosum_cond_id, t.rxpackage, t.rx";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "EXECUTE SQL: " + m_oAdo.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                    m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");

                    if (m_oAdo.m_OleDbDataReader.HasRows)
                    {
                        long lngCount = 0;
                        string strRxCycle = "1";
                        int intGroupPlaceholder = 999;
                        int intValuePlaceholder = 0;
                        //For each condition id/rxPackage combination returned by the query above
                        while (m_oAdo.m_OleDbDataReader.Read())
                        {
                            string cond_id = "";
                            string rxpackage = "";
                            string rx = "";
                            if (m_oAdo.m_OleDbDataReader["biosum_cond_id"] != System.DBNull.Value)
                                cond_id = m_oAdo.m_OleDbDataReader["biosum_cond_id"].ToString().Trim();
                            if (m_oAdo.m_OleDbDataReader["rxpackage"] != System.DBNull.Value)
                                rxpackage = m_oAdo.m_OleDbDataReader["rxpackage"].ToString().Trim();
                            if (m_oAdo.m_OleDbDataReader["rx"] != System.DBNull.Value)
                                rx = m_oAdo.m_OleDbDataReader["rx"].ToString().Trim();

                            //Insert a placeholder row with default values
                            m_oAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " " +
                                "(biosum_cond_id, rxpackage, rx, rxcycle, species_group, diam_group, " +
                                "merch_wt_gt, merch_val_dpa, merch_vol_cf, merch_to_chipbin_YN, " +
                                "chip_wt_gt, chip_val_dpa, chip_vol_cf, bc_vol_cf, bc_wt_gt, " +
                                "DateTimeCreated, place_holder) " +
                                "VALUES ('" + cond_id + "', '" + rxpackage + "', '" + rx + "', '" + strRxCycle + "', " +
                                intGroupPlaceholder + ", " + intGroupPlaceholder + ", " +
                                intValuePlaceholder + ", " + intValuePlaceholder + ", " + intValuePlaceholder + ", 'N', " +
                                intValuePlaceholder + ", " + intValuePlaceholder + ", " + intValuePlaceholder + ", " + intValuePlaceholder + ", " + intValuePlaceholder +
                                ", '" + m_strDateTimeCreated + "', 'Y')";

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n INSERT RECORD: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                            if (m_oAdo.m_intError != 0) break;
                            lngCount++;
                            //Console.WriteLine("Condition -> " + cond_id);
                        }
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, " \r\n END INSERTED " + lngCount + " RECORDS: " + System.DateTime.Now.ToString() + "\r\n");
                    }
                }

                // HARVEST COSTS table
                if (m_oAdo.m_intError == 0)
                {
                    // Query the conditions/rxpackage that have records in cycles 2,3, and 4 but not in cycle 1
                    m_oAdo.m_strSQL = "SELECT t.biosum_cond_id, t.rxpackage, t.rx " +
                        "FROM " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " t " +
                        "WHERE t.rxcycle in ('2','3','4') " +
                        "AND NOT EXISTS (" +
                        "SELECT t1.biosum_cond_id, t1.rxpackage " +
                        "FROM " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " t1 " +
                        "WHERE t1.rxcycle = '1' " +
                        "AND t.biosum_cond_id = t1.biosum_cond_id " +
                        "AND t.rxpackage = t1.rxpackage) " +
                        "GROUP BY t.biosum_cond_id, t.rxpackage, t.rx";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "EXECUTE SQL: " + m_oAdo.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                    m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");

                    if (m_oAdo.m_OleDbDataReader.HasRows)
                    {
                        long lngCount = 0;
                        string strRxCycle = "1";
                        int intValuePlaceholder = 0;
                        //For each condition id/rxPackage combination returned by the query above
                        while (m_oAdo.m_OleDbDataReader.Read())
                        {
                            string cond_id = "";
                            string rxpackage = "";
                            string rx = "";
                            if (m_oAdo.m_OleDbDataReader["biosum_cond_id"] != System.DBNull.Value)
                                cond_id = m_oAdo.m_OleDbDataReader["biosum_cond_id"].ToString().Trim();
                            if (m_oAdo.m_OleDbDataReader["rxpackage"] != System.DBNull.Value)
                                rxpackage = m_oAdo.m_OleDbDataReader["rxpackage"].ToString().Trim();
                            if (m_oAdo.m_OleDbDataReader["rx"] != System.DBNull.Value)
                                rx = m_oAdo.m_OleDbDataReader["rx"].ToString().Trim();

                            //Insert a placeholder row with default values
                            m_oAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " " +
                                "(biosum_cond_id, rxpackage, rx, rxcycle, " +
                                "complete_cpa, harvest_cpa," +
                                "DateTimeCreated, place_holder) " +
                                "VALUES ('" + cond_id + "', '" + rxpackage + "', '" + rx + "', '" + strRxCycle + "', " +
                                intValuePlaceholder + ", " + intValuePlaceholder +
                                ", '" + m_strDateTimeCreated + "', 'Y')";

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n INSERT RECORD: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                            if (m_oAdo.m_intError != 0) break;
                            lngCount++;
                            //Console.WriteLine("Condition -> " + cond_id);
                        }
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, " \r\n END INSERTED " + lngCount + " RECORDS: " + System.DateTime.Now.ToString() + "\r\n");
                    }
                }


                m_intError = m_oAdo.m_intError;
                m_strError = m_oAdo.m_strError;
            }
            else
            {
                RunScenario_AppendPlaceholdersToSqliteTreeVolValAndHarvestCostsTables();
            }

        }

        private void RunScenario_AppendPlaceholdersToSqliteTreeVolValAndHarvestCostsTables()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "// RunScenario_AppendPlaceholdersToSqliteTreeVolValAndHarvestCostsTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            // TREE VOL VAL table
            if (m_oDataMgr.m_intError == 0)
            {
                // Query the conditions/rxpackage that have records in cycles 2,3, and 4 but not in cycle 1
                m_oDataMgr.m_strSQL = "SELECT t.biosum_cond_id, t.rxpackage, t.rx " +
                    "FROM " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " t " +
                    "WHERE t.rxcycle in ('2','3','4') " +
                    "AND NOT EXISTS (" +
                    "SELECT t1.biosum_cond_id, t1.rxpackage " +
                    "FROM " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " t1 " +
                    "WHERE t1.rxcycle = '1' " +
                    "AND t.biosum_cond_id = t1.biosum_cond_id " +
                    "AND t.rxpackage = t1.rxpackage) " +
                    "GROUP BY t.biosum_cond_id, t.rxpackage, t.rx";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlQueryReader(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");

                if (m_oDataMgr.m_DataReader.HasRows)
                {
                    long lngCount = 0;
                    string strRxCycle = "1";
                    int intGroupPlaceholder = 999;
                    int intValuePlaceholder = 0;
                    //For each condition id/rxPackage combination returned by the query above
                    while (m_oDataMgr.m_DataReader.Read())
                    {
                        string cond_id = "";
                        string rxpackage = "";
                        string rx = "";
                        if (m_oDataMgr.m_DataReader["biosum_cond_id"] != System.DBNull.Value)
                            cond_id = m_oDataMgr.m_DataReader["biosum_cond_id"].ToString().Trim();
                        if (m_oDataMgr.m_DataReader["rxpackage"] != System.DBNull.Value)
                            rxpackage = m_oDataMgr.m_DataReader["rxpackage"].ToString().Trim();
                        if (m_oDataMgr.m_DataReader["rx"] != System.DBNull.Value)
                            rx = m_oDataMgr.m_DataReader["rx"].ToString().Trim();

                        //Insert a placeholder row with default values
                        m_oDataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " " +
                            "(biosum_cond_id, rxpackage, rx, rxcycle, species_group, diam_group, " +
                            "merch_wt_gt, merch_val_dpa, merch_vol_cf, merch_to_chipbin_YN, " +
                            "chip_wt_gt, chip_val_dpa, chip_vol_cf, bc_vol_cf, bc_wt_gt, " +
                            "DateTimeCreated, place_holder) " +
                            "VALUES ('" + cond_id + "', '" + rxpackage + "', '" + rx + "', '" + strRxCycle + "', " +
                            intGroupPlaceholder + ", " + intGroupPlaceholder + ", " +
                            intValuePlaceholder + ", " + intValuePlaceholder + ", " + intValuePlaceholder + ", 'N', " +
                            intValuePlaceholder + ", " + intValuePlaceholder + ", " + intValuePlaceholder + ", " + intValuePlaceholder + ", " + intValuePlaceholder +
                            ", '" + m_strDateTimeCreated + "', 'Y')";

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n INSERT RECORD: " + System.DateTime.Now.ToString() + "\r\n");
                        m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                        if (m_oDataMgr.m_intError != 0) break;
                        lngCount++;
                    }
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, " \r\n END INSERTED " + lngCount + " RECORDS: " + System.DateTime.Now.ToString() + "\r\n");
                }
            }

            // HARVEST COSTS table
            if (m_oDataMgr.m_intError == 0)
            {
                // Query the conditions/rxpackage that have records in cycles 2,3, and 4 but not in cycle 1
                m_oDataMgr.m_strSQL = "SELECT t.biosum_cond_id, t.rxpackage, t.rx " +
                    "FROM " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " t " +
                    "WHERE t.rxcycle in ('2','3','4') " +
                    "AND NOT EXISTS (" +
                    "SELECT t1.biosum_cond_id, t1.rxpackage " +
                    "FROM " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " t1 " +
                    "WHERE t1.rxcycle = '1' " +
                    "AND t.biosum_cond_id = t1.biosum_cond_id " +
                    "AND t.rxpackage = t1.rxpackage) " +
                    "GROUP BY t.biosum_cond_id, t.rxpackage, t.rx";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlQueryReader(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");

                if (m_oDataMgr.m_DataReader.HasRows)
                {
                    long lngCount = 0;
                    string strRxCycle = "1";
                    int intValuePlaceholder = 0;
                    //For each condition id/rxPackage combination returned by the query above
                    while (m_oDataMgr.m_DataReader.Read())
                    {
                        string cond_id = "";
                        string rxpackage = "";
                        string rx = "";
                        if (m_oDataMgr.m_DataReader["biosum_cond_id"] != System.DBNull.Value)
                            cond_id = m_oDataMgr.m_DataReader["biosum_cond_id"].ToString().Trim();
                        if (m_oDataMgr.m_DataReader["rxpackage"] != System.DBNull.Value)
                            rxpackage = m_oDataMgr.m_DataReader["rxpackage"].ToString().Trim();
                        if (m_oDataMgr.m_DataReader["rx"] != System.DBNull.Value)
                            rx = m_oDataMgr.m_DataReader["rx"].ToString().Trim();

                        //Insert a placeholder row with default values
                        m_oDataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " " +
                            "(biosum_cond_id, rxpackage, rx, rxcycle, " +
                            "complete_cpa, harvest_cpa," +
                            "DateTimeCreated, place_holder) " +
                            "VALUES ('" + cond_id + "', '" + rxpackage + "', '" + rx + "', '" + strRxCycle + "', " +
                            intValuePlaceholder + ", " + intValuePlaceholder +
                            ", '" + m_strDateTimeCreated + "', 'Y')";

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n INSERT RECORD: " + System.DateTime.Now.ToString() + "\r\n");
                        m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                        if (m_oDataMgr.m_intError != 0) break;
                        lngCount++;
                    }
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, " \r\n END INSERTED " + lngCount + " RECORDS: " + System.DateTime.Now.ToString() + "\r\n");
                }
            }
            m_intError = m_oDataMgr.m_intError;
            m_strError = m_oDataMgr.m_strError;
        }

        private string RunScenario_CopyOPCOSTTables(string p_strVariant, string p_strRxPackage, string p_strRx1, string p_strRx2,
            string p_strRx3, string p_strRx4, int p_intMinPercentOf2GB)
        {
            string strInputPath = frmMain.g_oFrmMain.getProjectDirectory() + "\\OPCOST\\Input";
            string strInputFile = System.IO.Path.GetFileNameWithoutExtension(frmMain.g_strOPCOSTDirectory) + "_Input_" +
                           p_strVariant + "_P" + p_strRxPackage + "_" + p_strRx1 + "_" + p_strRx2 + "_" + p_strRx3 + "_" + p_strRx4 + "_" + m_strDateTimeCreated + ".accdb";
            strInputFile = strInputFile.Replace(":", "_");
            strInputFile = strInputFile.Replace(" ", "_");
            System.IO.File.Copy(m_oQueries.m_strTempDbFile, strInputPath + "\\" + strInputFile, true);
            //I am setting this back to 5000 from 2500 to resolve an opcost file contention problem
            System.Threading.Thread.Sleep(5000);
            //delete the work tables and any links
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strInputPath + "\\" + strInputFile, "", ""), 5);
            //if (m_oAdo.m_intError == 0)
            //{
            string[] strTables = m_oAdo.getTableNames(m_oAdo.m_OleDbConnection);
            if (strTables != null)
            {
                for (int z = 0; z <= strTables.Length - 1; z++)
                {
                    if (strTables[z] != null)
                    {
                        switch (strTables[z].Trim().ToUpper())
                        {
                            case "OPCOST_ERRORS": break;
                            //case "OPCOST_IDEAL_ERRORS": break;
                            case "OPCOST_INPUT": break;
                            case "OPCOST_OUTPUT": break;
                            //case "OPCOST_IDEAL_OUTPUT": break;
                            default:
                                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + strTables[z].Trim());
                                break;

                        }
                    }
                }
            }

            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            System.Threading.Thread.Sleep(5000);
            if (uc_filesize_monitor1.CurrentPercent(strInputPath + "\\" + strInputFile, 2000000000) > p_intMinPercentOf2GB)
            {
                dao_data_access oDao = new dao_data_access();
                oDao.m_DaoDbEngine.Idle(1);
                oDao.m_DaoDbEngine.Idle(8);
                oDao.CompactMDB(strInputPath + "\\" + strInputFile);
                m_intError = oDao.m_intError;
                oDao.m_DaoWorkspace.Close();
                oDao.m_DaoDbEngine = null;
                oDao = null;
                System.Threading.Thread.Sleep(5000);
            }
            return strInputPath + "\\" + strInputFile;
        }

        private void RunScenario_StartNew()
        {
            ReferenceProcessorScenarioForm.tlbScenario.Enabled = false;
            ReferenceProcessorScenarioForm.EnableTabPage(ReferenceProcessorScenarioForm.tabControlScenario, "tbDataSources", false);
            ReferenceProcessorScenarioForm.EnableTabPage(ReferenceProcessorScenarioForm.tabControlScenario, "tbTreeGroupings", false);
            ReferenceProcessorScenarioForm.EnableTabPage(ReferenceProcessorScenarioForm.tabControlRules, "tbHarvestMethod,tbWoodValue,tbEscalators,tbAddHarvestCosts,tbFilterCond,tbMoveInCosts", false);
            string strPath = frmMain.g_oFrmMain.getProjectDirectory() + "\\OPCOST\\Input";
            if (!System.IO.Directory.Exists(strPath))
                System.IO.Directory.CreateDirectory(strPath);

            btnChkAll.Enabled = false;
            btnUncheckAll.Enabled = false;

            this.lblMsg.Text = "";
            this.lblMsg.Show();
            this.m_strDateTimeCreated = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt");

            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(this.RunScenario_MainNew));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.m_oThread.Start();
        }

        private void RunScenario_MainNew()
        {
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;

            int x, y, z;
            int intCount = 0;
            int intRowCount = 0;

            dao_data_access oDao = new dao_data_access();

            string strRx1, strRx2, strRx3, strRx4, strRxPackage, strVariant, strCutCount;

            frmMain.g_oDelegate.CurrentThreadProcessName = "main";
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;

            if (System.IO.File.Exists(m_strDebugFile))
                System.IO.File.Delete(m_strDebugFile);

            System.Threading.Thread.Sleep(2000);

            if (frmMain.g_bDebug)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "*****START*****" + System.DateTime.Now.ToString() + "\r\n");

            _lstErrorVariants = new System.Collections.Generic.List<string>();
            m_intLvCheckedCount = 0;
            m_intLvTotalCount = this.m_lvEx.Items.Count;
            for (x = 0; x <= this.m_lvEx.Items.Count - 1; x++)
            {
                ReferenceProgressBarEx = (ProgressBarEx.ProgressBarEx)this.m_lvEx.GetEmbeddedControl(COL_RUNSTATUS, x);
                ReferenceProgressBarEx.backgroundpainter.Color = ReferenceProgressBarEx.backgroundpainter.DefaultColor;
                frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", 0);
                if ((bool)frmMain.g_oDelegate.GetListViewExItemPropertyValue(this.m_lvEx, x, "Checked", false))
                {
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Text", "0%");
                    m_intLvCheckedCount++;

                }
                else
                {
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Text", "");

                }
                frmMain.g_oDelegate.ExecuteControlMethod(ReferenceProgressBarEx, "Refresh");

            }
            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Prepare for processing...Stand By");

            // Create temp SQLite database; The temp Access database is created in loadValues
            if (ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                m_strTempSqliteDbFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
                m_oDataMgr.CreateDbFile(m_strTempSqliteDbFile);
                //@ToDo: Close this connection
                m_oDataMgr.OpenConnection(m_oDataMgr.GetConnectionString(m_strTempSqliteDbFile));
                //attach to scenario_rule_definitions.db
                string strScenarioDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                m_oDataMgr.m_strSQL = "attach database '" + strScenarioDB + "' as definitions";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + m_oDataMgr.m_strSQL + "\r\n");
                m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                //attach to scenario_results.db
                m_oDataMgr.m_strSQL = "attach '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                      "\\processor\\" + ScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultSqliteResultsDbFile +
                      "' as 'results'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
            }

            string strOpcostInputPath = "";
            for (x = 0; x <= this.m_lvEx.Items.Count - 1; x++)
            {
                if ((bool)frmMain.g_oDelegate.GetListViewExItemPropertyValue(this.m_lvEx, x, "Checked", false))
                {

                    if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.UserControl)uc_filesize_monitor1, "Visible", false) == false)
                    {
                        uc_filesize_monitor1.BeginMonitoringFile(
                            m_oQueries.m_strTempDbFile, 2000000000, "2GB");
                        uc_filesize_monitor1.Information = "Work table containing table links";
                        uc_filesize_monitor2.BeginMonitoringFile(
                             frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                            "\\processor\\" + ScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultSqliteResultsDbFile
                            , 2000000000, "2GB");
                        uc_filesize_monitor2.Information = "Scenario results DB file containing Harvest Costs and Tree Volume and Value tables";
                    }
                    frmMain.g_oDelegate.EnsureListViewExItemVisible(this.m_lvEx, x);

                    frmMain.g_oDelegate.SetListViewExItemPropertyValue(this.m_lvEx, x, "Selected", true);
                    frmMain.g_oDelegate.SetListViewExItemPropertyValue(this.m_lvEx, x, "Focused", true);

                    this.m_intError = 0;
                    this.m_strError = "";
                    ReferenceProgressBarEx = (ProgressBarEx.ProgressBarEx)this.m_lvEx.GetEmbeddedControl(COL_RUNSTATUS, x);

                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Maximum", 100);

                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Minimum", 0);
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", 0);
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Text", "0%");

                    frmMain.g_oDelegate.SetStatusBarPanelTextValue(
                               frmMain.g_sbpInfo.Parent,
                               1,
                               "Processing " + Convert.ToString(intCount + 1) + " Of " + Convert.ToString(frmMain.g_oDelegate.GetListViewExCheckedItemsCount(m_lvEx, false)) + "...Stand By");

                    strVariant = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_VARIANT, "Text", false);
                    strVariant = strVariant.Trim();

                    //get the package and treatments
                    strRxPackage = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_PACKAGE, "Text", false);
                    strRxPackage = strRxPackage.Trim();

                    strRx1 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_RXCYCLE1, "Text", false);
                    strRx1 = strRx1.Trim();

                    strRx2 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_RXCYCLE2, "Text", false);
                    strRx2 = strRx2.Trim();

                    strRx3 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_RXCYCLE3, "Text", false);
                    strRx3 = strRx3.Trim();

                    strRx4 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_RXCYCLE4, "Text", false);
                    strRx4 = strRx4.Trim();

                    strCutCount = (string) frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_CUTCOUNT, "Text", false);
                    bool _bInactiveVarRxPackage = false;
                    if (strCutCount.Trim().Equals("0"))
                        _bInactiveVarRxPackage = true;

                    m_strOPCOSTBatchFile = frmMain.g_oEnv.strTempDir + "\\" +
                        "OPCOST_Input_P" + strRxPackage + "_" + strRx1 + "_" + strRx2 + "_" + strRx3 + "_" + strRx4 + ".BAT";

                    if (System.IO.File.Exists(m_strOPCOSTBatchFile))
                        System.IO.File.Delete(m_strOPCOSTBatchFile);

                    //find the package item in the package collection
                    for (y = 0; y <= this.m_oRxPackageItem_Collection.Count - 1; y++)
                    {
                        if (this.m_oRxPackageItem_Collection.Item(y).SimulationYear1Rx.Trim() == strRx1.Trim() &&
                            this.m_oRxPackageItem_Collection.Item(y).SimulationYear2Rx.Trim() == strRx2.Trim() &&
                            this.m_oRxPackageItem_Collection.Item(y).SimulationYear3Rx.Trim() == strRx3.Trim() &&
                            this.m_oRxPackageItem_Collection.Item(y).SimulationYear4Rx.Trim() == strRx4.Trim() &&
                            this.m_oRxPackageItem_Collection.Item(y).RxPackageId.Trim() == strRxPackage.Trim())
                            break;


                    }
                    if (y <= m_oRxPackageItem_Collection.Count - 1)
                    {
                        this.m_oRxPackageItem = new RxPackageItem();
                        m_oRxPackageItem.CopyProperties(m_oRxPackageItem_Collection.Item(y), m_oRxPackageItem);
                    }
                    else
                    {
                        this.m_oRxPackageItem = null;
                    }

                    //get the list of treatment cycle year fields to reference for this package
                    this.m_strRxCycleList = "";
                    if (strRx1.Trim().Length > 0 && strRx1.Trim() != "000") this.m_strRxCycleList = "1,";
                    if (strRx2.Trim().Length > 0 && strRx2.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "2,";
                    if (strRx3.Trim().Length > 0 && strRx3.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "3,";
                    if (strRx4.Trim().Length > 0 && strRx4.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "4,";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "// Dropping OpCost tables " + strVariant + strRxPackage + "\r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                    }

                    // Connection for m_oAdo is opened during loadvalues()
                    if (m_oAdo.m_OleDbConnection.State == ConnectionState.Closed)
                    {
                        m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""), 5);
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Reopening Access connection because it was closed" + "\r\n");
                        }
                    }
                    if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
                    {
                        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "HarvestCostsWorkTable") == true)
                            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE HarvestCostsWorkTable");

                        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "InactiveStandsWorkTable") == true)
                            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE InactiveStandsWorkTable");

                        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "TreeVolValLowSlope") == true)
                            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE TreeVolValLowSlope");
                    }
                    else
                    {
                        if (m_oDataMgr.TableExist(m_oDataMgr.m_Connection, "HarvestCostsWorkTable") == true)
                            m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, "DROP TABLE HarvestCostsWorkTable");

                        if (m_oDataMgr.TableExist(m_oDataMgr.m_Connection, "InactiveStandsWorkTable") == true)
                            m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, "DROP TABLE InactiveStandsWorkTable");

                        if (m_oDataMgr.TableExist(m_oDataMgr.m_Connection, "TreeVolValLowSlope") == true)
                            m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, "DROP TABLE TreeVolValLowSlope");
                    }

                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "opcost_input") == true)
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE opcost_input");

                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "opcost_output") == true)
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE opcost_output");

                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "opcost_errors") == true)
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE opcost_errors");

                    //Here we set the maximum number of ticks on the progress bar
                    //y cannot exceed this max number
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Maximum", 13);
                    
                    y = 0;

                    frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Load trees from cut list...Stand By");
                    y++;
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    string strSqliteConnection = "";
                    if (m_oDataMgr.m_Connection != null)
                    {
                        strSqliteConnection = m_oDataMgr.m_Connection.ConnectionString;
                    }
                    processor mainProcessor = new processor(m_strDebugFile, ScenarioId.Trim(), 
                        m_oAdo.m_OleDbConnection.ConnectionString, ReferenceProcessorScenarioForm.m_bUsingSqlite,
                        strSqliteConnection);
                    if (!_bInactiveVarRxPackage)
                    {
                        m_intError = mainProcessor.LoadTrees(strVariant, strRxPackage, m_oQueries.m_oFIAPlot.m_strCondTable,
                            m_oQueries.m_oFIAPlot.m_strPlotTable, m_oQueries.m_oReference.m_strRefHarvestMethodTable,
                            m_oQueries.m_oFvs.m_strRxTable);

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "//Processor.loadTrees return value: " + m_intError + "\r\n");
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                        }

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Update species codes and groups for trees...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            bool blnCreateReconcileTreesTable = false;
                            // print reconcile trees table if debug at highest level; This will be in temporary .accdb
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                blnCreateReconcileTreesTable = true;
                            m_intError = mainProcessor.UpdateTrees(strVariant, strRxPackage, m_oQueries.m_oFIAPlot.m_strTreeTable, 
                                m_oQueries.m_oTravelTime.m_strTravelTimeTable, this.m_oQueries.m_oFvs.m_strTreeSpcTable, blnCreateReconcileTreesTable);

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Processor.updateTrees return value: " + m_intError + "\r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                            }
                        }

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Creating OpCost Input...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            m_intError = mainProcessor.CreateOpcostInput(strVariant);

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Processor.createOpcostInput return value: " + m_intError + "\r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                            }
                        }

                        if (m_intError == 0)
                        {
                            m_oAdo.m_strSQL = "SELECT COUNT(*) AS reccount FROM opcost_input";

                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                        }

                        intCount++;

                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "OPCOST Processing Batch Input...Stand By");
                            RunScenario_ProcessOPCOST(strVariant, strRxPackage);
                        }
                        else
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "//OPCOST Processing Batch Input ERROR: " + m_strError + "\r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                            }
                        }
                        if (m_intError == 0)
                        {
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                        }
                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Update Tree Vol Val Table With Merch and Chip Market Values...Stand By");
                            y++;
                            frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                            // Note: This function disposes of local ado_data_access object
                            m_intError = mainProcessor.CreateTreeVolValWorkTable(m_strDateTimeCreated);

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Processor.createTreeVolValWorkTable return value: " + m_intError + "\r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                            }
                        }
                    }
                    else
                    {
                        y = y + 5;  // increment the progress indicator for inactive stands that skip steps
                    }

                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append OPCOST Data To Harvest Costs Work Table...Stand By");
                        RunScenario_AppendToHarvestCosts("HarvestCostsWorkTable", _bInactiveVarRxPackage);
                    }
                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }
                    string strInactiveStandsWorkTable = "InactiveStandsWorkTable";
                    if (m_intError == 0 && m_bIncludeInactiveStands == true)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Create Inactive Stands Work Table...Stand By");
                        RunScenario_CreateInactiveStandsHarvestCostsWorkTable("HarvestCostsWorkTable", strInactiveStandsWorkTable, strVariant, strRxPackage);
                    }
                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }

                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Update Harvest Costs Work Table With Additional Costs...Stand By");
                        RunScenario_UpdateHarvestCostsTableWithAdditionalCosts("HarvestCostsWorkTable", strInactiveStandsWorkTable);
                    }
                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }

                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Delete Old Variant=" + strVariant + " and RxPackage=" + strRxPackage + " Records From Harvest Costs And Tree Vol Val Table...Stand By");
                        RunScenario_DeleteFromTreeVolValAndHarvestCostsTable(strVariant, strRxPackage);
                    }
                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }

                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append New Variant=" + strVariant + " and RxPackage=" + strRxPackage + " Records To Harvest Costs And Tree Vol Val Table...Stand By");
                        RunScenario_AppendToTreeVolValAndHarvestCostsTable();
                    }

                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }


                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append Placeholder Records For Variant=" + strVariant + " and RxPackage=" + strRxPackage + " To Tree Vol Val And Harvest Cost Tables...Stand By");
                        RunScenario_AppendPlaceholdersToTreeVolValAndHarvestCostsTables();
                    }


                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Finalizing Processor Scenario Database Tables...Stand By");
                        //update counts
                        intRowCount = 0;
                        if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
                        {
                            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "TreeVolValLowSlope") && !_bInactiveVarRxPackage)
                                intRowCount = (int)m_oAdo.getRecordCount(m_oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM TreeVolValLowSlope", "temp");

                            frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_VOLVAL, intRowCount.ToString());

                            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "HarvestCostsWorkTable"))
                                intRowCount = (int)m_oAdo.getRecordCount(m_oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM HarvestCostsWorkTable", "temp");
                            frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_HVSTCOST, intRowCount.ToString());
                        }
                        else
                        {
                            if (m_oDataMgr.TableExist(m_oDataMgr.m_Connection, "TreeVolValLowSlope") && !_bInactiveVarRxPackage)
                                intRowCount = (int)m_oDataMgr.getRecordCount(m_oDataMgr.m_Connection, "SELECT COUNT(*) FROM TreeVolValLowSlope", "temp");

                            frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_VOLVAL, intRowCount.ToString());

                            if (m_oDataMgr.TableExist(m_oDataMgr.m_Connection, "HarvestCostsWorkTable"))
                                intRowCount = (int)m_oDataMgr.getRecordCount(m_oDataMgr.m_Connection, "SELECT COUNT(*) FROM HarvestCostsWorkTable", "temp");
                            frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_HVSTCOST, intRowCount.ToString());

                        }

                        // Checking to see if opcost_input has > rows than harvest costs; If so, opcost dropped some records
                        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "opcost_input"))
                        {
                            int intOpcostRowCount = (int)m_oAdo.getRecordCount(m_oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM opcost_input", "temp");
                            intRowCount = intOpcostRowCount - intRowCount;
                            frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_OPCOSTDROP, intRowCount.ToString());
                            // If OpCost dropped records, set text color to red
                            if (intRowCount > 0)
                            {
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(m_lvEx, x, COL_OPCOSTDROP, "ForeColor", System.Drawing.Color.Red);
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(m_lvEx, x, COL_OPCOSTDROP, "UseItemStyleForSubItems", "False");
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                {
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, "********* Updating listbox totals " +  System.DateTime.Now.ToString() + " ************" + "\r\n");
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, "*Warning: harvest_costs table has " + intRowCount + " less records that opcost_input.* "  + "\r\n");
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, "*The totals should be the same. Check opcost_error file!                             * " + "\r\n");
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, "************************************************************************************** " + "\r\n");
                                }

                            }
                        }
                        else
                        {
                            frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_OPCOSTDROP, "0");
                        }
                    }

                    if (System.IO.File.Exists(m_strOPCOSTBatchFile))
                        System.IO.File.Delete(m_strOPCOSTBatchFile);

                    //compact mdb
                    if (m_intError == 0)
                    {

                        string strConn = m_oAdo.m_OleDbConnection.ConnectionString;
                        string strDb = m_oQueries.m_strTempDbFile;
                        m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
                        System.Threading.Thread.Sleep(5000);

                        //check if file size greater than 70% of 2GB
                        if (uc_filesize_monitor1.CurrentPercent(m_oQueries.m_strTempDbFile, 2000000000) > 70)
                        {
                            oDao.m_DaoDbEngine.Idle(1);
                            oDao.m_DaoDbEngine.Idle(8);
                            oDao.DisplayErrors = false;
                            oDao.m_intErrorCount = 0;
                            oDao.m_intError = -1;
                            for (; ; )
                            {
                                oDao.CompactMDB(strDb);
                                if (oDao.m_intError == 0)
                                {
                                    break;
                                }
                                if (oDao.m_intErrorCount == 5) break;
                                if (oDao.m_intErrorCount == 4)
                                {
                                    int count = oDao.m_intErrorCount;
                                    oDao.m_DaoDbEngine.Idle(1);
                                    oDao.m_DaoDbEngine.Idle(8);
                                    oDao.m_DaoWorkspace.Close();
                                    oDao.m_DaoDbEngine = null;
                                    oDao = null;
                                    oDao = new dao_data_access();
                                    oDao.m_intErrorCount = count;

                                }

                                oDao.m_intErrorCount++;
                                System.Threading.Thread.Sleep(5000);


                            }
                            System.Threading.Thread.Sleep(5000);
                        }
                        oDao.DisplayErrors = true;
                        if (oDao.m_intError != 0)
                        {
                            MessageBox.Show("Failed to compact and repair file " + m_oQueries.m_strTempDbFile, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }

                        strOpcostInputPath = RunScenario_CopyOPCOSTTables(strVariant, strRxPackage, strRx1, strRx2, strRx3, strRx4, 0);

                        m_oAdo.OpenConnection(strConn, 5);
                    }
                    else
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Error encountered...Saving tables to OPCOST directory...Stand By");
                        strOpcostInputPath = RunScenario_CopyOPCOSTTables(strVariant, strRxPackage, strRx1, strRx2, strRx3, strRx4, 0);
                    }

                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                        frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_CHECKBOX, " ");
                        frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_PROCESSOR_PROCESSINGDATETIME, m_strDateTimeCreated);
                    }
                    else
                    {
                        ReferenceProgressBarEx.backgroundpainter.Color = Color.Red;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", 0);
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Text", "!!Error!!");
                    }

                    System.Threading.Thread.Sleep(2000);
                }
            }
            oDao.m_DaoDbEngine.Idle(1);
            oDao.m_DaoDbEngine.Idle(8);
            oDao.m_DaoWorkspace.Close();
            oDao.m_DaoDbEngine = null;
            oDao = null;

            if (_lstErrorVariants.Count == 0)
            {
                MessageBox.Show("Done", "FIA Biosum");
            }
            else
            {
                string strVariantInfo = "";
                foreach (string strNextVariant in _lstErrorVariants)
                {
                    string[] strPieces = strNextVariant.Split('|');
                    if (strPieces.Length == 2)
                        strVariantInfo = strVariantInfo + strPieces[0] + String.Format("{0,8}", strPieces[1]) + "\r\n";
                }
                string strMessage = "Done with warnings. Biosum could not estimate costs for some Variant/Sequence combinations. " +
                    "Please review the opcost_errors tables in the most recent OPCOST databases located in the " +
                    frmMain.g_oFrmMain.getProjectDirectory() + "\\OPCOST\\Input folder for these Variant/Sequence combinations: \r\n\r\n" +
                    strVariantInfo;
                MessageBox.Show(strMessage, "FIA Biosum");
            }

            if (frmMain.g_bDebug)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Write processor parameters to text file \r\n");
		    string strDateFormat = "yyyy-MM-dd_HH-mm";
			string strFileDate = System.DateTime.Now.ToString(strDateFormat);
            string strParameterFileName = "params_" + this._strScenarioId + "_" + strFileDate + ".txt";
            string strParameterFilePath = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + ScenarioId + "\\" + strParameterFileName;
            ProcessorScenarioTools oProcessorScenarioTools = new ProcessorScenarioTools();
            string strProperties = oProcessorScenarioTools.ScenarioProperties(_frmProcessorScenario.m_oProcessorScenarioItem);
            System.IO.File.WriteAllText(strParameterFilePath, strProperties);

            if (frmMain.g_bDebug)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "*****END*****" + System.DateTime.Now.ToString() + "\r\n");

            RunScenario_Finished();


            frmMain.g_oDelegate.CurrentThreadProcessDone = true;
            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            if (this.btnRun.Text.Trim().ToUpper() == "CANCEL")
            {
                bool bAbort = frmMain.g_oDelegate.AbortProcessing("FIA Biosum", "Cancel Running The Processor Scenario (Y/N)?");
                if (bAbort)
                {
                    if (frmMain.g_oDelegate.m_oThread.IsAlive)
                    {
                        frmMain.g_oDelegate.m_oThread.Join();
                    }
                    frmMain.g_oDelegate.StopThread();
                    RunScenario_Finished();

                    frmMain.g_oDelegate.m_oThread = null;
                    ReferenceProgressBarEx.backgroundpainter.Color = Color.Red;
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", 0);

                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Text", "Cancelled");
                    frmMain.g_oDelegate.ExecuteControlMethod(ReferenceProgressBarEx, "Refresh");
                }
            }
            else
            {
                this.m_intError = 0;
                this.m_strError = "";
                if (this.m_lvEx.CheckedItems.Count == 0)
                {
                    MessageBox.Show("No Boxes Are Checked", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                //@ToDo: May want to take some of the conditions out when everyone is on OpCost 10 for performance
                string strOPCOSTVersion = "";
                bool bUsesOPCOSTRefAccdb = false;
                if (!String.IsNullOrEmpty(frmMain.g_strOPCOSTDirectory))
                {
                    strOPCOSTVersion = System.IO.Path.GetFileName(frmMain.g_strOPCOSTDirectory);
                    string[] strPieces = strOPCOSTVersion.Split('_');
                    if (strPieces.Length > 0)
                    {
                        int intVersion = -1;
                        bool isInteger = int.TryParse(strPieces[0], out intVersion);
                        if (intVersion > 9)
                            bUsesOPCOSTRefAccdb = true;
                    }
                }

                if (bUsesOPCOSTRefAccdb == true && ! System.IO.File.Exists(m_strOPCOSTRefPath))
                {
                    string strMessage = "BioSum could not find the OPCOST reference database! ";
                    strMessage += "It should be located at " + m_strOPCOSTRefPath + ". Processing halted!";
                    MessageBox.Show(strMessage, "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                //Check to make sure OpCost path is set; If not, try to use the default
                bool bOpcostFileExists = false;
                if (frmMain.g_strOPCOSTDirectory.Trim().Length > 0)
                {
                    if (System.IO.File.Exists(frmMain.g_strOPCOSTDirectory))
                    {
                        bOpcostFileExists = true;
                    }
                }
                else
                {
                    frmMain.g_strOPCOSTDirectory = frmSettings.GetDefaultOpcostPath();
                    if (!String.IsNullOrEmpty(frmMain.g_strOPCOSTDirectory))
                    {
                        bOpcostFileExists = true;
                    }
                }
                if (bOpcostFileExists == false)
                {
                    MessageBox.Show("!!A valid OpCost .R file has not been configured!!", "FIA Biosum",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                System.Collections.Generic.IList<string> lstRx = new System.Collections.Generic.List<string>();
                for (int x = 0; x <= this.m_lvEx.Items.Count - 1; x++)
                {
                    if ((bool)frmMain.g_oDelegate.GetListViewExItemPropertyValue(this.m_lvEx, x, "Checked", false))
                    {
                        string strRx1 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_RXCYCLE1, "Text", false);
                        strRx1 = strRx1.Trim();
                        lstRx.Add(strRx1);

                        string strRx2 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_RXCYCLE2, "Text", false);
                        strRx2 = strRx2.Trim();
                        if (! lstRx.Contains(strRx2))
                            lstRx.Add(strRx2);

                        string strRx3 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_RXCYCLE3, "Text", false);
                        strRx3 = strRx3.Trim();
                        if (!lstRx.Contains(strRx3))
                            lstRx.Add(strRx3);

                        string strRx4 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_RXCYCLE4, "Text", false);
                        strRx4 = strRx4.Trim();
                        if (!lstRx.Contains(strRx4))
                            lstRx.Add(strRx4);
                    }
                }

                ReferenceProcessorScenarioForm.ValidateRuleDefinitions(lstRx);
                m_intError = ReferenceProcessorScenarioForm.m_intError;

                if (this.m_intError == 0)
                {
                    ReferenceProcessorScenarioForm.SaveRuleDefinitions();
                    m_intError = ReferenceProcessorScenarioForm.m_intError;
                }

                if (this.m_intError == 0 && (frmMain.g_oDelegate.m_oThread == null ||
                                             frmMain.g_oDelegate.m_oThread.IsAlive == false))
                {

                    // Here is where we branch off to old processing, if desired 
                    Button btnCaller = (Button)sender;
                    if (btnCaller.Text.Equals("Run"))
                    {
                        btnCaller.Text = "Cancel";
                        RunScenario_StartNew();
                    }
                    else
                    {
                        btnRunOC7.Text = "Cancel";
                        RunScenario_Start();
                    }
                }

            }
        }
        
    }
}

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
        string m_strOPCOSTTimeStamp = "";
        string m_strOPCOSTBatchFile="";
        private string m_strDebugFile = frmMain.g_oEnv.strTempDir + "\\biosum_processor_debug.txt";
        private string m_strOPCOSTRefPath;
        //@ToDo: Remove these when removing old Processor code; These variable stand-in for checkboxes that are removed
        bool m_blnLowSlope = true;
        bool m_blnSteepSlope = true;
        
        Queries m_oQueries = new Queries();
        RxTools m_oRxTools = new RxTools();
        excel_latebinding.excel_latebinding m_oExcel=null;
        FIA_Biosum_Manager.RxPackageItem_Collection m_oRxPackageItem_Collection = null;
        FIA_Biosum_Manager.RxItem_Collection m_oRxItem_Collection = null;
        FIA_Biosum_Manager.RxPackageItem m_oRxPackageItem = null;
        string m_strRxCycleList = "";
        private ListViewColumnSorter lvwColumnSorter;
        private bool m_bLinkFvsComputeTables = false;
        private IList<string> m_lstAdditionalCpaColumns = null;
        private string m_strTempSqliteDbFile = null;
        string m_strScenarioDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
            "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;

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
        private const int COL_FVSTREE_PROCESSINGDATETIME = 12;
        private const int COL_PROCESSOR_PROCESSINGDATETIME = 13;
       

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
            string strErrMsg="";
            bool bUpdate;
            string strDebugFile = frmMain.g_oEnv.strTempDir + "\\biosum_processor_scenario_run_loadvalues.txt";
            if (System.IO.File.Exists(strDebugFile))
                System.IO.File.Delete(strDebugFile);
            
            if (frmMain.g_bDebug)
                frmMain.g_oUtils.WriteText(strDebugFile, "*****START*****" + System.DateTime.Now.ToString() + "\r\n");

            // INITIALIZE OPCOST REF PATH WHEN WE LOAD THE SCENARIO 
            m_strOPCOSTRefPath = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + 
                "\\" + Tables.Reference.DefaultOpCostReferenceDbFile;
            
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
            this.m_lvEx.Columns.Add("FVS_CutTree_DateTimeCreated", 80, HorizontalAlignment.Left);
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
            //SCENARIO RESULTS DB
            //
            string strScenarioResultsDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + ScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultSqliteResultsDbFile;

            //
            //LOAD PROJECT DATATASOURCES INFO
            //
            m_oQueries.m_oFvs.LoadDatasource = true;
            m_oQueries.m_oReference.LoadDatasource = true;
            m_oQueries.m_oProcessor.LoadDatasource = true;
            m_oQueries.m_oTravelTime.LoadDatasource = true;
            // The following call creates the temp database that is used for running calculations
            m_oQueries.LoadDatasources(true, "processor", ScenarioId);
            //
            //LOAD RX PACKAGE INFO
            //
            //load rxpackage properties
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "START: LoadAllRxPackageItemsFromTableIntoRxPackageCollection - " + System.DateTime.Now.ToString() + "\r\n");
            m_oRxPackageItem_Collection = new RxPackageItem_Collection();
            m_oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(m_oQueries, this.m_oRxPackageItem_Collection);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "END: LoadAllRxPackageItemsFromTableIntoRxPackageCollection - " + System.DateTime.Now.ToString() + "\r\n");

            //
            //LOAD RX INFO
            //
            //load rx properties
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "START: LoadAllRxItemsFromTableIntoRxCollection - " + System.DateTime.Now.ToString() + "\r\n");
            m_oRxItem_Collection = new RxItem_Collection();
            m_oRxTools.LoadAllRxItemsFromTableIntoRxCollection(m_oQueries, m_oRxItem_Collection);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "END: LoadAllRxItemsFromTableIntoRxCollection - " + System.DateTime.Now.ToString() + "\r\n");

            // Check PRE_FVS_COMPUTE table
            string strFvsPrePostDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutPrePostDbFile;
            IList<string> lstComponents = new List<string>();
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(strFvsPrePostDb)))
            {
                conn.Open();
                string[] arrFvsCompute = { Tables.FVS.DefaultPreFVSComputeTableName, Tables.FVS.DefaultPostFVSComputeTableName };
                if (!m_oDataMgr.TableExist(conn, arrFvsCompute[0]))
                {
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(strDebugFile, "loadvalues(): PRE_FVS_COMPUTE table was not found. KCP additional " +
                            "CPA functionality will not be used ! \r\n");
                }
                else
                {
                    //Check for KCP additional CPA columns so we know if we should include inRunScenario_UpdateHarvestCostsTableWithAdditionalCost active stands
                    for (int i = 0; i < m_oRxItem_Collection.Count; i++)
                    {
                        var rxItem = m_oRxItem_Collection.Item(i);
                        var componentCollection = rxItem.ReferenceHarvestCostColumnCollection;
                        for (int j = 0; j < componentCollection.Count; j++)
                        {
                            var component = componentCollection.Item(j);
                            if (!lstComponents.Contains(component.HarvestCostColumn))
                            {
                                lstComponents.Add(component.HarvestCostColumn);
                            }
                        }
                    }
                    m_lstAdditionalCpaColumns = new List<string>();
                    foreach (var outTable in arrFvsCompute)
                    {
                        for (int i = 0; i < lstComponents.Count; i++)
                        {
                            if (m_oDataMgr.ColumnExist(conn, outTable, lstComponents[i]))
                            {
                                string strSql = $@"SELECT YEAR FROM {outTable} WHERE {lstComponents[i]} = 1 LIMIT 1";
                                long lngCount = m_oDataMgr.getRecordCount(conn, strSql, outTable);
                                if (lngCount > 0)
                                {
                                    m_bLinkFvsComputeTables = true;
                                    if (!m_lstAdditionalCpaColumns.Contains(lstComponents[i]))
                                    {
                                        m_lstAdditionalCpaColumns.Add(lstComponents[i]);
                                    }
                                }
                            }
                        }
                    }
                    if (m_bLinkFvsComputeTables == false)
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(strDebugFile, "loadvalues(): Additional CPA fields set to 1 not found in FVS_COMPUTE table. Using legacy additional CPA calculations!\r\n");
                    }
                }
            }

            //
            //OPEN CONNECTION TO TEMP DB FILE
            //
            string strTempDbFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(strTempDbFile)))
            {
                conn.Open();
                if (m_oDataMgr.TableExist(conn, "ProcessorVariantPackageDateTimeCreated_work_table"))
                    m_oDataMgr.SqlNonQuery(conn, "DROP TABLE ProcessorVariantPackageDateTimeCreated_work_table");
                //
                //attach scenario results db
                //
                m_oDataMgr.m_strSQL = $@"ATTACH '{strScenarioResultsDb}' as RESULTS";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");
                //
                //attach master db
                //
                m_oDataMgr.m_strSQL = $@"ATTACH '{m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Plot)}' as MASTER";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");

                //
                //create temp work table with datetimecreated
                //
                m_oDataMgr.m_strSQL = "CREATE TABLE ProcessorVariantPackageDateTimeCreated_work_table " +
                    "AS SELECT DISTINCT biosum_cond_id,rxpackage,DateTimeCreated FROM " +
                    Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.AddColumn(conn,
                    "ProcessorVariantPackageDateTimeCreated_work_table",
                    "fvs_variant", "text", "2");
                m_oDataMgr.AddIndex(conn, "ProcessorVariantPackageDateTimeCreated_work_table", 
                    "ProcessorVariantPackageDateTimeCreated_work_table_idx1", "biosum_cond_id");

                m_oDataMgr.m_strSQL = $@"UPDATE ProcessorVariantPackageDateTimeCreated_work_table
                    SET fvs_variant = (SELECT p.fvs_variant FROM {m_oQueries.m_oFIAPlot.m_strPlotTable} p
                    INNER JOIN {m_oQueries.m_oFIAPlot.m_strCondTable} c ON p.biosum_plot_id = c.biosum_plot_id
                    WHERE ProcessorVariantPackageDateTimeCreated_work_table.biosum_cond_id = c.biosum_cond_id)
                    WHERE EXISTS (SELECT 1 FROM {m_oQueries.m_oFIAPlot.m_strPlotTable} p
                    INNER JOIN {m_oQueries.m_oFIAPlot.m_strCondTable} c ON p.biosum_plot_id = c.biosum_plot_id
                    WHERE ProcessorVariantPackageDateTimeCreated_work_table.biosum_cond_id = c.biosum_cond_id)";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlQueryReader(conn, m_oDataMgr.m_strSQL);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");

                m_oDataMgr.SqlNonQuery(conn, "DETACH database RESULTS");
            }
                
            //
            //GET LIST OF VARIANTS
            //
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "START: GetListOfFVSVariantsInPlotTable - " + System.DateTime.Now.ToString() + "\r\n");
            IDictionary<string, RxPackageItem_Collection> dictFvsVariantPackage = m_oRxTools.GetFvsVariantPackageDictionary(m_oQueries);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "END: GetListOfFVSVariantsInPlotTable - " + System.DateTime.Now.ToString() + "\r\n");

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "START: Populate List " + System.DateTime.Now.ToString() + "\r\n");
            string strFvsTreeListDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSTreeListDbFile;
            string strFvsTreeListConn = m_oDataMgr.GetConnectionString(strFvsTreeListDb) + ";datetimeformat=CurrentCulture";
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strFvsTreeListConn))
            {
                conn.Open();
                //populate the listview object
                string strRxPackage = "";
                string strCount = "";
                string strVariant = "";

                // link to PRE_FVS_COMPUTE table
                string strComputeDbPath = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutPrePostDbFile;
                if (m_bLinkFvsComputeTables && System.IO.File.Exists(strComputeDbPath))
                {
                    m_oDataMgr.SqlNonQuery(conn, $@"ATTACH DATABASE '{strComputeDbPath}' AS FVSPREPOST");
                }
                m_oDataMgr.m_strSQL = $@"ATTACH '{strTempDbFile}' AS WORKTABLE";
                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                m_oDataMgr.m_strSQL = $@"ATTACH '{strScenarioResultsDb}' as RESULTS";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.m_strSQL = $@"ATTACH '{m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Plot)}' as MASTER";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");

                foreach (string key in dictFvsVariantPackage.Keys)
                {
                    strVariant = key;
                    var oRxPackageItem_Collection = dictFvsVariantPackage[key];
                    for (int j = 0; j < oRxPackageItem_Collection.Count; j++)
                    {
                        bool _bInactiveVarRxPackage = false;
                        RxPackageItem rxPackageItem = oRxPackageItem_Collection.Item(j);
                        strRxPackage = rxPackageItem.RxPackageId;
                        if (frmMain.g_bSuppressProcessorScenarioTableRowCount == false)
                        {
                            m_oDataMgr.m_strSQL = "SELECT rxpackage, fvs_variant,COUNT(*) AS rxpackage_variant_count " +
                                                  "FROM " + Tables.FVS.DefaultFVSCutTreeTableName +
                                                  " WHERE FVS_VARIANT = '" + strVariant + "' AND" +
                                                  " RXPACKAGE = '" + strRxPackage + "'" +
                                                  " GROUP BY rxpackage,fvs_variant";
                        }
                        else
                        {
                            m_oDataMgr.m_strSQL = "SELECT rxpackage,fvs_variant, 1 AS rxpackage_variant_count FROM " + Tables.FVS.DefaultFVSCutTreeTableName +
                                                  " WHERE FVS_VARIANT = '" + strVariant + "' AND" +
                                                  " RXPACKAGE = '" + strRxPackage + "' LIMIT 1";
                        }
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                        m_oDataMgr.SqlQueryReader(conn, m_oDataMgr.m_strSQL);
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");
                        if (m_oDataMgr.m_DataReader.HasRows)
                        {
                            while (m_oDataMgr.m_DataReader.Read())
                            {
                                if (m_oDataMgr.m_DataReader["rxpackage"] != System.DBNull.Value)
                                    strRxPackage = m_oDataMgr.m_DataReader["rxpackage"].ToString().Trim();
                                if (m_oDataMgr.m_DataReader["fvs_variant"] != System.DBNull.Value)
                                    strVariant = m_oDataMgr.m_DataReader["fvs_variant"].ToString().Trim();
                                if (m_oDataMgr.m_DataReader["rxpackage_variant_count"] != System.DBNull.Value)
                                    strCount = m_oDataMgr.m_DataReader["rxpackage_variant_count"].ToString().Trim();
                            }
                            m_oDataMgr.m_DataReader.Close();
                        }
                        else if (m_bLinkFvsComputeTables == true)
                        {
                            IList<RxItem> lstRxItem = this.LoadRxItemsForRxPackage(strRxPackage);
                            for (int l = 0; l < lstRxItem.Count; l++)
                            {
                                var rxItem = lstRxItem[l];
                                var componentCollection = rxItem.ReferenceHarvestCostColumnCollection;
                                for (int k = 0; k < componentCollection.Count; k++)
                                {
                                    var component = componentCollection.Item(k);
                                    if (m_lstAdditionalCpaColumns.Contains(component.HarvestCostColumn))
                                    {
                                        // Check PRE_FVS_COMPUTE for 1
                                        string strSQL = $@"SELECT COUNT(*) FROM {Tables.FVS.DefaultPreFVSComputeTableName} WHERE 
                                        FVS_VARIANT = '{strVariant}' AND RXPACKAGE = '{strRxPackage}' AND RX = '{rxItem.RxId}' AND {component.HarvestCostColumn}= 1";
                                        long count = m_oDataMgr.getRecordCount(conn, strSQL, Tables.FVS.DefaultPreFVSComputeTableName);
                                        if (count > 0)
                                        {
                                            _bInactiveVarRxPackage = true;
                                            break;
                                        }
                                        else
                                        {
                                            // Check POST_FVS_COMPUTE for 1
                                            strSQL = $@"SELECT COUNT(*) FROM {Tables.FVS.DefaultPostFVSComputeTableName} WHERE 
                                                FVS_VARIANT = '{strVariant}' AND RXPACKAGE = '{strRxPackage}' AND RX = '{rxItem.RxId}' AND {component.HarvestCostColumn}= 1";
                                            count = m_oDataMgr.getRecordCount(conn, strSQL, Tables.FVS.DefaultPreFVSComputeTableName);
                                            if (count > 0)
                                            {
                                                _bInactiveVarRxPackage = true;
                                                break;
                                            }
                                        }
                                    }
                                }
                                if (_bInactiveVarRxPackage)
                                {
                                    break;
                                }
                            }
                            if (_bInactiveVarRxPackage == true)
                            {
                                strCount = "1";
                            }
                            else
                            {
                                // Reset strVariant so it fails the test to be added to the list
                                strVariant = "";
                            }
                        }
                        else
                        {
                            // Reset strVariant so it fails the test to be added to the list
                            // This variant package has no records in FVS_CutTree and isn't using FVS_Compute to indicate treatment activity
                            strVariant = "";
                        }
                        if (!String.IsNullOrEmpty(strVariant))
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(strDebugFile, "Add To List Variant:" + strVariant + " RxPackage:" + strRxPackage + " " + System.DateTime.Now.ToString() + "\r\n");

                            if (!cmbFilter.Items.Contains(strVariant))
                            {
                                cmbFilter.Items.Add(strVariant);
                            }
                            //find the package item
                            for (y = 0; y <= oRxPackageItem_Collection.Count - 1; y++)
                            {
                                if (oRxPackageItem_Collection.Item(y).RxPackageId.Trim() == strRxPackage.Trim())
                                {
                                    break;
                                }
                            }
                            if (y <= oRxPackageItem_Collection.Count - 1)
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

                                // tree vol val count
                                m_oDataMgr.m_strSQL = "SELECT COUNT(*) as rowcount " +
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
                                        frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                                    entryListItem.SubItems.Add(Convert.ToString(m_oDataMgr.getRecordCount(conn,
                                        m_oDataMgr.m_strSQL, "temp")));
                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                        frmMain.g_oUtils.WriteText(strDebugFile, "END SQL " + System.DateTime.Now.ToString() + "\r\n");
                                }
                                else
                                    entryListItem.SubItems.Add(" ");
                                //tree harvest cost count
                                m_oDataMgr.m_strSQL = "SELECT COUNT(*) as rowcount " +
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
                                        frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");

                                    entryListItem.SubItems.Add(Convert.ToString(m_oDataMgr.getRecordCount(conn,
                                        m_oDataMgr.m_strSQL, "temp")));

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

                                //fvstree processing date and time variant,rxpackage
                                string strFvsTreeDateCreated = " ";
                                if (_bInactiveVarRxPackage == false)
                                {
                                    m_oDataMgr.m_strSQL = "SELECT DISTINCT DateTimeCreated " +
                                                  "FROM " + Tables.FVS.DefaultFVSCutTreeTableName + " t " +
                                                  "WHERE t.fvs_variant ='" + strVariant + "' and t.rxpackage='" + strRxPackage + "'";
                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                        frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                                    strFvsTreeDateCreated = (string)m_oDataMgr.getSingleStringValueFromSQLQuery(conn, m_oDataMgr.m_strSQL, "temp");
                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                        frmMain.g_oUtils.WriteText(strDebugFile, "END SQL: " + System.DateTime.Now.ToString() + "\r\n");
                                }
                                entryListItem.SubItems.Add((strFvsTreeDateCreated.Trim())); //date and time created

                                m_oDataMgr.m_strSQL = "SELECT DateTimeCreated " +
                                                  "FROM  ProcessorVariantPackageDateTimeCreated_work_table " +
                                                  "WHERE rxpackage='" + strRxPackage + "' AND " +
                                                        "fvs_variant='" + strVariant + "'";
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                                string strCreated = (string)m_oDataMgr.getSingleStringValueFromSQLQuery(conn, m_oDataMgr.m_strSQL, "temp");
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(strDebugFile, "END SQL: " + System.DateTime.Now.ToString() + "\r\n");

                                //fvstree processing date and time variant,rxpackage 
                                if (strCreated.Length > 0)
                                {
                                    entryListItem.SubItems.Add(strCreated); //date and time created
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
                                strErrMsg = strErrMsg + "Table " + Tables.FVS.DefaultFVSCutTreeTableName + " contains RXPACKAGE: " + strRxPackage + " but " + strRxPackage + " is not a defined package. \r\n";
                            }
                        }


                    }
                }
            }
    
                if (strErrMsg.Trim().Length > 0)
                {
                    MessageBox.Show("Error loading Processor Scenario run list \r\n" + strErrMsg, "FIA Biosum",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(strDebugFile, "END: Populate List " + System.DateTime.Now.ToString() + "\r\n");
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
                if (this.m_lvEx != null)
                {
                    this.m_lvEx.Height = this.lblMsg.Top - this.m_lvEx.Top + 10;
                    this.m_lvEx.Width = this.panel1.Width - (m_lvEx.Left * 2);
                    this.btnRun.Left = this.m_lvEx.Width - (int)(m_lvEx.Width * .5) - (int)(btnRun.Width * .5);
                    this.btnRunOC7.Left = (int)this.btnRun.Left + 150;
                    this.lblMsg.Width = this.m_lvEx.Width;
                }

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
        private void RunScenario_ProcessOPCOST(string p_strVariant, string p_strRxPackage)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_ProcessOPCOST\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            m_oDataMgr.m_strSQL = "DROP TABLE temp_year";
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(m_strTempSqliteDbFile)))
            {
                conn.Open();
                if (m_oDataMgr.TableExist(conn, "temp_year"))
                    m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                string strFvsTreeListDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSTreeListDbFile;
                m_oDataMgr.m_strSQL = $@"ATTACH DATABASE '{strFvsTreeListDb}' AS FVSOUT";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                m_oDataMgr.m_strSQL = $@"CREATE TABLE temp_year AS SELECT DISTINCT biosum_cond_id||rxpackage||rx||rxcycle AS STAND,RXYEAR FROM {Tables.FVS.DefaultFVSCutTreeTableName}";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                m_oDataMgr.m_strSQL = "UPDATE opcost_input SET YearCostCalc = (SELECT cast(rxyear as integer) FROM temp_year b WHERE opcost_input.stand = b.stand)";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                m_oDataMgr.m_strSQL = "DROP TABLE temp_year";
                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
            }

            string strOPCOSTErrorFilePath = frmMain.g_oFrmMain.getProjectDirectory() + "\\OPCOST\\Input\\" +
                                             p_strVariant + "_" + p_strRxPackage + "_opcost_error_log.txt";
            bool bOPCOSTWindow = RunScenario_CreateOPCOSTBatchFile(strOPCOSTErrorFilePath);
            RunScenario_ExecuteOPCOST(bOPCOSTWindow, strOPCOSTErrorFilePath);
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(m_strTempSqliteDbFile)))
            {
                conn.Open();
                if (m_oDataMgr.TableExist(conn, Tables.ProcessorScenarioRun.DefaultOpcostErrorsTableName))
                {
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Querying " + Tables.ProcessorScenarioRun.DefaultOpcostErrorsTableName + "table \r\n");
                    int intCount = Convert.ToInt32(m_oDataMgr.getRecordCount(conn, "select count(*) from " + Tables.ProcessorScenarioRun.DefaultOpcostErrorsTableName,
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
            }
        }
        
        private void RunScenario_ExecuteOPCOST(bool bOPCOSTWindow, string strOPCOSTErrorFilePath)
        {

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

            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(m_strTempSqliteDbFile)))
            {
                conn.Open();
                if (!m_oDataMgr.TableExist(conn, "OPCOST_OUTPUT"))
                {
                    m_intError = -1;
                    m_strError = "!!OPCOST processing did not complete successfully. Check the error log at " +
                                strOPCOSTErrorFilePath + " for details!!";

                    MessageBox.Show(m_strError, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            //m_oAdo.OpenConnection(strConn);
            //if (!m_oAdo.TableExist(m_oAdo.m_OleDbConnection,"OPCOST_OUTPUT"))
            //{
            //    m_intError=-1;
            //    m_strError="!!OPCOST processing did not complete successfully. Check the error log at " +
            //                strOPCOSTErrorFilePath + " for details!!";

            //    MessageBox.Show(m_strError,"FIA Biosum",MessageBoxButtons.OK,MessageBoxIcon.Error);
            //}
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
            oTextStreamWriter.Write("SET INPUTFILE=" + m_strTempSqliteDbFile + "\r\n");
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
        private void RunScenario_AppendToHarvestCostsAccess(string p_strHarvestCostTableName, string strSqliteDb, bool bInactiveVarRxPackage)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_AppendToHarvestCostsAccess\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            // Create the work table
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(strSqliteDb)))
            {
                conn.Open();
                frmMain.g_oTables.m_oProcessor.CreateSqliteHarvestCostsTable(m_oDataMgr, conn, p_strHarvestCostTableName);
                dao_data_access oDao = new dao_data_access();
                if (!m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strHarvestCostTableName))
                {
                    oDao.CreateSQLiteTableLink(m_oAdo.m_OleDbConnection.DataSource, p_strHarvestCostTableName,
                         p_strHarvestCostTableName, ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName,
                         conn.FileName);
                    oDao.m_DaoWorkspace.Close();
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
                oDao = null;
            }


            if (! bInactiveVarRxPackage)
            {
                m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.AppendToOPCOSTHarvestCostsTable(
                    "OpCost_Output", "OpCost_Input", p_strHarvestCostTableName, m_strDateTimeCreated);
                if (m_oAdo.m_intError == 0 && frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                if (m_oAdo.m_intError == 0) m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
        }
        private void RunScenario_AppendToHarvestCosts(string p_strHarvestCostTableName, string strSqliteDb, bool bInactiveVarRxPackage)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_AppendToHarvestCostss\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            // Create the work table
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(strSqliteDb)))
            {
                conn.Open();
                frmMain.g_oTables.m_oProcessor.CreateSqliteHarvestCostsTable(m_oDataMgr, conn, p_strHarvestCostTableName);
                if (!bInactiveVarRxPackage)
                {
                    m_oDataMgr.m_strSQL = Queries.ProcessorScenarioRun.AppendToOPCOSTHarvestCostsTable(
                        "OpCost_Output", "OpCost_Input", p_strHarvestCostTableName, m_strDateTimeCreated);
                    if (m_oDataMgr.m_intError == 0 && frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    if (m_oDataMgr.m_intError == 0) m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                }
            }
            // Adding link to so p_strHarvestCostTableName that it can be used later in RunScenario_DeleteInaccessibleCondFromWorktable()
            // Needs to happen in Access because plot table is referenced later
            ODBCMgr odbcmgr = new ODBCMgr();
            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName);
            }
            odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName,
                strSqliteDb);
            dao_data_access oDao = new dao_data_access();
            oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, p_strHarvestCostTableName,
                p_strHarvestCostTableName, ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName, strSqliteDb);
            System.Threading.Thread.Sleep(5000);
            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
        }

        private void RunScenario_UpdateHarvestCostsTableWithAdditionalCostsAccess(string p_strHarvestCostsTableName, string p_strTempDb,
            processor.Escalators oEscalators, ProcessorScenarioItem.HarvestCostItem_Collection harvestCostItemCollection)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_UpdateHarvestCostsWithAdditionalCostsAccess\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            int x, y;
            string strSum = "";
            dao_data_access oDao = new dao_data_access();
            string[] strRXArray = null;
            //create work table to hold total additional costs
            //if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            //{
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(p_strTempDb)))
            {
                conn.Open();
                frmMain.g_oTables.m_oProcessorScenarioRun.CreateSqliteTotalAdditionalHarvestCostsTable(
                    m_oDataMgr, conn, "HarvestCostsTotalAdditionalWorkTable");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created HarvestCostsTotalAdditionalWorkTable \r\n");
            }

            if (!oDao.TableExists(m_oQueries.m_strTempDbFile, "HarvestCostsTotalAdditionalWorkTable"))
            {
                oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, "HarvestCostsTotalAdditionalWorkTable",
                    "HarvestCostsTotalAdditionalWorkTable", ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName, p_strTempDb);
                Thread.Sleep(8000);
            }

            //@ToDo: Need to update this for SQLite
            if (m_oDataMgr.m_intError == 0)
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
                    string[] strScenarioColumnNameArray = null;
                    string[] strScenarioRxArray = null;
                    if (harvestCostItemCollection.Count > 0)
                    {
                        /*****************************************************************
                         *the key is that the strScenarioColumnNameArray 
                         *and the strScenarioRxArray match in length
                         *and values as found in the scenario_harvest_cost_columns table
                         *****************************************************************/

                        strScenarioColumnNameArray = new string[harvestCostItemCollection.Count];
                        strScenarioRxArray = new string[harvestCostItemCollection.Count];
                        IList<string> lstUniqueRx = new List<string>();
                        for (int i = 0; i < harvestCostItemCollection.Count; i++)
                        {
                            ProcessorScenarioItem.HarvestCostItem oItem = harvestCostItemCollection.Item(i);
                            strScenarioColumnNameArray[i] = oItem.ColumnName;
                            strScenarioRxArray[i] = oItem.Rx;
                            if (!string.IsNullOrEmpty(oItem.Rx) && !lstUniqueRx.Contains(oItem.Rx))
                            {
                                lstUniqueRx.Add(oItem.Rx);
                            }
                        }

                        //update by rx that have both unique and global costs
                        strRXArray = lstUniqueRx.ToArray();
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
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(p_strTempDb)))
                    {
                        conn.Open();
                        if (m_oAdo.m_intError == 0 && m_oDataMgr.m_intError == 0)
                        {
                            m_oDataMgr.m_strSQL = "UPDATE HarvestCostsTotalAdditionalWorkTable " +
                                              "SET complete_additional_cpa = 0 " +
                                              "WHERE complete_additional_cpa IS NULL";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                        }

                        if (m_oDataMgr.m_intError == 0)
                        {
                            //update the harvest costs table complete costs per acre column
                            m_oAdo.m_strSQL = Queries.ProcessorScenarioRun.UpdateHarvestCostsTableWithCompleteCostsPerAcre(
                                                 "scenario_cost_revenue_escalators",
                                                 "HarvestCostsTotalAdditionalWorkTable",
                                                 p_strHarvestCostsTableName, ScenarioId, false);
                            m_oDataMgr.m_strSQL = Queries.ProcessorScenarioRun.UpdateSqliteHarvestCostsTableWithCompleteCostsPerAcre(
                                "HarvestCostsTotalAdditionalWorkTable", p_strHarvestCostsTableName, oEscalators, false);

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                        }
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
            //}
            //else
            //{
            //    RunScenario_UpdateHarvestCostsTableWithAdditionalCostsSqlite(p_strHarvestCostsTableName);
            //}
        }

        private void RunScenario_UpdateHarvestCostsTableWithAdditionalCosts(string p_strHarvestCostsTableName, string p_strTempDb,
            processor.Escalators oEscalators, ProcessorScenarioItem.HarvestCostItem_Collection harvestCostItemCollection)
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
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(p_strTempDb)))
            {
                conn.Open();
                frmMain.g_oTables.m_oProcessorScenarioRun.CreateSqliteTotalAdditionalHarvestCostsTable(
                    m_oDataMgr, conn, "HarvestCostsTotalAdditionalWorkTable");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created HarvestCostsTotalAdditionalWorkTable \r\n");

                string strScenarioParametersDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                m_oDataMgr.m_strSQL = "attach database '" + strScenarioParametersDb + "' as params";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);

                if (m_oDataMgr.m_intError == 0)
                {
                    //insert plot+rx records for the current scenario
                    m_oDataMgr.m_strSQL = "INSERT INTO HarvestCostsTotalAdditionalWorkTable " +
                                      "(biosum_cond_id,rx) SELECT biosum_cond_id,rx " +
                                                          "FROM scenario_additional_harvest_costs " +
                                                          "WHERE TRIM(UPPER(scenario_id)) = " +
                                                          "'" + ScenarioId.ToUpper().Trim() + "'";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                }

                if (m_oDataMgr.m_intError == 0)
                {
                    //
                    //GET THE ADDITIONAL HARVEST COST COLUMNS AND THEIR ASSOCIATED TREATMENT
                    //
                    m_oDataMgr.m_strSQL = "SELECT columnname,rx " +
                        "FROM scenario_harvest_cost_columns " +
                        "WHERE trim(UPPER(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    m_oDataMgr.SqlQueryReader(conn, m_oDataMgr.m_strSQL);

                }

                string[] strScenarioColumnNameArray = null;
                string[] strScenarioRxArray = null;
                if (m_oDataMgr.m_intError == 0)
                {
                    if (harvestCostItemCollection.Count > 0)
                    {
                        /*****************************************************************
                         *the key is that the strScenarioColumnNameArray 
                         *and the strScenarioRxArray match in length
                         *and values as found in the scenario_harvest_cost_columns table
                         *****************************************************************/

                        strScenarioColumnNameArray = new string[harvestCostItemCollection.Count];
                        strScenarioRxArray = new string[harvestCostItemCollection.Count];
                        IList<string> lstUniqueRx = new List<string>();
                        for (int i = 0; i < harvestCostItemCollection.Count; i++)
                        {
                            ProcessorScenarioItem.HarvestCostItem oItem = harvestCostItemCollection.Item(i);
                            strScenarioColumnNameArray[i] = oItem.ColumnName;
                            strScenarioRxArray[i] = oItem.Rx;
                            if (!string.IsNullOrEmpty(oItem.Rx) && !lstUniqueRx.Contains(oItem.Rx))
                            {
                                lstUniqueRx.Add(oItem.Rx);
                            }
                        }
                        //update by rx that have both unique and global costs
                        strRXArray = lstUniqueRx.ToArray();
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
                                              "FROM scenario_additional_harvest_costs as b " +
                                              "WHERE a.biosum_cond_id = b.biosum_cond_id " +
                                              "AND a.RX = b.RX " +
                                              "AND TRIM(b.scenario_id) = '" + ScenarioId.Trim() + "' " +
                                              "AND b.RX = '" + strRXArray[x] + "') " +
                                              "WHERE EXISTS(" +
                                              "SELECT * " +
                                              "FROM scenario_additional_harvest_costs AS b " +
                                              "WHERE a.biosum_cond_id = b.biosum_cond_id " +
                                              "AND a.RX = b.RX AND " +
                                              "TRIM(b.scenario_id)='" + ScenarioId.Trim() + "' AND " +
                                              "b.RX = '" + strRXArray[x] + "')";

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                        }
                    }
                }

                if (m_oDataMgr.m_intError == 0)
                {
                    IList<string> lstRx = new List<string>();
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

                    lstRx = m_oDataMgr.getStringList(conn, m_oDataMgr.m_strSQL);
                    strRXArray = lstRx.ToArray();
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
                                    strSum = strSum + "COALESCE(b." + strScenarioColumnNameArray[y].Trim() + ", 0) + ";
                                }
                            }

                            if (strSum.Trim().Length > 0)
                            {
                                strSum = strSum.Substring(0, strSum.Length - 2);
                                m_oDataMgr.m_strSQL = "UPDATE HarvestCostsTotalAdditionalWorkTable AS a " +
                                    "SET complete_additional_cpa = (SELECT " + strSum + " " +
                                    "FROM scenario_additional_harvest_costs as b " +
                                    "WHERE a.biosum_cond_id = b.biosum_cond_id " +
                                    "AND a.RX = b.RX " +
                                    "AND TRIM(b.scenario_id) = '" + ScenarioId.Trim() + "' " +
                                    "AND b.RX = '" + strRXArray[x] + "') " +
                                    "WHERE EXISTS( SELECT * FROM scenario_additional_harvest_costs AS b " +
                                    "WHERE a.biosum_cond_id = b.biosum_cond_id AND a.RX = b.RX AND " +
                                    "TRIM(b.scenario_id)='" + ScenarioId.Trim() + "' AND b.RX = '" + strRXArray[x] + "')";

                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
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
                    m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                }

                if (m_oDataMgr.m_intError == 0)
                {
                    //update the harvest costs table complete costs per acre column
                    //@ToDo: This query is bad
                    m_oDataMgr.m_strSQL = Queries.ProcessorScenarioRun.UpdateSqliteHarvestCostsTableWithCompleteCostsPerAcre(
                        "HarvestCostsTotalAdditionalWorkTable", p_strHarvestCostsTableName, oEscalators, false);

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                }

            }

            m_intError = m_oDataMgr.m_intError;
            m_strError = m_oDataMgr.m_strError;
        }

        private void RunScenario_UpdateHarvestCostsTableWithKcpAdditionalCosts(string p_strHarvestCostsTableName, string p_strAddCostsWorktable,
            string p_strTempDb, string p_strVariant, string p_strRxPackage, string p_strDateTimeCreated, processor.Escalators p_oEscalators,
            ProcessorScenarioItem.HarvestCostItem_Collection p_harvestCostItemCollection)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_UpdateHarvestCostsTableWithKcpAdditionalCosts\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(p_strTempDb)))
            {
                conn.Open();
                //Delete link to p_strAddCostsWorktable -- TEMPORARY until we can convert to SQLite
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strAddCostsWorktable))
                {
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, $@"DROP TABLE {p_strAddCostsWorktable}");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, $@"Dropped table link to {p_strAddCostsWorktable} {Environment.NewLine}");
                }
                if (m_oDataMgr.TableExist(conn, p_strAddCostsWorktable))
                {
                    m_oDataMgr.SqlNonQuery(conn, $@"DROP TABLE {p_strAddCostsWorktable}");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, $@"Dropped {p_strAddCostsWorktable} {Environment.NewLine}");
                }
                //create additional cpa work table to itemize additional costs
                frmMain.g_oTables.m_oProcessorScenarioRun.CreateSqliteAdditionalKcpCpaTable(
                    m_oDataMgr, conn, p_strAddCostsWorktable, true);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, $@"Created {p_strAddCostsWorktable} {Environment.NewLine}");
            }

            dao_data_access oDao = new dao_data_access();
            IList<RxItem> lstRxItem = this.LoadRxItemsForRxPackage(p_strRxPackage);
            StringBuilder sb = new StringBuilder();
            if (m_oAdo.m_intError == 0)
            {
                //
                //GET THE ADDITIONAL HARVEST COST COLUMNS AND THEIR ASSOCIATED TREATMENT
                //
                string strWhereRx = "";
                if (lstRxItem.Count > 0)
                {
                    sb.Append(" AND RX IN ( ");
                    foreach (var rxItem in lstRxItem)
                    {
                        sb.Append($@"'{rxItem.RxId}',");
                    }
                    strWhereRx = sb.ToString().TrimEnd(',') + " )";
                }

                IList<string> lstScenarioColumnNameList = new List<string>();
                bool bHasScenarioCosts = false;
                for (int i = 0; i < p_harvestCostItemCollection.Count; i++)
                {
                    var oItem = p_harvestCostItemCollection.Item(i);
                    if (m_lstAdditionalCpaColumns.Contains(oItem.ColumnName))
                    {
                        bool bFoundIt = false;
                        foreach (var rx in lstRxItem)
                        {
                            if (rx.RxId.Equals(oItem.Rx))
                            {
                                bFoundIt = true;
                                break;
                            }
                        }
                        if (bFoundIt && !lstScenarioColumnNameList.Contains(oItem.ColumnName))
                        {
                            lstScenarioColumnNameList.Add(oItem.ColumnName);
                        }
                    }
                    if (string.IsNullOrEmpty(oItem.Rx))
                    {
                        bHasScenarioCosts = true;
                    }
                }

                string strFlagSuffix = "_FLAG";
                if (m_oAdo.m_intError == 0)
                {
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(p_strTempDb)))
                    {
                        // Open a connection to SQLite so we can add the additional cpa columns
                        conn.Open();
                        foreach (var oItem in lstScenarioColumnNameList)
                        {
                            if (m_lstAdditionalCpaColumns.Contains(oItem))
                            {
                                //Need to add the column to the table
                                m_oDataMgr.AddColumn(conn, p_strAddCostsWorktable, oItem, "DOUBLE", "");
                                //Add the associated flag column to the table
                                m_oDataMgr.AddColumn(conn, p_strAddCostsWorktable, oItem + strFlagSuffix, "DOUBLE", "");
                            }
                        }
                    }

                    // ADD SCENARIO NAME COLUMN IF THERE ARE SCENARIO LEVEL ADDITIONAL_CPA
                    if (bHasScenarioCosts == true)
                    {
                        using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(p_strTempDb)))
                        {
                            conn.Open();
                            m_oDataMgr.AddColumn(conn, p_strAddCostsWorktable, ScenarioId.Trim(), "DOUBLE", "");
                        }
                    }

                    // LINK TO KcpCpaWorkTable IF THERE ARE NO ADDITIONAL_CPA COLUMNS
                    if (lstScenarioColumnNameList.Count == 0)
                    {
                        if (!oDao.TableExists(m_oQueries.m_strTempDbFile, p_strAddCostsWorktable))
                        {
                            // We need to link the SQLite worktable to the temp Access database because we populate it from Opcost_Input
                            oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, p_strAddCostsWorktable,
                                p_strAddCostsWorktable, ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName, p_strTempDb);
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, $@"Created link to {p_strAddCostsWorktable} {Environment.NewLine}");
                        }
                        Thread.Sleep(4000);
                    }

                    // UPDATE THE COLUMNS AND CREATE THE TABLE LINK TO THE additional_kcp_cpa TABLE AS SOON AS WE KNOW
                    // WHAT THE COLUMNS SHOULD BE SO THE DAO HAS TIME TO CREATE THE LINK BEFORE WE NEED THE TABLE
                    string strScenarioResultsDb =
                        frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                        "\\processor\\" + ScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultScenarioResultsTableDbFile;
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(strScenarioResultsDb)))
                    {
                        conn.Open();
                        foreach (var sName in lstScenarioColumnNameList)
                        {
                            if (!m_oDataMgr.ColumnExist(conn, Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName, sName))
                            {
                                m_oDataMgr.AddColumn(conn, Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName, sName, "DOUBLE", "");
                            }
                        }
                        if (bHasScenarioCosts && !m_oDataMgr.ColumnExist(conn, Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName, ScenarioId.Trim()))
                        {
                            m_oDataMgr.AddColumn(conn, Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName, ScenarioId.Trim(), "DOUBLE", "");
                        }
                    }

                    // Create link to additional_kcp_cpa after we have finalized the schema
                    ODBCMgr odbcmgr = new ODBCMgr();
                    if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorResultsDsnName))
                    {
                        odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorResultsDsnName);
                    }
                    odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.ProcessorResultsDsnName,
                        strScenarioResultsDb);
                    // Refresh the link to additional_kcp_cpa if it exists
                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName))
                    {
                        m_oAdo.m_strSQL = $@"DROP TABLE {Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName}";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    }
                    oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName,
                        Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName, ODBCMgr.DSN_KEYS.ProcessorResultsDsnName, strScenarioResultsDb);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, $@"Created link to {Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName} {Environment.NewLine}");
                    Thread.Sleep(6000);

                    // INSERT ROWS FROM PRE_FVS_COMPUTE TABLE; ASSUME ROW COUNT IS SAME FOR BOTH PRE AND POST TABLE
                    // INCLUDES INITIALLY SETTING THE FLAGS FROM THE PRE_FVS_COMPUTE TABLE
                    if (lstScenarioColumnNameList.Count > 0)
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, $@"RxPackage {p_strRxPackage} has FVS_COMPUTE KCP columns{Environment.NewLine}");
                        using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(p_strTempDb)))
                        {
                            conn.Open();
                            string strComputeDbPath = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutPrePostDbFile;
                            // Attach pre-post database
                            m_oDataMgr.m_strSQL = $@"ATTACH DATABASE '{strComputeDbPath}' AS FVS";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                            sb.Clear();
                            sb.Append($@"INSERT INTO {p_strAddCostsWorktable} (biosum_cond_id,rxpackage, rx, rxcycle");
                            foreach (var sName in lstScenarioColumnNameList)
                            {
                                sb.Append($@", {sName}{strFlagSuffix}");
                            }
                            sb.Append(")");
                            StringBuilder sbSelect = new StringBuilder();
                            sbSelect.Append("SELECT biosum_cond_id,rxpackage, rx, rxcycle");
                            foreach (var sName in lstScenarioColumnNameList)
                            {
                                sbSelect.Append($@", CASE WHEN {sName} IS NOT NULL THEN {sName} ELSE 0 END");
                                //sbSelect.Append($@", IIF({sName} IS NOT NULL,{sName},0)");
                            }
                            m_oDataMgr.m_strSQL = $@"{sb.ToString()}{sbSelect.ToString()}
                                         FROM {Tables.FVS.DefaultPreFVSComputeTableName} WHERE FVS_VARIANT = '{p_strVariant}'
                                         AND RXPACKAGE = '{p_strRxPackage}'";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);

                            // UPDATE THE FLAGS FROM THE POST_FVS_COMPUTE TABLE
                            sb.Clear();
                            sb.Append($@"UPDATE {p_strAddCostsWorktable} SET (");
                            string strUpdateColumns = "";
                            foreach (var sName in lstScenarioColumnNameList)
                            {
                                strUpdateColumns = strUpdateColumns + $@"{sName}{strFlagSuffix},";
                            }
                            sb.Append(strUpdateColumns.TrimEnd(','));
                            sb.Append(")=(SELECT ");
                            string strTemp = "";
                            foreach (var sName in lstScenarioColumnNameList)
                            {
                                strTemp = strTemp + $@"CASE WHEN {Tables.FVS.DefaultPostFVSComputeTableName}.{sName} IS NOT NULL THEN {sName}{strFlagSuffix} + {Tables.FVS.DefaultPostFVSComputeTableName}.{sName} ELSE {sName}{strFlagSuffix} END,";
                                //sb.Append($@"{sName}{strFlagSuffix} = IIF({Tables.FVS.DefaultPostFVSComputeTableName}.{sName} IS NOT NULL,{sName}{strFlagSuffix} + {Tables.FVS.DefaultPostFVSComputeTableName}.{sName},{sName}{strFlagSuffix}),");
                            }
                            sb.Append(strTemp.TrimEnd(','));    // trim last comma
                            sb.Append($@" FROM {Tables.FVS.DefaultPostFVSComputeTableName} WHERE {p_strAddCostsWorktable}.biosum_cond_id = {Tables.FVS.DefaultPostFVSComputeTableName}.biosum_cond_id and ");
                            sb.Append($@"{p_strAddCostsWorktable}.rxpackage = {Tables.FVS.DefaultPostFVSComputeTableName}.rxpackage and {p_strAddCostsWorktable}.rxcycle = {Tables.FVS.DefaultPostFVSComputeTableName}.rxcycle)");
                            m_oDataMgr.m_strSQL = sb.ToString();
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                        }
                    }
                    else
                    {
                        // BUILD THE TABLE FROM OPCOST_OUTPUT WHEN THERE ARE NO FVS_COMPUTE KCP COLUMNS
                        using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(p_strTempDb)))
                        {
                            conn.Open();
                            m_oDataMgr.m_strSQL = $@"INSERT INTO {p_strAddCostsWorktable} (biosum_cond_id, rxpackage, rx, rxcycle) 
                                             SELECT biosum_cond_id, rxpackage, rx, rxcycle FROM OpCost_Output";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                        }
                    }

                    // APPLY THE $/CPA FOR EACH RX
                    // Contains $/CPA

                    if (p_harvestCostItemCollection != null && p_oEscalators != null)
                    {
                        if (lstScenarioColumnNameList.Count > 0)
                        {
                            // UPDATE ONE RX AT A TIME; NOT ALL RX'S HAVE ADDITIONAL KCP HARVEST COSTS
                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(p_strTempDb)))
                            {
                                conn.Open();
                                foreach (var oRx in lstRxItem)
                                {
                                    var componentCollection = oRx.ReferenceHarvestCostColumnCollection;
                                    if (componentCollection.Count > 0)
                                    {
                                        string strSetValues = "";
                                        foreach (var sName in lstScenarioColumnNameList)
                                        {
                                            for (int i = 0; i < p_harvestCostItemCollection.Count; i++)
                                            {
                                                var hCostItem = p_harvestCostItemCollection.Item(i);
                                                if (hCostItem.Rx.Equals(oRx.RxId) && hCostItem.ColumnName.Equals(sName))
                                                {
                                                    strSetValues += $@"{sName} = {hCostItem.DefaultCostPerAcre}*{sName}{strFlagSuffix},";
                                                }
                                            }
                                        }
                                        strSetValues = strSetValues.TrimEnd(',');
                                        m_oDataMgr.m_strSQL = $@"UPDATE {p_strAddCostsWorktable} SET {strSetValues} WHERE RX = '{oRx.RxId}'";
                                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                        m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);

                                        //APPLY THE ESCALATORS TO ADDITIONAL CPA
                                        sb.Clear();
                                        sb.Append($@"UPDATE {p_strAddCostsWorktable} SET ");
                                        foreach (var sName in lstScenarioColumnNameList)
                                        {
                                            sb.Append($@"{sName} = CASE RXCycle WHEN '2' THEN {sName} * {p_oEscalators.OperatingCostsCycle2}
                                            WHEN '3' THEN {sName} * {p_oEscalators.OperatingCostsCycle3}
                                            WHEN '4' THEN {sName} * {p_oEscalators.OperatingCostsCycle4} ELSE {sName} END,");
                                        }
                                        strSetValues = sb.ToString().TrimEnd(',');
                                        m_oDataMgr.m_strSQL = $@"{strSetValues} WHERE RX = '{oRx.RxId}'";
                                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                        m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);

                                        //SUM THE KCP CPA COSTS
                                        if (lstScenarioColumnNameList.Count > 0)
                                        {
                                            sb.Clear();
                                            sb.Append($@"UPDATE {p_strAddCostsWorktable} SET ADDITIONAL_CPA = (");
                                            foreach (var sName in lstScenarioColumnNameList)
                                            {
                                                sb.Append($@" {sName} +");
                                            }
                                            strSetValues = sb.ToString().TrimEnd('+');
                                            strSetValues = strSetValues + ") WHERE RX = '" + oRx.RxId + "'";
                                            m_oDataMgr.m_strSQL = strSetValues;
                                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                                        }
                                    }
                                }
                            }  // End Using. Close SQLite connection


                            //UPDATE THE SCENARIO-LEVEL COSTS (IF APPLICABLE)
                            // THESE HAVE A DEPENDENCY ON MS ACCESS TABLES
                            if (bHasScenarioCosts)
                            {
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, " This scenario has scenario costs \r\n");

                                // Refresh the link to the worktable if it exists
                                if (!oDao.TableExists(m_oQueries.m_strTempDbFile, p_strAddCostsWorktable) && !ReferenceProcessorScenarioForm.m_bUsingSqlite)
                                {
                                    oDao.CreateSQLiteTableLink(m_oQueries.m_strTempDbFile, p_strAddCostsWorktable,
                                        p_strAddCostsWorktable, ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName, p_strTempDb);
                                    Thread.Sleep(4000);
                                }
                                foreach (var oRx in lstRxItem)
                                {
                                    string strSetValues = "";
                                    sb.Clear();
                                    foreach (ProcessorScenarioItem.HarvestCostItem item in p_harvestCostItemCollection)
                                    {
                                        if (string.IsNullOrEmpty(item.Rx) && item.DefaultCostPerAcre != null)
                                        {
                                            if (Convert.ToDouble(item.DefaultCostPerAcre) > 0)
                                            {
                                                sb.Append($@"{item.ColumnName}");
                                                sb.Append("+");
                                            }
                                        }
                                    }
                                        using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(p_strTempDb)))
                                        {
                                            conn.Open();
                                            string strScenarioParametersDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                                                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                                            m_oDataMgr.m_strSQL = "attach database '" + strScenarioParametersDb + "' as params";
                                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);

                                            if (sb.Length > 0)
                                            {
                                                strSetValues = $@" SET {ScenarioId.Trim()} = " + sb.ToString().TrimEnd('+');
                                                m_oDataMgr.m_strSQL = $@"UPDATE {p_strAddCostsWorktable} 
                                                    SET {ScenarioId.Trim()} = (SELECT {sb.ToString().TrimEnd('+')} from {Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName} 
                                                    WHERE biosum_cond_id = {p_strAddCostsWorktable}.biosum_cond_id AND rx={p_strAddCostsWorktable}.rx
                                                    and RX = '{oRx.RxId}' AND TRIM(UPPER(SCENARIO_ID)) = '{ScenarioId.Trim().ToUpper()}')";
                                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);

                                                //APPLY THE ESCALATORS AND SET ADDITIONAL_CPA = 0 IF NULL
                                                m_oDataMgr.m_strSQL = $@"UPDATE {p_strAddCostsWorktable} 
                                                SET {ScenarioId.Trim()} = CASE RXCycle WHEN '2' THEN {ScenarioId.Trim()} * {p_oEscalators.OperatingCostsCycle2} WHEN '3' THEN {ScenarioId.Trim()} * {p_oEscalators.OperatingCostsCycle3}
                                                WHEN '4' THEN {ScenarioId.Trim()} * {p_oEscalators.OperatingCostsCycle4} ELSE {ScenarioId.Trim()} END,
                                                ADDITIONAL_CPA = CASE ADDITIONAL_CPA WHEN NULL THEN 0 ELSE ADDITIONAL_CPA END WHERE RX = '{oRx.RxId}'";
                                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);

                                            }
                                            //ADD ON SCENARIO LEVEL COSTS IF KCP COSTS WERE INCURRED
                                            m_oDataMgr.m_strSQL = $@"UPDATE {p_strAddCostsWorktable} SET ADDITIONAL_CPA = 
                                                CASE WHEN ADDITIONAL_CPA > 0 THEN ADDITIONAL_CPA + {ScenarioId.Trim()} ELSE {ScenarioId.Trim()} END WHERE RX = '{oRx.RxId}'";
                                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);

                                        }  // End Using
                                }
                            }

                            if (oDao != null)
                            {
                                oDao.m_DaoWorkspace.Close();
                                oDao = null;
                            }

                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(p_strTempDb)))
                            {
                                conn.Open();
                                foreach (var oRx in lstRxItem)
                                {
                                    //DELETE WHERE ADDITIONAL_CPA = 0
                                    m_oDataMgr.m_strSQL = $@"DELETE FROM {p_strAddCostsWorktable} WHERE (ADDITIONAL_CPA = 0 OR ADDITIONAL_CPA IS NULL) AND RX = '{oRx.RxId}'";
                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                                    m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                                }
                            }
                        }
                    }
                    //if (m_oAdo.m_intError == 0)
                    //{
                    //    //
                    //    //WITHOUT TREATMENTS:
                    //    //PROCESS TREATMENTS THAT ARE IN THE SCENARIO_ADDITIONAL_HARVEST_COSTS table
                    //    //BUT ARE NOT IN THE SCENARIO_ADDITIONAL_HARVEST_COST_COLUMNS table
                    //    //
                    //    m_oAdo.m_strSQL = "SELECT DISTINCT  a.rx " +
                    //                      "FROM scenario_additional_harvest_costs a " +
                    //                      "WHERE NOT EXISTS (SELECT b.rx " +
                    //                                        "FROM scenario_harvest_cost_columns b " +
                    //                                        "WHERE b.rx=a.rx AND TRIM(UCASE(scenario_id))='" + ScenarioId.Trim().ToUpper() + "')";
                    //    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    //        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    //    strRXArray = frmMain.g_oUtils.ConvertListToArray(
                    //               (string)m_oAdo.CreateCommaDelimitedList(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, ""), ",");

                    //    if (strRXArray != null && strScenarioRxArray != null)
                    //    {

                    //        for (x = 0; x <= strRXArray.Length - 1; x++)
                    //        {
                    //            strSum = "";

                    //            for (y = 0; y <= strScenarioRxArray.Length - 1; y++)
                    //            {
                    //                //check 
                    //                if (strScenarioRxArray[y].Trim().Length > 0 && strScenarioRxArray[y] != "*")
                    //                {

                    //                }
                    //                else
                    //                {
                    //                    //scenario harvest cost
                    //                    strSum = strSum + "IIF(b." + strScenarioColumnNameArray[y].Trim() + " IS NOT NULL,b." + strScenarioColumnNameArray[y].Trim() + ",0) +";
                    //                }
                    //            }                    
                    //if (strSum.Trim().Length > 0)
                    //{
                    //    strSum = strSum.Substring(0, strSum.Length - 1);
                    //    m_oAdo.m_strSQL = "UPDATE " + strAddCpaWorkTable + " a " +
                    //                      "INNER JOIN scenario_additional_harvest_costs b " +
                    //                      "ON a.biosum_cond_id=b.biosum_cond_id AND " +
                    //                         "a.RX = b.RX " +
                    //                      "SET a.complete_additional_cpa = " + strSum + " " +
                    //                      "WHERE TRIM(UCASE(b.scenario_id))='" + ScenarioId.ToUpper().Trim() + "' AND " +
                    //                             "b.RX = '" + strRXArray[x] + "'";

                    //    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    //        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    //    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    //    if (m_oAdo.m_intError != 0) break;
                    //}

                    //        }
                    //    }
                    //}
                    //@ToDo: review this with new design
                    //if (m_oAdo.m_intError == 0)
                    //{
                    //    m_oAdo.m_strSQL = $@"UPDATE {strAddCpaWorkTable} SET complete_additional_cpa = 0
                    //                          WHERE complete_additional_cpa IS NULL";
                    //    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    //        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    //    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    //}
                    //INSERT INACTIVE STANDS WITH HARVEST COSTS INTO THE HARVEST COSTS WORK TABLE WHERE ADDITIONAL_CPA > 0
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(p_strTempDb)))
                    {
                        conn.Open();
                        if (m_oDataMgr.m_intError == 0)
                        {
                            m_oDataMgr.m_strSQL = $@"INSERT INTO {p_strHarvestCostsTableName} ( biosum_cond_id, rxpackage, rx, rxcycle, additional_cpa, harvest_cpa, complete_cpa, 
                                             chip_cpa, assumed_movein_cpa, place_holder, override_YN, DateTimeCreated )
                                             SELECT {p_strAddCostsWorktable}.biosum_cond_id, {p_strAddCostsWorktable}.rxpackage, {p_strAddCostsWorktable}.rx, {p_strAddCostsWorktable}.rxcycle, 
                                             0, 0, 0,0,0,'N', 'N', '{p_strDateTimeCreated}' FROM {p_strAddCostsWorktable}
                                             WHERE {p_strAddCostsWorktable}.additional_cpa > 0 AND (Exists (SELECT 1 
                                             FROM {p_strHarvestCostsTableName}
                                             WHERE {p_strHarvestCostsTableName}.biosum_cond_id = {p_strAddCostsWorktable}.biosum_cond_id 
                                             and {p_strHarvestCostsTableName}.rxpackage = {p_strAddCostsWorktable}.rxpackage
                                             and {p_strHarvestCostsTableName}.rx = {p_strAddCostsWorktable}.rx 
                                             and {p_strHarvestCostsTableName}.rxcycle = {p_strAddCostsWorktable}.rxcycle))=False";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                        }

                        if (m_oDataMgr.m_intError == 0)
                        {
                            //Update the harvest costs work table complete costs per acre column;
                            //Also applies the escalators to harvest_costs
                            m_oDataMgr.m_strSQL = Queries.ProcessorScenarioRun.UpdateSqliteHarvestCostsTableWithKcpCostsPerAcre(
                                         p_strAddCostsWorktable,
                                         p_strHarvestCostsTableName, p_oEscalators, true);
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                        }

                        if (m_oDataMgr.m_intError == 0)
                        {
                            //Set additional_cpa = 0 where additional_cpa = null so that complete_cpa can be updated
                            m_oDataMgr.m_strSQL = $@"UPDATE {p_strHarvestCostsTableName} SET ADDITIONAL_CPA = 0 WHERE ADDITIONAL_CPA IS NULL";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                        }

                        if (m_oDataMgr.m_intError == 0)
                        {
                            //Update the complete_cpa for stands where additional_cpa = 0; This should be the case where stands are in a project
                            //that uses kcp cpa, but kcp cpa is not defined for them
                            m_oDataMgr.m_strSQL = Queries.ProcessorScenarioRun.UpdateSqliteHarvestCostsTableWhenZeroKcpCosts(p_strHarvestCostsTableName,
                                p_oEscalators);
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                        }
                    }
                }
            }

            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;
        }

        private bool DoesVariantPackageUseKcpCpa(string p_strVariant, string p_strRxPackage)
        {
            bool bUsesKcpCpa = false;
            IList<string> lstAddCpaCols = new List<string>();
            // Are any harvest cost components configured for this rxPackage?
            IList<RxItem> lstRxItem = this.LoadRxItemsForRxPackage(p_strRxPackage);
            foreach (var rxItem in lstRxItem)
            {
                var componentCollection = rxItem.ReferenceHarvestCostColumnCollection;
                for (int i = 0; i < componentCollection.Count; i++)
                {
                    if (!lstAddCpaCols.Contains(componentCollection.Item(i).HarvestCostColumn))
                    {
                        lstAddCpaCols.Add(componentCollection.Item(i).HarvestCostColumn);
                    }
                }
            }

            if (m_lstAdditionalCpaColumns != null)
            {
                foreach (var aColumn in lstAddCpaCols)
                {
                    if (m_lstAdditionalCpaColumns.Contains(aColumn))
                    {
                        // Any rows in PRE_FVS_COMPUTE FOR THIS CPA COLUMN FOR THIS VARIANT/PACKAGE ?
                        string strComputeDbPath = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutPrePostDbFile;
                        using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(strComputeDbPath)))
                        {
                            conn.Open();
                            m_oDataMgr.m_strSQL = $@"SELECT {aColumn} FROM {Tables.FVS.DefaultPreFVSComputeTableName} WHERE FVS_VARIANT = '{p_strVariant}'
                                    AND RXPACKAGE = '{p_strRxPackage}' AND {aColumn} = 1 LIMIT 1";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            long lngCount = m_oDataMgr.getRecordCount(conn, m_oDataMgr.m_strSQL, Tables.FVS.DefaultPreFVSComputeTableName);
                            if (lngCount > 0)
                            {
                                bUsesKcpCpa = true;
                                break;
                            }

                            // Any rows in PRE_FVS_COMPUTE FOR THIS CPA COLUMN FOR THIS VARIANT/PACKAGE ?
                            m_oDataMgr.m_strSQL = $@"SELECT {aColumn} FROM {Tables.FVS.DefaultPostFVSComputeTableName} WHERE FVS_VARIANT = '{p_strVariant}'
                                    AND RXPACKAGE = '{p_strRxPackage}' AND {aColumn} = 1 LIMIT 1";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            lngCount = m_oDataMgr.getRecordCount(conn, m_oDataMgr.m_strSQL, Tables.FVS.DefaultPostFVSComputeTableName);
                            if (lngCount > 0)
                            {
                                bUsesKcpCpa = true;
                                break;
                            }
                        }
                    }
                }
            }

            return bUsesKcpCpa;
        }

        private void RunScenario_DeleteFromTreeVolValAndHarvestCostsTable(string p_strVariant, string p_strRxPackage,
            bool bRxPackageUseKcpAdditionalCpa)
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
            if (m_oAdo.m_intError == 0 && bRxPackageUseKcpAdditionalCpa == true)
            {
                //
                //DELETE THE VARIANT+RXPACKAGE FROM THE additional_kcp_cpa TABLE
                //
                m_oAdo.m_strSQL = $@"DELETE FROM {Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName} t 
                                  WHERE EXISTS (SELECT c.biosum_cond_id,p.fvs_variant FROM {this.m_oQueries.m_oFIAPlot.m_strCondTable} c,
                                  {this.m_oQueries.m_oFIAPlot.m_strPlotTable} p
                                  WHERE trim(t.rxpackage)='{p_strRxPackage}' AND (trim(t.biosum_cond_id)=c.biosum_cond_id AND 
                                  p.biosum_plot_id=c.biosum_plot_id AND p.fvs_variant='{p_strVariant}'))";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            m_intError = m_oAdo.m_intError;
            m_strError = m_oAdo.m_strError;

        }
        private void RunScenario_AppendToTreeVolValAndHarvestCostsTable(string p_strHarvestCostsTableName, string p_strAddCostsWorktable,
            string p_strTempDb)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunScenario_AppendDataToTreeVolVal\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            //if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            //{
            //    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "TreeVolValLowSlope"))
            //    {
            //        //
            //        //APPEND TO SCENARIO TREE VOL VAL TABLE
            //        //
            //        m_oAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " " +
            //                          "SELECT * FROM TreeVolValLowSlope";
            //        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            //            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            //        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            //    }
            //    if (m_oAdo.m_intError == 0 && m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strHarvestCostsTableName))
            //    {
            //        //
            //        //APPEND TO SCENARIO HARVEST COST TABLE
            //        //
            //        m_oAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " " +
            //                          "SELECT * FROM " + p_strHarvestCostsTableName;
            //        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            //            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            //        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            //    }
            //    if (!string.IsNullOrEmpty(p_strAddCostsWorktable))
            //    {
            //        if (m_oAdo.m_intError == 0 && m_oAdo.TableExist(m_oAdo.m_OleDbConnection, p_strAddCostsWorktable))
            //        {
            //            //
            //            //APPEND TO SCENARIO HARVEST COST TABLE
            //            //
            //            var dtSourceSchema = m_oAdo.getTableSchema(m_oAdo.m_OleDbConnection, "select * from " + p_strAddCostsWorktable);
            //            var dtDestSchema = m_oAdo.getTableSchema(m_oAdo.m_OleDbConnection, "select * from " + Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName);
            //            //build field list string to insert sql by matching FIADB and BioSum Tree columns
            //            string strCol = "";
            //            string strFields = "";
            //            StringBuilder sb = new StringBuilder();
            //            sb.Append("SET ");
            //            for (int x = 0; x <= dtDestSchema.Rows.Count - 1; x++)
            //            {
            //                strCol = dtDestSchema.Rows[x]["columnname"].ToString().Trim();
            //                //see if there is an equivalent source column
            //                for (int y = 0; y <= dtSourceSchema.Rows.Count - 1; y++)
            //                {
            //                    if (strCol.Trim().ToUpper() == dtSourceSchema.Rows[y]["columnname"].ToString().ToUpper())
            //                    {
            //                        strFields += strCol + ",";

            //                        break;
            //                    }
            //                }
            //                if (dtDestSchema.Rows[x]["datatype"].ToString().ToUpper().Equals("SYSTEM.DOUBLE"))
            //                {
            //                    sb.Append($@"{strCol}=IIF({strCol} IS NULL, 0, {strCol}),");
            //                }
            //            }
            //            strFields += $@"'{m_strDateTimeCreated}' AS DATETIMECREATED ";

            //            m_oAdo.m_strSQL = $@"INSERT INTO {Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName} 
            //                              SELECT {strFields} FROM {p_strAddCostsWorktable}";
            //            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            //                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            //            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

            //            //SET NULL VALUES TO 0
            //            if (sb.Length > 4)
            //            {
            //                string strSetValues = sb.ToString().TrimEnd(',');
            //                m_oAdo.m_strSQL = $@"UPDATE {Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName} {strSetValues}";
            //                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            //                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            //                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            //            }
            //        }
            //    }

            //    m_intError = m_oAdo.m_intError;
            //    m_strError = m_oAdo.m_strError;
            //}
            //else
            //{
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(p_strTempDb)))
            {
                conn.Open();
                string strScenarioResultsDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + ScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultSqliteResultsDbFile;
                m_oDataMgr.m_strSQL = "attach database '" + strScenarioResultsDB + "' as results";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + m_oDataMgr.m_strSQL + "\r\n");
                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);

                if (m_oDataMgr.TableExist(conn, "TreeVolValLowSlope"))
                {
                    //
                    //APPEND TO SCENARIO TREE VOL VAL TABLE
                    //
                    m_oDataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " (biosum_cond_id, rxpackage," +
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
                    m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                }
                if (m_oDataMgr.m_intError == 0 && m_oDataMgr.TableExist(conn, p_strHarvestCostsTableName))
                {
                    //
                    //APPEND TO SCENARIO HARVEST COST TABLE
                    //
                    m_oDataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " " +
                                      "SELECT * FROM " + p_strHarvestCostsTableName;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                    m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                }
                if (!string.IsNullOrEmpty(p_strAddCostsWorktable))
                {
                    if (m_oDataMgr.m_intError == 0 && m_oDataMgr.TableExist(conn, p_strAddCostsWorktable))
                    {
                        //
                        //APPEND TO SCENARIO HARVEST COST TABLE
                        //
                        var dtSourceSchema = m_oDataMgr.getTableSchema(conn, "select * from " + p_strAddCostsWorktable);
                        var dtDestSchema = m_oDataMgr.getTableSchema(conn, "select * from " + Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName);
                        //build field list string to insert sql by matching FIADB and BioSum Tree columns
                        string strCol = "";
                        string strFields = "";
                        StringBuilder sb = new StringBuilder();
                        sb.Append("SET ");
                        for (int x = 0; x <= dtDestSchema.Rows.Count - 1; x++)
                        {
                            strCol = dtDestSchema.Rows[x]["columnname"].ToString().Trim();
                            //see if there is an equivalent source column
                            for (int y = 0; y <= dtSourceSchema.Rows.Count - 1; y++)
                            {
                                if (strCol.Trim().ToUpper() == dtSourceSchema.Rows[y]["columnname"].ToString().ToUpper())
                                {
                                    strFields += strCol + ",";

                                    break;
                                }
                            }
                            if (dtDestSchema.Rows[x]["datatype"].ToString().ToUpper().Equals("SYSTEM.DOUBLE"))
                            {
                                //sb.Append($@"{strCol}=IIF({strCol} IS NULL, 0, {strCol}),");
                                sb.Append($@"{strCol}=CASE WHEN {strCol} IS NULL THEN 0 ELSE {strCol} END,");
                            }
                        }
                        string strTargetFields = strFields + "DATETIMECREATED ";
                        strFields += $@"'{m_strDateTimeCreated}' AS DATETIMECREATED ";

                        m_oDataMgr.m_strSQL = $@"INSERT INTO {Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName} 
                                              ({strTargetFields}) SELECT {strFields} FROM {p_strAddCostsWorktable}";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                        m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);

                        //SET NULL VALUES TO 0
                        if (sb.Length > 4)
                        {
                            string strSetValues = sb.ToString().TrimEnd(',');
                            m_oDataMgr.m_strSQL = $@"UPDATE {Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName} {strSetValues}";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                        }
                    }
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
                string strScenarioResultsDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + ScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultScenarioResultsTableDbFile;

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
                        using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(strScenarioResultsDb)))
                        {
                            conn.Open();
                            System.Data.SQLite.SQLiteTransaction transaction;
                            System.Data.SQLite.SQLiteCommand command = conn.CreateCommand();
                            // Start a local transaction
                            transaction = conn.BeginTransaction(IsolationLevel.ReadCommitted);
                            // Assign transaction object for a pending local transaction
                            command.Transaction = transaction;

                            try
                            {
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
                                    command.CommandText = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " " +
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
                                        frmMain.g_oUtils.WriteText(m_strDebugFile, command.CommandText + " \r\n INSERT RECORD: " + System.DateTime.Now.ToString() + "\r\n");
                                    command.ExecuteNonQuery();
                                    lngCount++;
                                }
                                transaction.Commit();
                            }
                            catch (Exception err)
                            {
                                m_intError = -1;
                                MessageBox.Show(err.Message);
                                transaction.Rollback();
                            }
                            transaction.Dispose();
                            transaction = null;
                        }
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, " \r\n END INSERTED " + lngCount + " RECORDS: " + System.DateTime.Now.ToString() + "\r\n");
                    }
                    m_oAdo.m_OleDbDataReader.Close();
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
                        using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(strScenarioResultsDb)))
                        {
                            conn.Open();
                            System.Data.SQLite.SQLiteTransaction transaction;
                            System.Data.SQLite.SQLiteCommand command = conn.CreateCommand();
                            // Start a local transaction
                            transaction = conn.BeginTransaction(IsolationLevel.ReadCommitted);
                            // Assign transaction object for a pending local transaction
                            command.Transaction = transaction;
                            string strRxCycle = "1";
                            int intValuePlaceholder = 0;
                            //For each condition id/rxPackage combination returned by the query above
                            try
                            {
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
                                    command.CommandText = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " " +
                                "(biosum_cond_id, rxpackage, rx, rxcycle, " +
                                "complete_cpa, harvest_cpa, chip_cpa, assumed_movein_cpa," +
                                "DateTimeCreated, place_holder) " +
                                "VALUES ('" + cond_id + "', '" + rxpackage + "', '" + rx + "', '" + strRxCycle + "', " +
                                intValuePlaceholder + ", " + intValuePlaceholder + ", " + intValuePlaceholder + ", " + intValuePlaceholder +
                                ", '" + m_strDateTimeCreated + "', 'Y')";

                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                        frmMain.g_oUtils.WriteText(m_strDebugFile, command.CommandText + " \r\n INSERT RECORD: " + System.DateTime.Now.ToString() + "\r\n");
                                    command.ExecuteNonQuery();
                                    lngCount++;
                                }
                                transaction.Commit();
                            }
                            catch (Exception err)
                            {
                                m_intError = -1;
                                MessageBox.Show(err.Message);
                                transaction.Rollback();
                            }
                            transaction.Dispose();
                            transaction = null;
                        }
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, " \r\n END INSERTED " + lngCount + " RECORDS: " + System.DateTime.Now.ToString() + "\r\n");
                    }
                    m_oAdo.m_OleDbDataReader.Close();
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

            string strScenarioResultsDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + ScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultScenarioResultsTableDbFile;

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
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(strScenarioResultsDb)))
                {
                    conn.Open();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + " " + System.DateTime.Now.ToString() + "\r\n");
                    m_oDataMgr.SqlQueryReader(conn, m_oDataMgr.m_strSQL);
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
                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                            if (m_oDataMgr.m_intError != 0) break;
                            lngCount++;
                        }
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, " \r\n END INSERTED " + lngCount + " RECORDS: " + System.DateTime.Now.ToString() + "\r\n");
                    }
                    m_oDataMgr.m_DataReader.Close();

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
                        m_oDataMgr.SqlQueryReader(conn, m_oDataMgr.m_strSQL);
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
                                    "complete_cpa, harvest_cpa, chip_cpa, assumed_movein_cpa," +
                                    "DateTimeCreated, place_holder) " +
                                    "VALUES ('" + cond_id + "', '" + rxpackage + "', '" + rx + "', '" + strRxCycle + "', " +
                                    intValuePlaceholder + ", " + intValuePlaceholder + ", " + intValuePlaceholder + ", " + intValuePlaceholder +
                                    ", '" + m_strDateTimeCreated + "', 'Y')";

                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + " \r\n INSERT RECORD: " + System.DateTime.Now.ToString() + "\r\n");
                                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                                if (m_oDataMgr.m_intError != 0) break;
                                lngCount++;
                            }
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, " \r\n END INSERTED " + lngCount + " RECORDS: " + System.DateTime.Now.ToString() + "\r\n");
                        }
                        m_oDataMgr.m_DataReader.Close();
                    }
                }   // Closing connection
            }
            m_intError = m_oDataMgr.m_intError;
            m_strError = m_oDataMgr.m_strError;
        }

        private string RunScenario_CopyOPCOSTTables(string p_strVariant, string p_strRxPackage, string p_strRx1, string p_strRx2,
            string p_strRx3, string p_strRx4, int p_intMinPercentOf2GB)
        {
            string strInputPath = frmMain.g_oFrmMain.getProjectDirectory() + "\\OPCOST\\Input";
            string strOpcostScript = System.IO.Path.GetFileNameWithoutExtension(frmMain.g_strOPCOSTDirectory).ToUpper();
            if (strOpcostScript.IndexOf("OPCOST") > -1)
            {
                strOpcostScript = strOpcostScript.Substring(strOpcostScript.IndexOf("OPCOST") + "OPCOST".Length);
            }
            string strInputFile = $@"{p_strVariant}_P{p_strRxPackage}_{m_strOPCOSTTimeStamp}{strOpcostScript}.db";
            System.IO.File.Copy(m_strTempSqliteDbFile, strInputPath + "\\" + strInputFile, true);
            //delete the work tables and any links
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(strInputPath + "\\" + strInputFile)))
            {
                conn.Open();
                string[] strTables = m_oDataMgr.getTableNames(conn);
                if (strTables != null)
                {
                    for (int z = 0; z <= strTables.Length - 1; z++)
                    {
                        if (strTables[z] != null)
                        {
                            switch (strTables[z].Trim().ToUpper())
                            {
                                case "OPCOST_ERRORS": break;
                                case "OPCOST_INPUT": break;
                                case "OPCOST_OUTPUT": break;
                                default:
                                    m_oDataMgr.SqlNonQuery(conn, "DROP TABLE " + strTables[z].Trim());
                                    break;
                            }
                        }
                    }
                }

            }
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

        private void RunScenario_Start()
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
            this.m_strOPCOSTTimeStamp = DateTime.Now.ToString("yyyyMMdd_HHmm");

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

            int x, y;
            int intCount = 0;
            int intRowCount = 0;

            dao_data_access oDao = new dao_data_access();
            ProcessorScenarioTools oProcessorScenarioTools = new ProcessorScenarioTools();

            string strRx1, strRx2, strRx3, strRx4, strRxPackage, strVariant, strCutCount;

            frmMain.g_oDelegate.CurrentThreadProcessName = "main";
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;

            if (System.IO.File.Exists(m_strDebugFile))
                System.IO.File.Delete(m_strDebugFile);

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
            m_strTempSqliteDbFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
            m_oDataMgr.CreateDbFile(m_strTempSqliteDbFile);

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
                    // Keep this flag to prevent Processor from trying to load trees
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

                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(m_strTempSqliteDbFile)))
                    {
                        conn.Open();
                        if (m_oDataMgr.TableExist(conn, "HarvestCostsWorkTable") == true)
                            m_oDataMgr.SqlNonQuery(conn, "DROP TABLE HarvestCostsWorkTable");

                        if (m_oDataMgr.TableExist(conn, "InactiveStandsWorkTable") == true)
                            m_oDataMgr.SqlNonQuery(conn, "DROP TABLE InactiveStandsWorkTable");

                        if (m_oDataMgr.TableExist(conn, "TreeVolValLowSlope") == true)
                            m_oDataMgr.SqlNonQuery(conn, "DROP TABLE TreeVolValLowSlope");

                        if (m_oDataMgr.TableExist(conn, "HarvestCostsTotalAdditionalWorkTable") == true)
                            m_oDataMgr.SqlNonQuery(conn, "DROP TABLE HarvestCostsTotalAdditionalWorkTable");

                        if (m_oDataMgr.TableExist(conn, "opcost_input") == true)
                            m_oDataMgr.SqlNonQuery(conn, "DROP TABLE opcost_input");

                        if (m_oDataMgr.TableExist(conn, "opcost_output") == true)
                            m_oDataMgr.SqlNonQuery(conn, "DROP TABLE opcost_output");

                        if (m_oDataMgr.TableExist(conn, "opcost_errors") == true)
                            m_oDataMgr.SqlNonQuery(conn, "DROP TABLE opcost_errors");
                    }

                    //Here we set the maximum number of ticks on the progress bar
                    //y cannot exceed this max number
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Maximum", 12);
                    
                    y = 0;

                    frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Load trees from cut list...Stand By");
                    y++;
                    frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    processor mainProcessor = new processor(m_strDebugFile, ScenarioId.Trim(), 
                        m_oAdo, ReferenceProcessorScenarioForm.m_bUsingSqlite,
                        m_strTempSqliteDbFile);
                    if (!_bInactiveVarRxPackage)
                    {
                        m_intError = mainProcessor.LoadTrees(strVariant, strRxPackage, m_oQueries.m_oFIAPlot.m_strCondTable,
                            m_oQueries.m_oFIAPlot.m_strPlotTable, m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.HarvestMethods),
                            m_oQueries.m_oReference.m_strRefHarvestMethodTable, m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Rx), 
                            m_oQueries.m_oFvs.m_strRxTable, m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Plot));
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
                                this.m_oQueries.m_oDataSource.getValidDataSourceTableName(Datasource.TableTypes.FiaTreeSpeciesReference), this.m_oQueries.m_oTravelTime.m_strDbFile,
                                this.m_oQueries.m_oTravelTime.m_strTravelTimeTable, blnCreateReconcileTreesTable);

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
                            //m_intError = mainProcessor.CreateOpcostInput(strVariant, strRxPackage);
                            m_intError = mainProcessor.CreateSqliteOpcostInput(strVariant, strRxPackage);

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Processor.createOpcostInput return value: " + m_intError + "\r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                            }
                        }

                        if (m_intError == 0)
                        {
                            //m_oAdo.m_strSQL = "SELECT COUNT(*) AS reccount FROM opcost_input";
                            //m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
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
                        RunScenario_AppendToHarvestCosts("HarvestCostsWorkTable", m_strTempSqliteDbFile, _bInactiveVarRxPackage);
                    }
                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }
                    string strKcpCpaWorkTable = "";
                    bool bRxPackageUsesKcpAdditionalCpa = DoesVariantPackageUseKcpCpa(strVariant, strRxPackage);

                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Update Harvest Costs Work Table With Additional Costs...Stand By");

                        // Had to get creative here because the ReferenceScenarioProcessorForm doesn't reflect configuration changes from the current session
                        processor.Escalators oEscalators = mainProcessor.LoadEscalators();
                        ProcessorScenarioItem oTempProcessorScenarioItem = new ProcessorScenarioItem();
                        oTempProcessorScenarioItem.ScenarioId = ScenarioId;
                        oProcessorScenarioTools.LoadHarvestCostComponents(m_strScenarioDb, oTempProcessorScenarioItem);
                        if (bRxPackageUsesKcpAdditionalCpa)
                        {
                            strKcpCpaWorkTable = "KcpCpaWorkTable";
                            RunScenario_UpdateHarvestCostsTableWithKcpAdditionalCosts("HarvestCostsWorkTable", strKcpCpaWorkTable,
                                m_strTempSqliteDbFile,strVariant, strRxPackage, m_strDateTimeCreated, oEscalators, oTempProcessorScenarioItem.m_oHarvestCostItem_Collection);
                        }
                        else
                        {
                            RunScenario_UpdateHarvestCostsTableWithAdditionalCosts("HarvestCostsWorkTable", m_strTempSqliteDbFile, oEscalators,
                                oTempProcessorScenarioItem.m_oHarvestCostItem_Collection);
                        }
                    }
                    if (m_intError == 0)
                    {
                        RunScenario_DeleteInaccessibleCondFromWorktable("HarvestCostsWorkTable");
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }
                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Delete Old Variant=" + strVariant + " and RxPackage=" + strRxPackage + " Records From Harvest Costs And Tree Vol Val Table...Stand By");
                        RunScenario_DeleteFromTreeVolValAndHarvestCostsTable(strVariant, strRxPackage, bRxPackageUsesKcpAdditionalCpa);
                    }
                    if (m_intError == 0)
                    {
                        y++;
                        frmMain.g_oDelegate.SetControlPropertyValue(ReferenceProgressBarEx, "Value", y);
                    }

                    if (m_intError == 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(lblMsg, "Text", "Append New Variant=" + strVariant + " and RxPackage=" + strRxPackage + " Records To Harvest Costs And Tree Vol Val Table...Stand By");
                        RunScenario_AppendToTreeVolValAndHarvestCostsTable("HarvestCostsWorkTable", strKcpCpaWorkTable, m_strTempSqliteDbFile);
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
                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_oDataMgr.GetConnectionString(m_strTempSqliteDbFile)))
                            {
                                conn.Open();
                                if (m_oDataMgr.TableExist(conn, "TreeVolValLowSlope") && !_bInactiveVarRxPackage)
                                    intRowCount = (int)m_oDataMgr.getRecordCount(conn, "SELECT COUNT(*) FROM TreeVolValLowSlope", "temp");

                                frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_VOLVAL, intRowCount.ToString());

                                if (m_oDataMgr.TableExist(conn, "HarvestCostsWorkTable"))
                                    intRowCount = (int)m_oDataMgr.getRecordCount(conn, "SELECT COUNT(*) FROM HarvestCostsWorkTable WHERE HARVEST_CPA <> 0", "temp");
                            }
                            frmMain.g_oDelegate.SetListViewTextValue(m_lvEx, x, COL_HVSTCOST, intRowCount.ToString());

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
                    string strOpcostInputPath = "";
                    if (m_intError == 0)
                    {
                        string strConn = m_oAdo.m_OleDbConnection.ConnectionString;
                        string strDb = m_oQueries.m_strTempDbFile;
                        m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
                        //System.Threading.Thread.Sleep(5000);

                        //check if file size greater than 70% of 2GB
                        //if (uc_filesize_monitor1.CurrentPercent(m_oQueries.m_strTempDbFile, 2000000000) > 70)
                        //{
                        //    oDao.m_DaoDbEngine.Idle(1);
                        //    oDao.m_DaoDbEngine.Idle(8);
                        //    oDao.DisplayErrors = false;
                        //    oDao.m_intErrorCount = 0;
                        //    oDao.m_intError = -1;
                        //    for (; ; )
                        //    {
                        //        oDao.CompactMDB(strDb);
                        //        if (oDao.m_intError == 0)
                        //        {
                        //            break;
                        //        }
                        //        if (oDao.m_intErrorCount == 5) break;
                        //        if (oDao.m_intErrorCount == 4)
                        //        {
                        //            int count = oDao.m_intErrorCount;
                        //            oDao.m_DaoDbEngine.Idle(1);
                        //            oDao.m_DaoDbEngine.Idle(8);
                        //            oDao.m_DaoWorkspace.Close();
                        //            oDao.m_DaoDbEngine = null;
                        //            oDao = null;
                        //            oDao = new dao_data_access();
                        //            oDao.m_intErrorCount = count;

                        //        }

                        //        oDao.m_intErrorCount++;
                        //        System.Threading.Thread.Sleep(5000);


                        //    }
                        //    System.Threading.Thread.Sleep(5000);
                        //}
                        //oDao.DisplayErrors = true;
                        //if (oDao.m_intError != 0)
                        //{
                        //    MessageBox.Show("Failed to compact and repair file " + m_oQueries.m_strTempDbFile, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        //}

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

                        string strVariant = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_VARIANT, "Text", false);
                        strVariant = strVariant.Trim();

                        string strRxPackage = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_PACKAGE, "Text", false);
                        strRxPackage = strRxPackage.Trim();

                        string strCutCount = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(m_lvEx, x, COL_CUTCOUNT, "Text", false);

                        // Validate that there are trees on the cut list if the stand isn't inactive
                        if (!strCutCount.Trim().Equals("0"))
                        {
                            ReferenceProcessorScenarioForm.ValidateCutList(strVariant, strRxPackage);
                            m_intError = ReferenceProcessorScenarioForm.m_intError;
                            if (this.m_intError != 0)
                            {
                                return;
                            }
                        }
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
                    Button btnCaller = (Button)sender;
                    if (btnCaller.Text.Equals("Run"))
                    {
                        btnCaller.Text = "Cancel";
                        RunScenario_Start();
                    }
                }

            }
        }

        private IList<RxItem> LoadRxItemsForRxPackage(string strRxPackage)
        {
            IList<RxItem> lstRxItem = new List<RxItem>();
            if (m_oRxPackageItem_Collection != null)
            {
                for (int i = 0; i < m_oRxPackageItem_Collection.Count; i++)
                {
                    var rxPackageItem = m_oRxPackageItem_Collection.Item(i);
                    if (rxPackageItem.RxPackageId == strRxPackage)
                    {
                        string[] arrRx = new string[] {rxPackageItem.SimulationYear1Rx, rxPackageItem.SimulationYear2Rx,
                                  rxPackageItem.SimulationYear3Rx, rxPackageItem.SimulationYear4Rx};
                        for (int j = 0; j < arrRx.Length; j++)
                        {
                            for (int k = 0; k < m_oRxItem_Collection.Count; k++)
                            {
                                var rxItem = m_oRxItem_Collection.Item(k);
                                if (rxItem.RxId == arrRx[j])
                                {
                                    if (!lstRxItem.Contains(rxItem))
                                    {
                                        lstRxItem.Add(rxItem);
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            return lstRxItem;
        }

        private void RunScenario_DeleteInaccessibleCondFromWorktable(string p_strHarvestCostsTableName)
        {
            // This relies on an MS Access table link to HarvestCostsWorkTable created in RunScenario_AppendToHarvestCosts
            m_oAdo.m_strSQL = $@"DELETE {p_strHarvestCostsTableName}.*
                FROM ({p_strHarvestCostsTableName} INNER JOIN cond ON {p_strHarvestCostsTableName}.biosum_cond_id = cond.biosum_cond_id) 
                INNER JOIN plot ON cond.biosum_plot_id = plot.biosum_plot_id
                WHERE plot.gis_yard_dist_ft < 0;";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
        }
    }
}

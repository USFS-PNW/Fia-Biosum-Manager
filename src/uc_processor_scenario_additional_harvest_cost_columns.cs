using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    public partial class uc_processor_scenario_additional_harvest_cost_columns : UserControl
    {
        public int m_intError = 0;
        public string m_strError = "";
        private uc_processor_scenario_additional_harvest_cost_column_collection uc_collection = new uc_processor_scenario_additional_harvest_cost_column_collection();
        ado_data_access m_oAdo=null;
        private string _strScenarioId = "";
        private frmProcessorScenario _frmProcessorScenario = null;
        System.Data.OleDb.OleDbConnection m_oConnAdditionalHarvestCosts = new System.Data.OleDb.OleDbConnection();
        RxTools m_oRxTools = new RxTools();
        private string _strTempDb = "";

        private FIA_Biosum_Manager.frmGridView m_frmHarvestCosts;
        public string[] m_strColumnsToEdit;
        public int m_intColumnsToEditCount = 0;
        public string m_strAddHarvCostsWorkTable = "additional_harvest_costs_work_table";

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

        public string TempDb
        {
            get { return _strTempDb; }
            set { _strTempDb = value; }
        }
        public uc_processor_scenario_additional_harvest_cost_column_collection ReferenceUserControlAdditionalHarvestCostColumnsCollection
        {
            get { return uc_collection; }
        }
        public uc_processor_scenario_additional_harvest_cost_columns()
        {
            InitializeComponent();
            this.uc_processor_scenario_additional_harvest_cost_column_item1.ReferenceAdditionalHarvestCostColumnsUserControl = this;
        }
        public void loadvalues()
        {
            int x;
            int y;
            int z;
            int zz;
            bool bFound = false;

            //
            //SCENARIO DB
            //
            string strScenarioDB =
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            //
            //SCENARIO ID
            //
            ScenarioId = this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();

            //
            //OPEN CONNECTION TO SCENARIO DEFINITIONS
            //
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            TempDb = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
            string strConn = dataMgr.GetConnectionString(TempDb);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();

                //attach processor scenario definitions
                dataMgr.m_strSQL = $@"ATTACH '{strScenarioDB}' AS DEFINITIONS";
                dataMgr.SqlNonQuery(oConn, dataMgr.m_strSQL);

                //delete work table if it exists
                if (dataMgr.TableExist(oConn, m_strAddHarvCostsWorkTable))
                {
                    dataMgr.m_strSQL = "DROP TABLE " + m_strAddHarvCostsWorkTable;
                    dataMgr.SqlNonQuery(oConn, dataMgr.m_strSQL);
                }
                //create a work table from the additional harvests costs table
                //
                dataMgr.m_strSQL = Tables.Processor.CreateAdditionalHarvestCostsTableSQL(m_strAddHarvCostsWorkTable);
                dataMgr.SqlNonQuery(oConn, dataMgr.m_strSQL);
                dataMgr.AddIndex(oConn, m_strAddHarvCostsWorkTable, $@"{m_strAddHarvCostsWorkTable}__plotrx", "biosum_cond_id,rx");

                //
                //add columns from the SOURCE scenario additional harvest costs table to the DESTINATION work table
                //
                System.Data.DataTable oSourceTableSchema = dataMgr.getTableSchema(oConn, "SELECT * FROM scenario_additional_harvest_costs");
                string strSourceColumnsList = dataMgr.getFieldNames(oConn, "SELECT * FROM scenario_additional_harvest_costs");
                string strDestColumnsList = dataMgr.getFieldNames(oConn, "SELECT * FROM " + m_strAddHarvCostsWorkTable);
                string[] strDestColumnsArray = frmMain.g_oUtils.ConvertListToArray(strDestColumnsList, ",");
                dataMgr.m_strSQL = "";
                for (z = 0; z <= oSourceTableSchema.Rows.Count - 1; z++)
                {

                    if (oSourceTableSchema.Rows[z]["ColumnName"] != System.DBNull.Value && oSourceTableSchema.Rows[z]["ColumnName"].ToString().Trim().ToUpper() != "SCENARIO_ID")
                    {
                        bFound = false;
                        for (zz = 0; zz <= strDestColumnsArray.Length - 1; zz++)
                        {
                            if (oSourceTableSchema.Rows[z]["ColumnName"].ToString().Trim().ToUpper() ==
                                strDestColumnsArray[zz].Trim().ToUpper())
                            {
                                bFound = true;
                                break;
                            }
                        }
                        if (!bFound)
                        {
                            //column not found so let's add it
                            dataMgr.m_strSQL = dataMgr.FormatCreateTableSqlFieldItem(oSourceTableSchema.Rows[z]);
                            dataMgr.SqlNonQuery(oConn, "ALTER TABLE " + m_strAddHarvCostsWorkTable + " " +
                                "ADD COLUMN " + dataMgr.m_strSQL);
                        }
                    }
                }
                //
                //append all the current scenario rows into the work table
                //
                strDestColumnsList = dataMgr.getFieldNames(oConn, "SELECT * FROM " + m_strAddHarvCostsWorkTable);
                dataMgr.m_strSQL = "INSERT INTO " + m_strAddHarvCostsWorkTable + " (" + strDestColumnsList + ") " +
                                    "SELECT " + strDestColumnsList + " FROM scenario_additional_harvest_costs WHERE UPPER(TRIM(scenario_id))='" + this.ScenarioId.Trim().ToUpper() + "'";
                dataMgr.SqlNonQuery(oConn, dataMgr.m_strSQL);

                //
                //GET ALL BIOSUM_COND_ID + RX possible combinations from tree and rx tables
                //
                //attach master.db for access to rx table
                dataMgr.m_strSQL = $@"ATTACH '{ this.ReferenceProcessorScenarioForm.LoadedQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Rx)}' AS MASTER";
                dataMgr.SqlNonQuery(oConn, dataMgr.m_strSQL);
                dataMgr.m_strSQL = "CREATE TABLE plotrx AS " +
                    "SELECT b.biosum_cond_id, a.rx " +
                    "FROM " + this.ReferenceProcessorScenarioForm.LoadedQueries.m_oFvs.m_strRxTable + " a " +
                    "CROSS JOIN (SELECT DISTINCT biosum_cond_id FROM " + this.ReferenceProcessorScenarioForm.LoadedQueries.m_oFIAPlot.m_strTreeTable + ") b";
                dataMgr.SqlNonQuery(oConn, dataMgr.m_strSQL);
                dataMgr.AddIndex(oConn, "plotrx", "plotrx_idx", "biosum_cond_id,rx");

                dataMgr.m_strSQL = "CREATE TABLE NewPlotRx AS " +
                  "SELECT a.biosum_cond_id, a.rx FROM PlotRx a " +
                  "LEFT JOIN additional_harvest_costs_work_table b on " +
                  "a.biosum_cond_id = b.biosum_cond_id AND a.rx = b.rx " +
                  "WHERE b.biosum_cond_id IS NULL";
                dataMgr.SqlNonQuery(oConn, dataMgr.m_strSQL);
                this.lblNewlyAdded.Text = Convert.ToString(dataMgr.getRecordCount(oConn, "SELECT COUNT(*) FROM NewPlotRx", "NewPlotRx"));

                //
                //insert new plot + rx combos
                //
                if (Convert.ToInt32(this.lblNewlyAdded.Text) > 0)
                {
                    dataMgr.m_strSQL = "INSERT INTO " + m_strAddHarvCostsWorkTable + " (biosum_cond_id,rx) " +
                                        "SELECT biosum_cond_id,rx FROM NewPlotRx";
                    dataMgr.SqlNonQuery(oConn, dataMgr.m_strSQL);

                    this.lblNewlyAdded.Visible = true;
                    this.lblNewAddedDescription.Visible = true;
                    this.ReferenceProcessorScenarioForm.m_bSave = true;
                }


            }

            //
            //load up any scenario columns and the default values
            //
            ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadHarvestCostComponents(
                    strScenarioDB, ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);
                if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Count > 0)
                {
                    for (x = 0; x <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Count - 1; x++)
                    {
                        ProcessorScenarioItem.HarvestCostItem oHarvestCostItem =
                            ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Item(x);
                        if (oHarvestCostItem.Type == "S")
                        {
                            if (this.uc_collection.Count == 0)
                            {
                                uc_processor_scenario_additional_harvest_cost_column_item1.ColumnName = oHarvestCostItem.ColumnName;
                                uc_processor_scenario_additional_harvest_cost_column_item1.Name = "uc_processor_scenario_additional_harvest_cost_column_item" + Convert.ToString(uc_collection.Count + 1);
                                uc_processor_scenario_additional_harvest_cost_column_item1.Type = "Scenario";
                                uc_processor_scenario_additional_harvest_cost_column_item1.Description = oHarvestCostItem.Description;
                                uc_processor_scenario_additional_harvest_cost_column_item1.DefaultCubicFootDollarValue = oHarvestCostItem.DefaultCostPerAcre;
                                uc_processor_scenario_additional_harvest_cost_column_item1.EnableColumnNameRemoveButton = true;
                                uc_processor_scenario_additional_harvest_cost_column_item1.ReferenceAdditionalHarvestCostColumnsUserControl = this;
                                uc_processor_scenario_additional_harvest_cost_column_item1.Visible = true;
                                uc_processor_scenario_additional_harvest_cost_column_item1.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
                                uc_collection.Add(this.uc_processor_scenario_additional_harvest_cost_column_item1);
                            }
                            else
                            {
                                //create another harvest cost item
                                FIA_Biosum_Manager.uc_processor_scenario_additional_harvest_cost_column_item oItem = new uc_processor_scenario_additional_harvest_cost_column_item();
                                oItem.ColumnName = oHarvestCostItem.ColumnName;
                                oItem.Type = "Scenario";
                                oItem.Description = oHarvestCostItem.Description;
                                oItem.EnableColumnNameRemoveButton = true;
                                oItem.ReferenceAdditionalHarvestCostColumnsUserControl = this;
                                oItem.DefaultCubicFootDollarValue = oHarvestCostItem.DefaultCostPerAcre;
                                oItem.Name = "uc_processor_scenario_additional_harvest_cost_column_item" + Convert.ToString(uc_collection.Count + 1);
                                oItem.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
                                panel1.Controls.Add(oItem);
                                oItem.Left = this.uc_processor_scenario_additional_harvest_cost_column_item1.Left;
                                oItem.Top = uc_collection.Item(uc_collection.Count - 1).Top + oItem.Height;
                                oItem.Visible = true;
                                uc_collection.Add(oItem);
                            }
                        }
                        else
                        {
                            //find the column and rx
                            for (y = 0; y <= uc_collection.Count - 1; y++)
                            {
                                if (oHarvestCostItem.ColumnName.Trim().ToUpper() == uc_collection.Item(y).ColumnName.Trim().ToUpper() &&
                                    oHarvestCostItem.Rx.Trim() == uc_collection.Item(y).Type.Trim())
                                {
                                    uc_collection.Item(y).DefaultCubicFootDollarValue = oHarvestCostItem.DefaultCostPerAcre;
                                    break;
                                }
                            }
                        }
                    }
                }

            if (dataMgr.m_intError == 0)
            {
                //
                //make sure columns exist in the additional harvest cost columns table
                //
                CreateColumns(dataMgr, strConn);
                //
                //get null counts for each column
                //
                UpdateNullCounts(TempDb);
                if (this.uc_collection.Count == 0)
                    uc_processor_scenario_additional_harvest_cost_column_item1.Visible = false;
            }
        }

        public void  DeleteWorkTable()
        {
                SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
                string strScenarioDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                string strConn = dataMgr.GetConnectionString(strScenarioDB);
                using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    oConn.Open();
                    //delete work table if it exists
                    if (dataMgr.TableExist(oConn, m_strAddHarvCostsWorkTable))
                    {
                        dataMgr.m_strSQL = "DROP TABLE " + m_strAddHarvCostsWorkTable;
                        dataMgr.SqlNonQuery(oConn, dataMgr.m_strSQL);
                    }
                }
        }
        public void loadvaluesFromProperties()
        {
            int x;
            int y;
            int z;
            int zz;
            bool bFound = false;

            if (m_oAdo != null && m_oAdo.m_OleDbConnection != null)
                if (m_oAdo.m_OleDbConnection.State == System.Data.ConnectionState.Open) m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);

            //
            //SCENARIO MDB
            //
            string strScenarioDB =
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            //
            //SCENARIO ID
            //
            ScenarioId = this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();

            //
            //OPEN CONNECTION TO TEMP DB FILE
            //
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(this.ReferenceProcessorScenarioForm.LoadedQueries.m_strTempDbFile, "", ""));
            if (m_oAdo.m_intError == 0)
            {
                //
                //OPEN CONNECTION TO SCENARIO DEFINITIONS
                //
                SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
                string strConn = dataMgr.GetConnectionString(strScenarioDB);
                dataMgr.OpenConnection(strConn);
                //delete work table if it exists
                if (dataMgr.TableExist(dataMgr.m_Connection, m_strAddHarvCostsWorkTable))
                {
                    dataMgr.m_strSQL = "DROP TABLE " + m_strAddHarvCostsWorkTable;
                    dataMgr.SqlNonQuery(dataMgr.m_Connection, dataMgr.m_strSQL);
                }
                //create a work table from the additional harvests costs table
                //
                //@ToDo: Switch to SQLite datatypes when ready
                //dataMgr.m_strSQL = Tables.Processor.CreateSqliteAdditionalHarvestCostsTableSQL(m_strAddHarvCostsWorkTable);
                dataMgr.m_strSQL = Tables.Processor.CreateAdditionalHarvestCostsTableSQL(m_strAddHarvCostsWorkTable);

                dataMgr.SqlNonQuery(dataMgr.m_Connection, dataMgr.m_strSQL);
                //
                //add columns from the SOURCE scenario additional harvest costs table to the DESTINATION work table
                //
                DataTable oSourceTableSchema = dataMgr.getTableSchema(dataMgr.m_Connection, "SELECT * FROM scenario_additional_harvest_costs");
                string strSourceColumnsList = dataMgr.getFieldNames(dataMgr.m_Connection, "SELECT * FROM scenario_additional_harvest_costs");
                string strDestColumnsList = dataMgr.getFieldNames(dataMgr.m_Connection, "SELECT * FROM " + m_strAddHarvCostsWorkTable);
                string[] strDestColumnsArray = frmMain.g_oUtils.ConvertListToArray(strDestColumnsList, ",");
                dataMgr.m_strSQL = "";
                for (z = 0; z <= oSourceTableSchema.Rows.Count - 1; z++)
                {
                    if (oSourceTableSchema.Rows[z]["ColumnName"] != System.DBNull.Value && oSourceTableSchema.Rows[z]["ColumnName"].ToString().Trim().ToUpper() != "SCENARIO_ID")
                    {
                        bFound = false;
                        for (zz = 0; zz <= strDestColumnsArray.Length - 1; zz++)
                        {
                          if (oSourceTableSchema.Rows[z]["ColumnName"].ToString().Trim().ToUpper() ==
                              strDestColumnsArray[zz].Trim().ToUpper())
                          {
                                bFound = true;
                                break;
                          }
                        }
                        if (!bFound)
                        {
                            //column not found so let's add it
                            dataMgr.m_strSQL = dataMgr.FormatCreateTableSqlFieldItem(oSourceTableSchema.Rows[z]);
                            dataMgr.SqlNonQuery(dataMgr.m_Connection, "ALTER TABLE " + m_strAddHarvCostsWorkTable + " " +
                                "ADD COLUMN " + dataMgr.m_strSQL);
                        }
                    }
                }
                //
                //append all the current scenario rows into the work table
                //
                strDestColumnsList = dataMgr.getFieldNames(dataMgr.m_Connection, "SELECT * FROM " + m_strAddHarvCostsWorkTable);
                dataMgr.m_strSQL = "INSERT INTO " + m_strAddHarvCostsWorkTable + " (" + strDestColumnsList + ") " +
                                    "SELECT " + strDestColumnsList + " FROM scenario_additional_harvest_costs WHERE UPPER(TRIM(scenario_id))='" + this.ScenarioId.Trim().ToUpper() + "'";
                dataMgr.SqlNonQuery(dataMgr.m_Connection, dataMgr.m_strSQL);

                //CREATE LINK IN TEMP MDB TO TEMPORARY SCENARIO ADDITIONAL HARVEST COSTS TABLE
                //
                ODBCMgr odbcmgr = new ODBCMgr();
                // Check to see if the input SQLite DSN exists and if so, delete so we can add
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName);
                }
                odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName,
                    this.ReferenceProcessorScenarioForm.LoadedQueries.m_strTempDbFile);
                if (! m_oAdo.TableExist(m_oAdo.m_OleDbConnection, m_strAddHarvCostsWorkTable))
                {
                    dao_data_access oDao = new dao_data_access();
                    oDao.CreateSQLiteTableLink(this.ReferenceProcessorScenarioForm.LoadedQueries.m_strTempDbFile, m_strAddHarvCostsWorkTable,
                        m_strAddHarvCostsWorkTable, ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName, strScenarioDB);
                    oDao.m_DaoWorkspace.Close();
                    oDao = null;
                    frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioAdditionalHarvestCostsTableIndexes(m_oAdo,
                        m_oAdo.m_OleDbConnection, m_strAddHarvCostsWorkTable);
                }

                //
                //DON'T append all the current scenario rows into the work table
                //This is a copy. We want to start fresh with none of the old harvest costs
                //


                // DELETE OLD TABLE LINKS IF THEY EXIST IN TEMP DB
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "plotrx"))
                {
                    m_oAdo.m_strSQL = "DROP TABLE " + "plotrx";
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                }
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "NewPlotRx"))
                {
                    m_oAdo.m_strSQL = "DROP TABLE " + "NewPlotRx";
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                }
                //
                //GET ALL BIOSUM_COND_ID + RX possible combinations from tree and rx tables
                //
                m_oAdo.m_strSQL = "SELECT b.biosum_cond_id,a.rx " +
                                  "INTO plotrx " +
                                  "FROM " + this.ReferenceProcessorScenarioForm.LoadedQueries.m_oFvs.m_strRxTable + " a," +
                                    "(SELECT DISTINCT biosum_cond_id " +
                                     "FROM " + this.ReferenceProcessorScenarioForm.LoadedQueries.m_oFIAPlot.m_strTreeTable + ") b";

                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                m_oAdo.AddIndex(m_oAdo.m_OleDbConnection, "plotrx", "plotrx_idx", "biosum_cond_id,rx");
                m_oAdo.m_strSQL = "SELECT a.biosum_cond_id, a.rx " +
                                  "INTO NewPlotRx FROM PlotRx a " +
                                  "LEFT JOIN additional_harvest_costs_work_table b on " +
                                  "a.biosum_cond_id = b.biosum_cond_id AND a.rx = b.rx " +
                                  "WHERE b.biosum_cond_id IS NULL";
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

                this.lblNewlyAdded.Text = Convert.ToString(m_oAdo.getRecordCount(m_oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM NewPlotRx", "NewPlotRx"));
                //
                //insert new plot + rx combos
                //
                if (Convert.ToInt32(this.lblNewlyAdded.Text) > 0)
                {
                    m_oAdo.m_strSQL = "INSERT INTO " + this.m_strAddHarvCostsWorkTable + "  (biosum_cond_id,rx) " +
                                        "SELECT biosum_cond_id,rx FROM NewPlotRx";
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

                    this.lblNewlyAdded.Visible = true;
                    this.lblNewAddedDescription.Visible = true;
                    this.ReferenceProcessorScenarioForm.m_bSave = true;
                }

                //
                //Remove any existing scenario items before adding current
                //
                List<int> lstItemsToRemove = new List<int>();
                int j = 0;
                foreach (FIA_Biosum_Manager.uc_processor_scenario_additional_harvest_cost_column_item oItem in this.uc_collection)
                {
                    if (oItem.Type.Equals("Scenario".Trim()))
                        lstItemsToRemove.Add(j);
                    j++;
                }
                foreach (int k in lstItemsToRemove)
                {
                    this.uc_collection.Remove(k);
                }

                //
                //load up any scenario columns and the default values
                //
                if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Count > 0)
                {
                    for (x = 0; x <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Count - 1; x++)
                    {
                        ProcessorScenarioItem.HarvestCostItem oHarvestCostItem =
                            ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Item(x);
                        if (oHarvestCostItem.Type == "S")
                        {
                            if (this.uc_collection.Count == 0)
                            {
                                uc_processor_scenario_additional_harvest_cost_column_item1.ColumnName = oHarvestCostItem.ColumnName;
                                uc_processor_scenario_additional_harvest_cost_column_item1.Name = "uc_processor_scenario_additional_harvest_cost_column_item" + Convert.ToString(uc_collection.Count + 1);
                                uc_processor_scenario_additional_harvest_cost_column_item1.Type = "Scenario";
                                uc_processor_scenario_additional_harvest_cost_column_item1.Description = oHarvestCostItem.Description;
                                uc_processor_scenario_additional_harvest_cost_column_item1.DefaultCubicFootDollarValue = oHarvestCostItem.DefaultCostPerAcre;
                                uc_processor_scenario_additional_harvest_cost_column_item1.EnableColumnNameRemoveButton = true;
                                uc_processor_scenario_additional_harvest_cost_column_item1.ReferenceAdditionalHarvestCostColumnsUserControl = this;
                                uc_processor_scenario_additional_harvest_cost_column_item1.Visible = true;
                                uc_processor_scenario_additional_harvest_cost_column_item1.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
                                uc_collection.Add(this.uc_processor_scenario_additional_harvest_cost_column_item1);
                            }
                            else
                            {
                                //create another harvest cost item
                                FIA_Biosum_Manager.uc_processor_scenario_additional_harvest_cost_column_item oItem = new uc_processor_scenario_additional_harvest_cost_column_item();
                                oItem.ColumnName = oHarvestCostItem.ColumnName;
                                oItem.Type = "Scenario";
                                oItem.Description = oHarvestCostItem.Description;
                                oItem.EnableColumnNameRemoveButton = true;
                                oItem.ReferenceAdditionalHarvestCostColumnsUserControl = this;
                                oItem.DefaultCubicFootDollarValue = oHarvestCostItem.DefaultCostPerAcre;
                                oItem.Name = "uc_processor_scenario_additional_harvest_cost_column_item" + Convert.ToString(uc_collection.Count + 1);
                                oItem.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
                                panel1.Controls.Add(oItem);
                                oItem.Left = this.uc_processor_scenario_additional_harvest_cost_column_item1.Left;
                                oItem.Top = uc_collection.Item(uc_collection.Count - 1).Top + oItem.Height;
                                oItem.Visible = true;
                                uc_collection.Add(oItem);
                            }
                        }
                        else
                        {
                            //find the column and rx
                            for (y = 0; y <= uc_collection.Count - 1; y++)
                            {
                                if (oHarvestCostItem.ColumnName.Trim().ToUpper() == uc_collection.Item(y).ColumnName.Trim().ToUpper() &&
                                    oHarvestCostItem.Rx.Trim() == uc_collection.Item(y).Type.Trim())
                                {
                                    uc_collection.Item(y).DefaultCubicFootDollarValue = oHarvestCostItem.DefaultCostPerAcre;
                                    break;
                                }
                            }
                        }

                    }
                }
                //
                //make sure columns exist in the additional harvest cost columns table
                //
                CreateColumns(dataMgr, dataMgr.m_Connection.ConnectionString);

                // Copy values from reference to target scenario
                dataMgr.m_strSQL = "";
                string strColumn = "";

                for (int q = 0; q <= uc_collection.Count - 1; q++)
                {
                    strColumn = uc_collection.Item(q).ColumnName.Trim();
                    //make sure the source scenario has this column
                    if (dataMgr.ColumnExist(dataMgr.m_Connection, "scenario_additional_harvest_costs", strColumn))
                    {
                        //make sure columnname not already referenced
                        if (dataMgr.m_strSQL.ToUpper().IndexOf("B." + strColumn.ToUpper() + " IS NOT NULL", 0) < 0)
                        {
                            dataMgr.m_strSQL = dataMgr.m_strSQL + strColumn + " = (SELECT COALESCE(b." + strColumn + ", a." + strColumn + ") " +
                                               "FROM scenario_additional_harvest_costs as b " +
                                               "WHERE a.biosum_cond_id = b.biosum_cond_id AND a.rx = b.rx), ";
                        }


                    }
                }
                if (dataMgr.m_strSQL.Trim().Length > 0)
                {
                    frmMain.g_oFrmMain.ActivateStandByAnimation(
                           this.ParentForm.WindowState,
                           this.ParentForm.Left,
                           this.ParentForm.Height,
                           this.ParentForm.Width,
                           this.ParentForm.Top);
                    //remove the comma at the end of the string
                    dataMgr.m_strSQL = dataMgr.m_strSQL.Substring(0, dataMgr.m_strSQL.Length - 2);

                    String sourceScenarioId = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.SourceScenarioId;
                    dataMgr.m_strSQL = "UPDATE additional_harvest_costs_work_table AS a SET " + dataMgr.m_strSQL +
                                       " WHERE EXISTS( " +
                                       "SELECT* FROM scenario_additional_harvest_costs AS b " +
                                       "WHERE b.biosum_cond_id = a.biosum_cond_id AND b.RX = a.RX " +
                                       "AND upper(trim(b.scenario_id)) = '" + sourceScenarioId.Trim().ToUpper() + "')";
                    frmMain.g_sbpInfo.Text = "Updating Harvest Cost Component $/A/C Values...Stand By";

                    dataMgr.SqlNonQuery(dataMgr.m_Connection, dataMgr.m_strSQL);

                    frmMain.g_sbpInfo.Text = "Ready";

                    if (dataMgr.m_intError == 0)
                    {
                        UpdateNullCounts(TempDb);
                        frmMain.g_oFrmMain.DeactivateStandByAnimation();
                        //MessageBox.Show("Done");
                    }
                    else
                        frmMain.g_oFrmMain.DeactivateStandByAnimation();
                }
                dataMgr.CloseAndDisposeConnection(dataMgr.m_Connection, true);

                //
                //get null counts for each column
                //
                UpdateNullCounts(TempDb);
                if (this.uc_collection.Count == 0)
                    uc_processor_scenario_additional_harvest_cost_column_item1.Visible = false;

            }
        }
        private void PositionControls()
        {
            for (int x = 0; x <= uc_collection.Count - 1; x++)
            {
                if (x != 0)
                {
                    uc_collection.Item(x).Top = uc_collection.Item(x - 1).Top + uc_collection.Item(x).Height;
                    //uc_collection.Item(x).ShowSeparator = true;
                }
                //else uc_collection.Item(x).ShowSeparator = false;

            }
        }
        public void savevalues()
        {
            int z;
            int zz;
            bool bFound = false;
            string strFields = "";
            string strValues = "";
            m_strError = "";
            m_intError = 0;
            try
            {
                //
                //SCENARIO MDB
                //
                string strScenarioMDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\db\\scenario_processor_rule_definitions.mdb";
                System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();

                m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strScenarioMDB, "", ""), ref oConn);

                //
                //add columns from the SOURCE additional harvest costs work table to the DESTINATION scenario table
                //
                System.Data.DataTable oSourceTableSchema = m_oAdo.getTableSchema(m_oAdo.m_OleDbConnection, "SELECT * FROM additional_harvest_costs_work_table");
                string strSourceColumnsList = m_oAdo.getFieldNames(m_oAdo.m_OleDbConnection, "SELECT * FROM additional_harvest_costs_work_table");
                string strSourceColumnsReservedWordFormattedList = m_oAdo.FormatReservedWordsInColumnNameList(strSourceColumnsList, ",");
                string[] strSourceColumnsArray = frmMain.g_oUtils.ConvertListToArray(strSourceColumnsList, ",");
                string strDestColumnsList = m_oAdo.getFieldNames(m_oAdo.m_OleDbConnection, "SELECT * FROM scenario_additional_harvest_costs");
                string[] strDestColumnsArray = frmMain.g_oUtils.ConvertListToArray(strDestColumnsList, ",");

                m_oAdo.m_strSQL = "";
                for (z = 0; z <= oSourceTableSchema.Rows.Count - 1; z++)
                {

                    if (oSourceTableSchema.Rows[z]["ColumnName"] != System.DBNull.Value && oSourceTableSchema.Rows[z]["ColumnName"].ToString().Trim().ToUpper() != "SCENARIO_ID")
                    {
                        bFound = false;
                        for (zz = 0; zz <= strDestColumnsArray.Length - 1; zz++)
                        {
                            if (oSourceTableSchema.Rows[z]["ColumnName"].ToString().Trim().ToUpper() ==
                                strDestColumnsArray[zz].Trim().ToUpper())
                            {
                                bFound = true;
                                break;
                            }
                        }
                        if (!bFound)
                        {
                            //column not found so let's add it
                            m_oAdo.m_strSQL = m_oAdo.FormatCreateTableSqlFieldItem(oSourceTableSchema.Rows[z]);
                            m_oAdo.SqlNonQuery(oConn, "ALTER TABLE scenario_additional_harvest_costs " +
                                "ADD COLUMN " + m_oAdo.m_strSQL);
                        }
                    }
                }
                //
                //delete all records of the current scenario
                //
                m_oAdo.m_strSQL = "DELETE FROM scenario_additional_harvest_costs WHERE UCASE(TRIM(scenario_id))='" + ScenarioId + "'";
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                //
                //append all the current scenario rows into the work table
                //
                strDestColumnsList = m_oAdo.getFieldNames(m_oAdo.m_OleDbConnection, "SELECT * FROM additional_harvest_costs_work_table");
                m_oAdo.m_strSQL = "INSERT INTO scenario_additional_harvest_costs (scenario_id," + strDestColumnsList + ") " +
                                    "SELECT '" + ScenarioId + "'," + strDestColumnsList + " FROM additional_harvest_costs_work_table";
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                //
                //save scenario column information
                //
                //
                //delete all records of the current scenario
                //
                m_oAdo.m_strSQL = "DELETE FROM scenario_harvest_cost_columns WHERE UCASE(TRIM(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'";
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                //
                //insert all the scenario harvest cost columns into the scenario_harvest_cost_columns table
                //
                strFields = "scenario_id,ColumnName,Description,rx,Default_CPA";
                for (z = 0; z <= this.uc_collection.Count - 1; z++)
                {
                    strValues = "'" + ScenarioId + "',";
                    strValues = strValues + "'" + uc_collection.Item(z).ColumnName.Trim() + "',";
                    if (uc_collection.Item(z).Type.Trim().ToUpper() == "SCENARIO")
                    {

                        strValues = strValues + "'" + uc_collection.Item(z).Description.Trim() + "',";
                        strValues = strValues + "'',";
                    }
                    else
                    {
                        strValues = strValues + "'',";
                        strValues = strValues + "'" + uc_collection.Item(z).Type.Trim() + "',";
                    }
                    strValues = strValues + uc_collection.Item(z).DefaultCubicFootDollarValue.Replace("$", "");

                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, Queries.GetInsertSQL(strFields, strValues, "scenario_harvest_cost_columns"));
                    //
                    //update the rx_harvest_cost_columns description field
                    //
                    if (uc_collection.Item(z).Type.Trim().ToUpper() != "SCENARIO")
                    {
                        m_oAdo.m_strSQL = "UPDATE " + this.ReferenceProcessorScenarioForm.LoadedQueries.m_oFvs.m_strRxHarvestCostColumnsTable + " " +
                                           "SET Description='" + uc_collection.Item(z).Description + "'" +
                                           "WHERE rx='" + uc_collection.Item(z).Type.Trim() + "' AND " +
                                                 "TRIM(ColumnName)='" + uc_collection.Item(z).ColumnName.Trim() + "'";
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    }
                }
                m_oAdo.CloseConnection(oConn);
            }
            catch (Exception e)
            {
                
                m_intError = -1;
                m_strError = e.Message;
            }
        }
        private void CreateColumns(SQLite.ADO.DataMgr p_DataMgr, string p_strConn)
        {
            int x,y;
            string strColumnsAddedList = "";
            string strColumnName = "";
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(p_strConn))
            {
                oConn.Open();
                string[] strColumnsArray = p_DataMgr.getFieldNamesArray(oConn, "SELECT * FROM additional_harvest_costs_work_table ");
                for (y = 0; y <= uc_collection.Count - 1; y++)
                {
                    if (!p_DataMgr.ColumnExist(p_DataMgr.m_Connection, "additional_harvest_costs_work_table", uc_collection.Item(y).ColumnName.Trim()))
                    {
                        if (strColumnsAddedList.IndexOf("," + uc_collection.Item(y).ColumnName.Trim().ToUpper() + ",", 0) < 0)
                        {
                            for (x = 0; x <= strColumnsArray.Length - 1; x++)
                            {
                                strColumnName = strColumnsArray[x].Trim();
                                if (uc_collection.Item(y).ColumnName.Trim().ToUpper() ==
                                            strColumnName.ToUpper()) break;


                            }
                            if (x > strColumnsArray.Length - 1)
                            {
                                p_DataMgr.m_strSQL = "ALTER TABLE additional_harvest_costs_work_table  " +
                                                  "ADD COLUMN " + uc_collection.Item(y).ColumnName.Trim() + " DOUBLE ";
                                p_DataMgr.SqlNonQuery(p_DataMgr.m_Connection, p_DataMgr.m_strSQL);
                                strColumnsAddedList = strColumnsAddedList + "," + strColumnName + ",";
                            }
                        }
                    }
                }
            }
        }

        public void savevaluesSqlite()
        {
            int z;
            int zz;
            bool bFound = false;
            string strFields = "";
            string strValues = "";
            m_strError = "";
            m_intError = 0;
            try
            {
                //
                //SCENARIO MDB
                //
                string strScenarioDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
                oDataMgr = new SQLite.ADO.DataMgr();
                oDataMgr.OpenConnection(oDataMgr.GetConnectionString(strScenarioDB));
                if (oDataMgr.m_intError != 0)
                {
                    m_intError = oDataMgr.m_intError;
                    m_strError = oDataMgr.m_strError;
                    oDataMgr = null;
                    return;
                }

                //
                //add columns from the SOURCE additional harvest costs work table to the DESTINATION scenario table
                //
                DataTable oSourceTableSchema = oDataMgr.getTableSchema(oDataMgr.m_Connection, "SELECT * FROM " + m_strAddHarvCostsWorkTable);
                string strSourceColumnsList = oDataMgr.getFieldNames(oDataMgr.m_Connection, "SELECT * FROM " + m_strAddHarvCostsWorkTable);
                string strSourceColumnsReservedWordFormattedList = oDataMgr.FormatReservedWordsInColumnNameList(strSourceColumnsList, ",");
                string[] strSourceColumnsArray = frmMain.g_oUtils.ConvertListToArray(strSourceColumnsList, ",");
                string strDestColumnsList = oDataMgr.getFieldNames(oDataMgr.m_Connection, "SELECT * FROM scenario_additional_harvest_costs");
                string[] strDestColumnsArray = frmMain.g_oUtils.ConvertListToArray(strDestColumnsList, ",");

                oDataMgr.m_strSQL = "";
                for (z = 0; z <= oSourceTableSchema.Rows.Count - 1; z++)
                {

                    if (oSourceTableSchema.Rows[z]["ColumnName"] != DBNull.Value && oSourceTableSchema.Rows[z]["ColumnName"].ToString().Trim().ToUpper() != "SCENARIO_ID")
                    {
                        bFound = false;
                        for (zz = 0; zz <= strDestColumnsArray.Length - 1; zz++)
                        {
                            if (oSourceTableSchema.Rows[z]["ColumnName"].ToString().Trim().ToUpper() ==
                                strDestColumnsArray[zz].Trim().ToUpper())
                            {
                                bFound = true;
                                break;
                            }
                        }
                        if (!bFound)
                        {
                            //column not found so let's add it
                            oDataMgr.m_strSQL = oDataMgr.FormatCreateTableSqlFieldItem(oSourceTableSchema.Rows[z]);
                            oDataMgr.SqlNonQuery(oDataMgr.m_Connection, "ALTER TABLE scenario_additional_harvest_costs " +
                                "ADD COLUMN " + oDataMgr.m_strSQL);
                        }
                    }
                }
                //
                //delete all records of the current scenario
                //
                oDataMgr.m_strSQL = "DELETE FROM scenario_additional_harvest_costs WHERE UPPER(TRIM(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'";
                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                //
                //append all the current scenario rows into the work table
                //
                strDestColumnsList = oDataMgr.getFieldNames(oDataMgr.m_Connection, "SELECT * FROM " + m_strAddHarvCostsWorkTable);
                oDataMgr.m_strSQL = "INSERT INTO scenario_additional_harvest_costs (scenario_id," + strDestColumnsList + ") " +
                                    "SELECT '" + ScenarioId + "'," + strDestColumnsList + " FROM " + m_strAddHarvCostsWorkTable;
                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                //
                //save scenario column information
                //
                //
                //delete all records of the current scenario
                //
                oDataMgr.m_strSQL = "DELETE FROM scenario_harvest_cost_columns WHERE UPPER(TRIM(scenario_id))='" + ScenarioId.Trim().ToUpper() + "'";
                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                //
                //insert all the scenario harvest cost columns into the scenario_harvest_cost_columns table
                //
                strFields = "scenario_id,ColumnName,Description,rx,Default_CPA";
                for (z = 0; z <= this.uc_collection.Count - 1; z++)
                {
                    strValues = "'" + ScenarioId + "',";
                    strValues = strValues + "'" + uc_collection.Item(z).ColumnName.Trim() + "',";
                    if (uc_collection.Item(z).Type.Trim().ToUpper() == "SCENARIO")
                    {

                        strValues = strValues + "'" + uc_collection.Item(z).Description.Trim() + "',";
                        strValues = strValues + "'',";
                    }
                    else
                    {
                        strValues = strValues + "'',";
                        strValues = strValues + "'" + uc_collection.Item(z).Type.Trim() + "',";
                    }
                    strValues = strValues + uc_collection.Item(z).DefaultCubicFootDollarValue.Replace("$", "");

                    oDataMgr.SqlNonQuery(oDataMgr.m_Connection, Queries.GetInsertSQL(strFields, strValues, "scenario_harvest_cost_columns"));
                    //
                    //update the rx_harvest_cost_columns description field
                    //
                    if (uc_collection.Item(z).Type.Trim().ToUpper() != "SCENARIO")
                    {
                        //
                        //db\master.db
                        //
                        string rxHarvestCostConn = ReferenceProcessorScenarioForm.LoadedQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.RxHarvestCostColumns);
                        using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(rxHarvestCostConn)))
                        {
                            conn.Open();
                            oDataMgr.m_strSQL = "UPDATE " + this.ReferenceProcessorScenarioForm.LoadedQueries.m_oFvs.m_strRxHarvestCostColumnsTable + " " +
                                           "SET Description='" + uc_collection.Item(z).Description + "'" +
                                           "WHERE rx='" + uc_collection.Item(z).Type.Trim() + "' AND " +
                                                 "TRIM(ColumnName)='" + uc_collection.Item(z).ColumnName.Trim() + "'";
                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                        }
                    }
                }
                oDataMgr.CloseConnection(oDataMgr.m_Connection);
                oDataMgr = null;
            }
            catch (Exception e)
            {

                m_intError = -1;
                m_strError = e.Message;
            }
        }
        public void UpdateNullCounts(string strTempDb)
        {
            int x;
            string strColumnName = "";
            string strRx = "";
            int intCount;

                SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
                string strConn = dataMgr.GetConnectionString(strTempDb);
                string addHarvCostsWorkTable = "additional_harvest_costs_work_table";
                using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    oConn.Open();
                    string[] strColumnsArray = dataMgr.getFieldNamesArray(oConn, "SELECT * FROM " + addHarvCostsWorkTable);
                    for (x = 0; x <= uc_collection.Count - 1; x++)
                    {
                        strColumnName = uc_collection.Item(x).ColumnName.Trim();
                        if (uc_collection.Item(x).Type != "Scenario")
                        {
                            strRx = uc_collection.Item(x).Type.Trim();
                            dataMgr.m_strSQL = "SELECT COUNT(*) AS null_count FROM " + addHarvCostsWorkTable +
                                              " WHERE TRIM(rx)='" + strRx + "' AND " + strColumnName + " IS NULL";

                        }
                        else
                        {
                            dataMgr.m_strSQL = "SELECT COUNT(*) AS null_count  FROM " + addHarvCostsWorkTable +
                                              " WHERE " + strColumnName + " IS NULL";
                        }
                        intCount = Convert.ToInt32(dataMgr.getSingleDoubleValueFromSQLQuery(oConn, dataMgr.m_strSQL,
                            this.ReferenceProcessorScenarioForm.LoadedQueries.m_oProcessor.m_strAdditionalHarvestCostsTable));
                        uc_collection.Item(x).NullCount = Convert.ToString(intCount);
                    }
                }

        }
        

        private void groupBox1_Resize(object sender, EventArgs e)
        {
            try
            {
                this.btnAdd.Top = this.ClientSize.Height - this.Top - this.btnAdd.Height - 5;
                this.btnEditAll.Top = this.btnAdd.Top;
                this.btnEditNull.Top = this.btnAdd.Top;
                this.btnRemoveAll.Top = this.btnAdd.Top;
                this.btnEditPrev.Top = this.btnAdd.Top;

                this.panel1.Height = this.btnAdd.Top - this.panel1.Top - 5;
                this.panel1.Width = this.Width - this.panel1.Left * 2;

            }
            catch
            {
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddColumn();
        }
        private void AddColumn()
        {

            int y,intCount;
            FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
            frmTemp.MaximizeBox = false;
            frmTemp.MinimizeBox = false;
            frmTemp.BackColor = System.Drawing.SystemColors.Control;
            frmTemp.Initialize_Scenario_Harvest_Costs_Column_Edit_Control();
            string strAvailableColumnList = "";
            string strUsedColumnList = "";

            //get all the columns in the table
            string strScenarioDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(strScenarioDB);
            string[] strColumnsArray = null;
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                strColumnsArray = dataMgr.getFieldNamesArray(oConn, "SELECT * FROM additional_harvest_costs_work_table ");
            }

            for (y = 0; y <= strColumnsArray.Length - 1; y++)
            {
                switch (strColumnsArray[y].Trim().ToUpper())
                {
                    case "BIOSUM_COND_ID":
                        break;
                    case "RX":
                        break;
                    default:
                        strAvailableColumnList = strAvailableColumnList + strColumnsArray[y].Trim() + ",";
                        break;
                }
            }
            if (strAvailableColumnList.Trim().Length > 0) strAvailableColumnList = strAvailableColumnList.Substring(0, strAvailableColumnList.Length - 1);



            frmTemp.Height = 0;
            frmTemp.Width = 0;
            if (frmTemp.Top + frmTemp.uc_scenario_harvest_cost_column_edit1.Height > frmTemp.ClientSize.Height + 2)
            {
                for (y = 1; ; y++)
                {
                    frmTemp.Height = y;
                    if (frmTemp.uc_scenario_harvest_cost_column_edit1.Top +
                        frmTemp.uc_scenario_harvest_cost_column_edit1.Height <
                        frmTemp.ClientSize.Height)
                    {
                        break;
                    }
                }

            }
            if (frmTemp.uc_scenario_harvest_cost_column_edit1.Left + frmTemp.uc_scenario_harvest_cost_column_edit1.Width > frmTemp.ClientSize.Width + 2)
            {
                for (y = 1; ; y++)
                {
                    frmTemp.Width = y;
                    if (frmTemp.uc_scenario_harvest_cost_column_edit1.Left +
                        frmTemp.uc_scenario_harvest_cost_column_edit1.Width <
                        frmTemp.ClientSize.Width)
                    {
                        break;
                    }
                }

            }
            frmTemp.Left = 0;
            frmTemp.Top = 0;

            frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            frmTemp.uc_scenario_harvest_cost_column_edit1.Dock = System.Windows.Forms.DockStyle.Fill;

            frmTemp.uc_scenario_harvest_cost_column_edit1.EditType = "New";


            frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnList = strAvailableColumnList;
            strUsedColumnList = "";
            for (y = 0; y <= uc_collection.Count - 1; y++)
            {
                strUsedColumnList = strUsedColumnList + this.uc_collection.Item(y).ColumnName + ",";
            }
            if (strUsedColumnList.Trim().Length > 0)
                strUsedColumnList = strUsedColumnList.Substring(0, strUsedColumnList.Length - 1);
            frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText = "";
            frmTemp.uc_scenario_harvest_cost_column_edit1.CurrentSelectedColumnList = strUsedColumnList;
            frmTemp.uc_scenario_harvest_cost_column_edit1.HarvestCostTableColumnList = strAvailableColumnList;
            frmTemp.uc_scenario_harvest_cost_column_edit1.loadvalues();
            frmTemp.uc_scenario_harvest_cost_column_edit1.lblEdit.Show();



            frmTemp.Text = "Additional CPA Component";
            System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                ReferenceProcessorScenarioForm.m_bSave = true;
                //check if column exists in the table
                for (y = 0; y <= strColumnsArray.Length - 1; y++)
                {
                    if (strColumnsArray[y].ToUpper().Trim() ==
                        frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText.ToUpper().Trim())
                        break;
                }
                //check if we need to add the column to the table
                using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    oConn.Open();
                    if (y > strColumnsArray.Length - 1)
                    {
                        dataMgr.m_strSQL = "ALTER TABLE additional_harvest_costs_work_table " +
                                          "ADD COLUMN " + frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText.Trim() + " " +
                                          "DOUBLE";
                        dataMgr.SqlNonQuery(oConn, dataMgr.m_strSQL);
                    }

                    //get the number of nulls in the table for this column
                    dataMgr.m_strSQL = "SELECT COUNT(*) FROM additional_harvest_costs_work_table WHERE " +
                                      frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText.Trim() + " " +
                                      "IS NULL";
                    intCount = Convert.ToInt32(dataMgr.getSingleDoubleValueFromSQLQuery(oConn, dataMgr.m_strSQL, "temp"));
                }

                if (uc_collection.Count == 0)
                {
                    uc_processor_scenario_additional_harvest_cost_column_item1.ColumnName = frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText.Trim();
                    uc_processor_scenario_additional_harvest_cost_column_item1.Name = "uc_processor_scenario_additional_harvest_cost_column_item" + Convert.ToString(uc_collection.Count + 1);
                    uc_processor_scenario_additional_harvest_cost_column_item1.Type = "Scenario";
                    uc_processor_scenario_additional_harvest_cost_column_item1.Description = frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnDescription.Trim(); ;
                    uc_processor_scenario_additional_harvest_cost_column_item1.EnableColumnNameRemoveButton = true;
                    //uc_processor_scenario_additional_harvest_cost_column_item1.ReferenceAdo = m_oAdo;
                    //uc_processor_scenario_additional_harvest_cost_column_item1.ReferenceOleDbConnection = m_oAdo.m_OleDbConnection;
                    uc_processor_scenario_additional_harvest_cost_column_item1.ReferenceAdditionalHarvestCostColumnsUserControl = this;
                    uc_processor_scenario_additional_harvest_cost_column_item1.NullCount = Convert.ToString(intCount);
                    uc_processor_scenario_additional_harvest_cost_column_item1.Visible = true;
                    uc_processor_scenario_additional_harvest_cost_column_item1.ReferenceProcessorScenarioForm = this.ReferenceProcessorScenarioForm;
                    uc_collection.Add(this.uc_processor_scenario_additional_harvest_cost_column_item1);
                    
                }
                else
                {
                    FIA_Biosum_Manager.uc_processor_scenario_additional_harvest_cost_column_item oItem = new uc_processor_scenario_additional_harvest_cost_column_item();
                    oItem.ColumnName = frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText.Trim();
                    oItem.Type = "Scenario";
                    oItem.Name = "uc_processor_scenario_additional_harvest_cost_column_item" + Convert.ToString(uc_collection.Count + 1);
                    oItem.Description = frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnDescription.Trim();
                    oItem.EnableColumnNameRemoveButton = true;
                    //oItem.ReferenceAdo = m_oAdo;
                    //oItem.ReferenceOleDbConnection = m_oAdo.m_OleDbConnection;
                    oItem.ReferenceAdditionalHarvestCostColumnsUserControl = this;
                    oItem.ReferenceProcessorScenarioForm = this.ReferenceProcessorScenarioForm;
                    panel1.Controls.Add(oItem);
                    oItem.Left = this.uc_processor_scenario_additional_harvest_cost_column_item1.Left;
                    oItem.Top = uc_collection.Item(uc_collection.Count - 1).Top + oItem.Height;
                    oItem.Visible = true;
                    oItem.NullCount = Convert.ToString(intCount);
                    uc_collection.Add(oItem);
                }
            }
            frmTemp.Dispose();
        }
        public void RemoveColumn(FIA_Biosum_Manager.uc_processor_scenario_additional_harvest_cost_column_item p_oItem)
        {
            for (int x = 0; x <= uc_collection.Count - 1; x++)
            {
                if (p_oItem.Name.Trim() == uc_collection.Item(x).Name.Trim())
                {
                    uc_collection.Remove(x);
                    if (x != 0)
                        p_oItem.Dispose();
                    else
                        p_oItem.Visible = false;
                    break;
                }
            }
            PositionControls();
        }

        private void btnRemoveAll_Click(object sender, EventArgs e)
        {
            
            uc_processor_scenario_additional_harvest_cost_column_item oItem;
            for (int x = uc_collection.Count - 1; x >=0 ; x--)
            {
                if (uc_collection.Item(x).Type.Trim().ToUpper()=="SCENARIO")
                {
                    ReferenceProcessorScenarioForm.m_bSave = true;
                    oItem=uc_collection.Item(x);
                    uc_collection.Remove(x);
                    if (x != 0)
                        oItem.Dispose();
                    else
                        oItem.Visible = false;
   
                }
            }
            
           
        }

        private void btnEditAll_Click(object sender, EventArgs e)
        {
            EditAll();
        }
        public void EditAll()
        {
            if (uc_collection.Count == 0) return;

            int x,y;
            string strColumnsToEditList="";
            string[] strColumnsToEditArray = null;
            string[] strAllColumnsArray=null;
            string strAllColumnsList = "";
            string str = "";
            /*****************************************************************
             **lets see if this harvest costs edit form is already open
             *****************************************************************/
            utils oUtils = new utils();
            oUtils.m_intLevel = 1;
            if (oUtils.FindWindowLike(frmMain.g_oFrmMain.Handle, "Processor: Edit Additional Harvest Costs " + " (" + ScenarioId + ")", "*", true, false) > 0)
            {
                MessageBox.Show("!!Harvest Costs Component Edit Form Is  Already Open!!", "Harvest Costs Componenet Edit Form", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                if (this.m_frmHarvestCosts.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    this.m_frmHarvestCosts.WindowState = System.Windows.Forms.FormWindowState.Normal;
                this.m_frmHarvestCosts.Focus();
                return;
            }
            frmMain.g_oFrmMain.ActivateStandByAnimation(
                              this.ParentForm.WindowState,
                              this.ParentForm.Left,
                              this.ParentForm.Height,
                              this.ParentForm.Width,
                              this.ParentForm.Top);

            for (x = 0; x <= uc_collection.Count - 1; x++)
            {
                if (uc_collection.Item(x).ColumnName.Trim().Length > 0)
                {
                    if (str.IndexOf("," + uc_collection.Item(x).ColumnName.Trim().ToUpper() + ",") < 0)
                    {
                        strColumnsToEditList = strColumnsToEditList + uc_collection.Item(x).ColumnName.Trim() + ",";
                        str = str + "," + uc_collection.Item(x).ColumnName.Trim().ToUpper() + ",";
                    }
                }
            }
            if (strColumnsToEditList.Trim().Length > 0) strColumnsToEditList = strColumnsToEditList.Substring(0, strColumnsToEditList.Length - 1);
            strColumnsToEditArray = oUtils.ConvertListToArray(strColumnsToEditList,",");

            string strScenarioDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(strScenarioDB);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                strAllColumnsList = dataMgr.getFieldNames(oConn, "select * from additional_harvest_costs_work_table");
                strAllColumnsArray = oUtils.ConvertListToArray(strAllColumnsList, ",");
            }

            string strSQL = "";
            for (x = 0; x <= strAllColumnsArray.Length - 1; x++)
            {
                if (strAllColumnsArray[x].Trim().Length > 0)
                {
                    if (strAllColumnsArray[x].Trim().ToUpper() == "BIOSUM_COND_ID")
                    {
                        strSQL = "biosum_cond_id,rx,";
                        strSQL = strSQL + "substr(biosum_cond_id,6,2) as statecd,substr(biosum_cond_id,12,3) as countycd,substr(biosum_cond_id,15,7) as plot,substr(biosum_cond_id,25,1) as condid,";
                    }
                    else
                    {
                        for (y = 0; y <= strColumnsToEditArray.Length - 1; y++)
                        {

                            if (strAllColumnsArray[x].Trim().ToUpper() == strColumnsToEditArray[y].Trim().ToUpper())
                            {
                                strSQL = strSQL + strColumnsToEditArray[y].Trim() + ",";
                            }
                        }
                    }
                }
            }
            strSQL = strSQL.Substring(0, strSQL.Trim().Length - 1);

            strSQL = "SELECT DISTINCT " + strSQL + " FROM additional_harvest_costs_work_table";

            this.m_strColumnsToEdit = strColumnsToEditArray;
            this.m_intColumnsToEditCount = m_strColumnsToEdit.Length;

            string[] strRecordKeyField = new string[2];
            strRecordKeyField[0] = "biosum_cond_id";
            strRecordKeyField[1] = "rx";

            
            this.m_frmHarvestCosts = new frmGridView();
            this.m_frmHarvestCosts.HarvestCostColumns = true;
            this.m_frmHarvestCosts.UsingSQLite = true;
            this.m_frmHarvestCosts.ReferenceProcessorScenarioForm = this.ReferenceProcessorScenarioForm;
            this.m_frmHarvestCosts.LoadDataSetToEdit(strConn, strSQL, "additional_harvest_costs_work_table", this.m_strColumnsToEdit, this.m_intColumnsToEditCount, strRecordKeyField);
            frmMain.g_oFrmMain.DeactivateStandByAnimation();
            if (this.m_frmHarvestCosts.Visible == false)
            {
                m_frmHarvestCosts.MinimizeMainForm = true;
                m_frmHarvestCosts.ParentControl = this.ReferenceProcessorScenarioForm;
                m_frmHarvestCosts.ParentControl.Enabled = false;
                m_frmHarvestCosts.CallingClient = "ProcessorScenario:HarvesCostColumns_EditAll";
                m_frmHarvestCosts.ReferenceProcessorScenarioAdditionalHarvestCostColumns = this;

                this.m_frmHarvestCosts.Text = "Processor: Edit Additional Harvest Costs " + " (" + ScenarioId + ")";
               
                this.m_frmHarvestCosts.Show();
            }


        }
        public void EditAll(uc_processor_scenario_additional_harvest_cost_column_item p_oItem)
        {
            if (uc_collection.Count == 0) return;

            int x, y;
            string strColumnsToEditList = p_oItem.ColumnName.Trim() ;
            string[] strColumnsToEditArray = new string[1];
            strColumnsToEditArray[0] = p_oItem.ColumnName.Trim();
            string[] strAllColumnsArray = null;
            string strAllColumnsList = "";
            string strWhereSql = "";
            /*****************************************************************
             **lets see if this harvest costs edit form is already open
             *****************************************************************/
            utils oUtils = new utils();
            oUtils.m_intLevel = 1;
            if (oUtils.FindWindowLike(frmMain.g_oFrmMain.Handle, "Processor: Edit Additional Harvest Costs " + " (" + ScenarioId + ")", "*", true, false) > 0)
            {
                MessageBox.Show("!!Harvest Costs Component Edit Form Is  Already Open!!", "Harvest Costs Componenet Edit Form", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                if (this.m_frmHarvestCosts.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    this.m_frmHarvestCosts.WindowState = System.Windows.Forms.FormWindowState.Normal;
                this.m_frmHarvestCosts.Focus();
                return;
            }

            if (p_oItem.Type.Trim().ToUpper() != "SCENARIO")
            {
                strWhereSql = "WHERE rx = '" + p_oItem.Type.Trim() + "'";
            }

            string strScenarioDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(strScenarioDB);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                strAllColumnsList = dataMgr.getFieldNames(oConn, "select * from additional_harvest_costs_work_table");
                strAllColumnsArray = oUtils.ConvertListToArray(strAllColumnsList, ",");
            }

            string strSQL = "";
            for (x = 0; x <= strAllColumnsArray.Length - 1; x++)
            {
                if (strAllColumnsArray[x].Trim().Length > 0)
                {
                    if (strAllColumnsArray[x].Trim().ToUpper() == "BIOSUM_COND_ID")
                    {
                        strSQL = "biosum_cond_id,rx,";
                        strSQL = strSQL + "substr(biosum_cond_id,6,2) as statecd,substr(biosum_cond_id,12,3) as countycd,substr(biosum_cond_id,15,7) as plot,substr(biosum_cond_id,25,1) as condid,";
                    }
                    else
                    {
                        for (y = 0; y <= strColumnsToEditArray.Length - 1; y++)
                        {

                            if (strAllColumnsArray[x].Trim().ToUpper() == strColumnsToEditArray[y].Trim().ToUpper())
                            {
                                strSQL = strSQL + strColumnsToEditArray[y].Trim() + ",";
                            }
                        }
                    }
                }
            }
            strSQL = strSQL.Substring(0, strSQL.Trim().Length - 1);

            strSQL = "SELECT DISTINCT " + strSQL + " FROM additional_harvest_costs_work_table " + strWhereSql;

            this.m_strColumnsToEdit = strColumnsToEditArray;
            this.m_intColumnsToEditCount = m_strColumnsToEdit.Length;

            string[] strRecordKeyField = new string[2];
            strRecordKeyField[0] = "biosum_cond_id";
            strRecordKeyField[1] = "rx";

            this.m_frmHarvestCosts = new frmGridView();
            this.m_frmHarvestCosts.HarvestCostColumns = true;
            this.m_frmHarvestCosts.UsingSQLite = true;
            this.m_frmHarvestCosts.ReferenceProcessorScenarioForm = this.ReferenceProcessorScenarioForm;
            this.m_frmHarvestCosts.LoadDataSetToEdit(strConn, strSQL, "additional_harvest_costs_work_table", this.m_strColumnsToEdit, this.m_intColumnsToEditCount, strRecordKeyField);
            if (this.m_frmHarvestCosts.Visible == false)
            {
                this.m_frmHarvestCosts.Text = "Processor: Edit Additional Harvest Costs " + " (" + ScenarioId + ")";
                this.m_frmHarvestCosts.ShowDialog();
                this.UpdateNullCounts(TempDb);
            }
        }
        public void EditAllNulls(uc_processor_scenario_additional_harvest_cost_column_item p_oItem)
        {
            if (uc_collection.Count == 0) return;

            int x, y;
            string strColumnsToEditList = p_oItem.ColumnName.Trim();
            string[] strColumnsToEditArray = new string[1];
            strColumnsToEditArray[0] = p_oItem.ColumnName.Trim();
            string[] strAllColumnsArray = null;
            string strAllColumnsList = "";
            string strWhereSql = "";
            /*****************************************************************
             **lets see if this harvest costs edit form is already open
             *****************************************************************/
            utils oUtils = new utils();
            oUtils.m_intLevel = 1;
            if (oUtils.FindWindowLike(frmMain.g_oFrmMain.Handle, "Processor: Edit Additional Harvest Costs " + " (" + ScenarioId + ")", "*", true, false) > 0)
            {
                MessageBox.Show("!!Harvest Costs Component Edit Form Is  Already Open!!", "Harvest Costs Componenet Edit Form", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                if (this.m_frmHarvestCosts.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    this.m_frmHarvestCosts.WindowState = System.Windows.Forms.FormWindowState.Normal;
                this.m_frmHarvestCosts.Focus();
                return;
            }

            if (p_oItem.Type.Trim().ToUpper() != "SCENARIO")
            {
                strWhereSql = "WHERE rx = '" + p_oItem.Type.Trim() + "' AND " + p_oItem.ColumnName.Trim() + " IS NULL";
            }
            else
            {
                strWhereSql = "WHERE " +  p_oItem.ColumnName.Trim() + " IS NULL";
            }

            string strScenarioDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(strScenarioDB);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                strAllColumnsList = dataMgr.getFieldNames(oConn, "select * from additional_harvest_costs_work_table");
                strAllColumnsArray = oUtils.ConvertListToArray(strAllColumnsList, ",");
            }

            string strSQL = "";
            for (x = 0; x <= strAllColumnsArray.Length - 1; x++)
            {
                if (strAllColumnsArray[x].Trim().Length > 0)
                {
                    if (strAllColumnsArray[x].Trim().ToUpper() == "BIOSUM_COND_ID")
                    {
                        strSQL = "biosum_cond_id,rx,";
                        strSQL = strSQL + "substr(biosum_cond_id,6,2) as statecd,substr(biosum_cond_id,12,3) as countycd,substr(biosum_cond_id,15,7) as plot,substr(biosum_cond_id,25,1) as condid,";
                    }
                    else
                    {
                        for (y = 0; y <= strColumnsToEditArray.Length - 1; y++)
                        {

                            if (strAllColumnsArray[x].Trim().ToUpper() == strColumnsToEditArray[y].Trim().ToUpper())
                            {
                                strSQL = strSQL + strColumnsToEditArray[y].Trim() + ",";
                            }
                        }
                    }
                }
            }
            strSQL = strSQL.Substring(0, strSQL.Trim().Length - 1);

            strSQL = "SELECT DISTINCT " + strSQL + " FROM additional_harvest_costs_work_table " + strWhereSql;

            this.m_strColumnsToEdit = strColumnsToEditArray;
            this.m_intColumnsToEditCount = m_strColumnsToEdit.Length;

            string[] strRecordKeyField = new string[2];
            strRecordKeyField[0] = "biosum_cond_id";
            strRecordKeyField[1] = "rx";


            this.m_frmHarvestCosts = new frmGridView();
            this.m_frmHarvestCosts.HarvestCostColumns = true;
            this.m_frmHarvestCosts.UsingSQLite = true;
            this.m_frmHarvestCosts.ReferenceProcessorScenarioForm = this.ReferenceProcessorScenarioForm;
            this.m_frmHarvestCosts.LoadDataSetToEdit(strConn, strSQL, "additional_harvest_costs_work_table", this.m_strColumnsToEdit, this.m_intColumnsToEditCount, strRecordKeyField);
            if (this.m_frmHarvestCosts.Visible == false)
            {


                this.m_frmHarvestCosts.Text = "Processor: Edit Additional Harvest Costs " + " (" + ScenarioId + ")";
                this.m_frmHarvestCosts.ShowDialog();
                this.UpdateNullCounts(TempDb);
            }


        }
        private void EditAllNulls()
        {
            if (uc_collection.Count == 0) return;

            int x, y;
            string strColumnsToEditList = "";
            string[] strColumnsToEditArray = null;
            string[] strAllColumnsArray = null;
            string strAllColumnsList = "";
            string strWhereSql = "";
            string str = "";
            /*****************************************************************
             **lets see if this harvest costs edit form is already open
             *****************************************************************/
            utils oUtils = new utils();
            oUtils.m_intLevel = 1;
            if (oUtils.FindWindowLike(frmMain.g_oFrmMain.Handle, "Processor: Edit Additional Harvest Costs " + " (" + ScenarioId + ")", "*", true, false) > 0)
            {
                MessageBox.Show("!!Harvest Costs Component Edit Form Is  Already Open!!", "Harvest Costs Componenet Edit Form", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                if (this.m_frmHarvestCosts.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    this.m_frmHarvestCosts.WindowState = System.Windows.Forms.FormWindowState.Normal;
                this.m_frmHarvestCosts.Focus();
                return;
            }
            frmMain.g_oFrmMain.ActivateStandByAnimation(
                             this.ParentForm.WindowState,
                             this.ParentForm.Left,
                             this.ParentForm.Height,
                             this.ParentForm.Width,
                             this.ParentForm.Top);
            for (x = 0; x <= uc_collection.Count - 1; x++)
            {
                if (uc_collection.Item(x).ColumnName.Trim().Length > 0)
                {
                    if (str.IndexOf("," + uc_collection.Item(x).ColumnName.Trim().ToUpper() + ",") < 0)
                    {
                        strColumnsToEditList = strColumnsToEditList + uc_collection.Item(x).ColumnName.Trim() + ",";
                        str = str + "," + uc_collection.Item(x).ColumnName.Trim().ToUpper() + ",";
                        strWhereSql = strWhereSql + uc_collection.Item(x).ColumnName.Trim() + " IS NULL OR ";
                    }
                }
            }
            if (strColumnsToEditList.Trim().Length > 0) strColumnsToEditList = strColumnsToEditList.Substring(0, strColumnsToEditList.Length - 1);
            if (strWhereSql.Trim().Length > 0) strWhereSql = strWhereSql.Substring(0, strWhereSql.Length - 3);


            strColumnsToEditArray = oUtils.ConvertListToArray(strColumnsToEditList, ",");

            string strScenarioDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(strScenarioDB);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                strAllColumnsList = dataMgr.getFieldNames(oConn, "select * from additional_harvest_costs_work_table");
                strAllColumnsArray = oUtils.ConvertListToArray(strAllColumnsList, ",");
            }

            string strSQL = "";
            for (x = 0; x <= strAllColumnsArray.Length - 1; x++)
            {
                if (strAllColumnsArray[x].Trim().Length > 0)
                {
                    if (strAllColumnsArray[x].Trim().ToUpper() == "BIOSUM_COND_ID")
                    {
                        strSQL = "biosum_cond_id,rx,";
                        strSQL = strSQL + "substr(biosum_cond_id,6,2) as statecd,substr(biosum_cond_id,12,3) as countycd,substr(biosum_cond_id,15,7) as plot,substr(biosum_cond_id,25,1) as condid,";
                    }
                    else
                    {
                        for (y = 0; y <= strColumnsToEditArray.Length - 1; y++)
                        {

                            if (strAllColumnsArray[x].Trim().ToUpper() == strColumnsToEditArray[y].Trim().ToUpper())
                            {
                                strSQL = strSQL + strColumnsToEditArray[y].Trim() + ",";
                            }
                        }
                    }
                }
            }
            strSQL = strSQL.Substring(0, strSQL.Trim().Length - 1);

            strSQL = "SELECT DISTINCT " + strSQL + " FROM additional_harvest_costs_work_table WHERE " + strWhereSql;

            this.m_strColumnsToEdit = strColumnsToEditArray;
            this.m_intColumnsToEditCount = m_strColumnsToEdit.Length;

            string[] strRecordKeyField = new string[2];
            strRecordKeyField[0] = "biosum_cond_id";
            strRecordKeyField[1] = "rx";


            this.m_frmHarvestCosts = new frmGridView();
            this.m_frmHarvestCosts.HarvestCostColumns = true;
            this.m_frmHarvestCosts.UsingSQLite = true;
            this.m_frmHarvestCosts.ReferenceProcessorScenarioForm = this.ReferenceProcessorScenarioForm;
            this.m_frmHarvestCosts.LoadDataSetToEdit(strConn, strSQL, "additional_harvest_costs_work_table", this.m_strColumnsToEdit, this.m_intColumnsToEditCount, strRecordKeyField);
            frmMain.g_oFrmMain.DeactivateStandByAnimation();
            if (this.m_frmHarvestCosts.Visible == false)
            {
                m_frmHarvestCosts.MinimizeMainForm = true;
                m_frmHarvestCosts.ParentControl = this.ReferenceProcessorScenarioForm;
                m_frmHarvestCosts.ParentControl.Enabled = false;
                m_frmHarvestCosts.CallingClient = "ProcessorScenario:HarvesCostColumns_EditAll";
                m_frmHarvestCosts.ReferenceProcessorScenarioAdditionalHarvestCostColumns = this;


                this.m_frmHarvestCosts.Text = "Processor: Edit Additional Harvest Costs " + " (" + ScenarioId + ")";
                this.m_frmHarvestCosts.Show();

            }


        }

        private void btnEditNull_Click(object sender, EventArgs e)
        {
            EditAllNulls();
        }

        public void UpdateItemFromPrevious(uc_processor_scenario_additional_harvest_cost_column_item p_oItem)
        {
            string strColumn = "";
            string strRx = "";

            DialogResult result;

            frmMain.g_oFrmMain.ActivateStandByAnimation(
                   this.ParentForm.WindowState,
                   this.ParentForm.Left,
                   this.ParentForm.Height,
                   this.ParentForm.Width,
                   this.ParentForm.Top);

            strColumn = p_oItem.ColumnName.Trim();
            if (p_oItem.Type.Trim().ToUpper() != "SCENARIO")
                strRx = p_oItem.Type.Trim();

            frmDialog frmPrevExp = new frmDialog();

            frmPrevExp.Width = frmPrevExp.uc_previous_expressions1.m_intFullWd;
            frmPrevExp.Height = frmPrevExp.uc_previous_expressions1.m_intFullHt;
            frmPrevExp.Text = "Processor: Harvest Cost";

            frmPrevExp.uc_previous_expressions1.Visible = true;

            string strScenarioDB =
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(strScenarioDB);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                if (strRx.Trim().Length == 0)
                {
                    dataMgr.m_strSQL = "SELECT DISTINCT a.scenario_id, a.Description, b.Record_Count " +
                                      "FROM scenario a," +
                                        "(SELECT COUNT(*) AS Record_Count , scenario_id " +
                                         "FROM scenario_additional_harvest_costs GROUP BY scenario_id)  b " +
                                       "WHERE a.scenario_id=b.scenario_id AND a.scenario_id <> '" + ScenarioId + "'";
                }
                else
                {
                    dataMgr.m_strSQL = "SELECT DISTINCT a.scenario_id, a.Description, b.Record_Count " +
                                        "FROM scenario a," +
                                          "(SELECT COUNT(*) AS Record_Count , scenario_id " +
                                           "FROM scenario_additional_harvest_costs WHERE rx='" + strRx + "' GROUP BY scenario_id )  b " +
                                         "WHERE a.scenario_id=b.scenario_id AND a.scenario_id <> '" + ScenarioId + "'";
                }
                frmPrevExp.uc_previous_expressions1.lblTitle.Text = "Previous Scenario Harvest Cost Component Values";
                frmPrevExp.uc_previous_expressions1.loadvalues(dataMgr, oConn, dataMgr.m_strSQL, "DESCRIPTION", "SCENARIO", "scenario");

            frmPrevExp.uc_previous_expressions1.ShowDeleteButton = false;
            frmPrevExp.uc_previous_expressions1.ShowRecallButton = false;
            frmMain.g_oFrmMain.DeactivateStandByAnimation();
            result = frmPrevExp.ShowDialog();
            if (result == DialogResult.OK)
            {
                if (strRx.Trim().Length == 0)
                {
                    result = MessageBox.Show("All harvest cost component $/A/C values in scenario " + ScenarioId + " will be replaced \r\n" +
                                             "with harvest cost component $/A/C values from scenario " + frmPrevExp.uc_previous_expressions1.listView1.SelectedItems[0].SubItems[1].Text + "\r\n" +
                                             "Do you wish to continue with this action?(Y/N)", "FIA Biosum", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                }
                else
                {
                    result = MessageBox.Show("All harvest cost component $/A/C values in scenario " + ScenarioId + " with treatment " + strRx + " will be replaced \r\n" +
                         "with harvest cost component $/A/C values for treatment " + strRx + " from scenario " + frmPrevExp.uc_previous_expressions1.listView1.SelectedItems[0].SubItems[1].Text + "\r\n" +
                         "Do you wish to continue with this action?(Y/N)", "FIA Biosum", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                }
                if (result == DialogResult.Yes)
                {
                       
                    if (strRx.Trim().Length > 0)
                    {
                            //Clear out previous values before copying
                            dataMgr.m_strSQL = "UPDATE additional_harvest_costs_work_table " +
                                          "SET " + strColumn + " = NULL " +
                                          "WHERE rx='" + strRx + "' ";
                            dataMgr.SqlNonQuery(oConn, dataMgr.m_strSQL);

                            //dataMgr.m_strSQL = "UPDATE additional_harvest_costs_work_table a " +
                            //                  "INNER JOIN scenario_additional_harvest_costs b " +
                            //                  "ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " +
                            //                  "SET a." + strColumn + "=IIF(b." + strColumn + " IS NOT NULL,b." + strColumn + ",a." + strColumn + ") " +
                            //                  "WHERE a.rx='" + strRx + "' " +
                            //                  "AND b.scenario_id = '" + frmPrevExp.uc_previous_expressions1.listView1.SelectedItems[0].SubItems[1].Text.Trim() + "' ";
                            dataMgr.m_strSQL = $@"UPDATE additional_harvest_costs_work_table 
                                SET ({strColumn}) = (select CASE WHEN b.{strColumn} IS NOT NULL THEN b.{strColumn} ELSE additional_harvest_costs_work_table.{strColumn} END 
                                FROM scenario_additional_harvest_costs b WHERE additional_harvest_costs_work_table.biosum_cond_id=b.biosum_cond_id 
                                AND additional_harvest_costs_work_table.rx=b.rx and b.scenario_id = '{frmPrevExp.uc_previous_expressions1.listView1.SelectedItems[0].SubItems[1].Text.Trim()}')
                                WHERE rx = '{strRx}'";
                    }
                    else
                    {
                            //Clear out previous values before copying
                            dataMgr.m_strSQL = "UPDATE additional_harvest_costs_work_table " +
                                          "SET " + strColumn + " = NULL ";
                            dataMgr.SqlNonQuery(oConn, dataMgr.m_strSQL);

                            dataMgr.m_strSQL = $@"UPDATE additional_harvest_costs_work_table 
                                SET ({strColumn}) = (select CASE WHEN b.{strColumn} IS NOT NULL THEN b.{strColumn} ELSE additional_harvest_costs_work_table.{strColumn} END 
                                FROM scenario_additional_harvest_costs b WHERE additional_harvest_costs_work_table.biosum_cond_id=b.biosum_cond_id 
                                AND additional_harvest_costs_work_table.rx=b.rx and b.scenario_id = '{frmPrevExp.uc_previous_expressions1.listView1.SelectedItems[0].SubItems[1].Text.Trim()}')";
                    }

                       frmMain.g_oFrmMain.ActivateStandByAnimation(
                       this.ParentForm.WindowState,
                       this.ParentForm.Left,
                       this.ParentForm.Height,
                       this.ParentForm.Width,
                       this.ParentForm.Top);

                        frmMain.g_sbpInfo.Text = "Updating Harvest Cost Component $/A/C Values...Stand By";

                        dataMgr.SqlNonQuery(oConn, dataMgr.m_strSQL);

                    }   // End SQLite using SQLite conn

                    frmMain.g_sbpInfo.Text = "Ready";

                        if (dataMgr.m_intError == 0)
                        {
                            UpdateNullCounts(TempDb);
                            frmMain.g_oFrmMain.DeactivateStandByAnimation();
                            MessageBox.Show("Done");
                        }
                        else frmMain.g_oFrmMain.DeactivateStandByAnimation();                
                }

            }
            frmPrevExp.Close();
            frmPrevExp = null;

        }
        private void UpdateAllFromPrevious()
        {
            string strColumn = "";
           
            DialogResult result;



            frmDialog frmPrevExp = new frmDialog();

            frmPrevExp.Width = frmPrevExp.uc_previous_expressions1.m_intFullWd;
            frmPrevExp.Height = frmPrevExp.uc_previous_expressions1.m_intFullHt;
            frmPrevExp.Text = "Processor: Harvest Cost";

            frmPrevExp.uc_previous_expressions1.Visible = true;

            string strScenarioDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(strScenarioDB);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                dataMgr.m_strSQL = "SELECT DISTINCT a.scenario_id, a.Description, b.Record_Count " +
                    "FROM scenario a, (SELECT COUNT(*) AS Record_Count , scenario_id " +
                 "FROM scenario_additional_harvest_costs GROUP BY scenario_id)  b " +
               "WHERE a.scenario_id=b.scenario_id AND a.scenario_id <> '" + ScenarioId + "'";

                frmPrevExp.uc_previous_expressions1.lblTitle.Text = "Previous Scenario Harvest Cost Component Values";
                frmPrevExp.uc_previous_expressions1.loadvalues(dataMgr, oConn, dataMgr.m_strSQL, "DESCRIPTION", "SCENARIO", "scenario");
            }

            frmPrevExp.MinimizeBox = false;
            frmPrevExp.uc_previous_expressions1.ShowDeleteButton = false;
            frmPrevExp.uc_previous_expressions1.ShowRecallButton = false;
            result = frmPrevExp.ShowDialog();
            if (result == DialogResult.OK)
            {
                result = MessageBox.Show("All harvest cost component $/A/C values in scenario " + ScenarioId + " will be replaced \r\n" +
                                         "with harvest cost component $/A/C values from scenario " + frmPrevExp.uc_previous_expressions1.listView1.SelectedItems[0].SubItems[1].Text + "\r\n" +
                                         "Do you wish to continue with this action?(Y/N)", "FIA Biosum", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    m_oAdo.m_strSQL = "";

                    // Query the work table for the column names so we can clear their values before copying
                    System.Data.DataTable oTargetTableSchema = m_oAdo.getTableSchema(m_oAdo.m_OleDbConnection, "SELECT * FROM additional_harvest_costs_work_table");
                    string strTargetColumnsList = m_oAdo.getFieldNames(m_oAdo.m_OleDbConnection, "SELECT * FROM additional_harvest_costs_work_table");
                    string strTargetColumnsReservedWordFormattedList = m_oAdo.FormatReservedWordsInColumnNameList(strTargetColumnsList, ",");
                    string[] strTargetColumnsArray = frmMain.g_oUtils.ConvertListToArray(strTargetColumnsList, ",");
                    String strClearSQL = "UPDATE additional_harvest_costs_work_table SET ";
                    foreach (String strColName in strTargetColumnsArray)
                    {
                        // Add column to ClearSQL so we clear out the value before updating it from source
                        if (!strColName.ToUpper().Equals("SCENARIO_ID") &&
                            !strColName.ToUpper().Equals("BIOSUM_COND_ID") &&
                            !strColName.ToUpper().Equals("RX"))
                        strClearSQL = strClearSQL + strColName + " = NULL, ";
                    }
                    if (strClearSQL.Trim().Length > 0)
                    {
                        strClearSQL = strClearSQL.Substring(0, strClearSQL.Length - 2);
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, strClearSQL);
                    }

                    for (int x = 0; x <= uc_collection.Count - 1; x++)
                    {
                        strColumn = uc_collection.Item(x).ColumnName.Trim();

                        //make sure the source scenario has this column
                        if (m_oAdo.ColumnExist(m_oAdo.m_OleDbConnection, "scenario_additional_harvest_costs", strColumn))
                        {
                            
                            //make sure columnname not already referenced
                            if (m_oAdo.m_strSQL.ToUpper().IndexOf("B." + strColumn.ToUpper() + " IS NOT NULL", 0) < 0)
                            {
                                m_oAdo.m_strSQL = m_oAdo.m_strSQL + "a." + strColumn + "= IIF(b." + strColumn + " IS NOT NULL,b." + strColumn + ",a." + strColumn + "),";
                            }

                        }
                    }

                    if (m_oAdo.m_strSQL.Trim().Length > 0)
                    {
                        frmMain.g_oFrmMain.ActivateStandByAnimation(
                               this.ParentForm.WindowState,
                               this.ParentForm.Left,
                               this.ParentForm.Height,
                               this.ParentForm.Width,
                               this.ParentForm.Top);
                        //remove the comma at the end of the strings
                        m_oAdo.m_strSQL = m_oAdo.m_strSQL.Substring(0, m_oAdo.m_strSQL.Length - 1);

                        m_oAdo.m_strSQL = "UPDATE additional_harvest_costs_work_table a " +
                                          "INNER JOIN  scenario_additional_harvest_costs b " +
                                          "ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx " +
                                          "SET " + m_oAdo.m_strSQL +
                                          " WHERE b.scenario_id = '" + frmPrevExp.uc_previous_expressions1.listView1.SelectedItems[0].SubItems[1].Text.Trim() + "'";


                        frmMain.g_sbpInfo.Text = "Updating Harvest Cost Component $/A/C Values...Stand By";

                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

                        frmMain.g_sbpInfo.Text = "Ready";

                        if (m_oAdo.m_intError == 0)
                        {
                            UpdateNullCounts(TempDb);
                            frmMain.g_oFrmMain.DeactivateStandByAnimation();
                            MessageBox.Show("Done");
                        }
                        else
                            frmMain.g_oFrmMain.DeactivateStandByAnimation();

                    }
                }

            }
            frmPrevExp.Close();
            frmPrevExp = null;
        }
        private void btnEditPrev_Click(object sender, EventArgs e)
        {

            UpdateAllFromPrevious();
        }

    }
}

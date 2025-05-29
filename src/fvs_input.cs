using System;
using System.IO;
using System.Windows.Forms;
using System.Collections;
using System.ComponentModel;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Drawing.Text;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;


namespace FIA_Biosum_Manager
{
    /// <summary>
    /// Summary description for fvs_input.
    /// </summary>
    public class fvs_input
    {
        public ado_data_access m_ado;
        public dao_data_access m_dao;
        public int m_intError;
        private string m_strPlotTable;
        private string m_strCondTable;
        private string m_strTreeTable;
        private string m_strTreeSpcTable;
        private string m_strSiteTreeTable;
        private string m_strFVSTreeSpcTable;

        private Datasource m_DataSource;
        private string m_strProjDir;
        private string m_strFVSInMDBFile;
        public string m_strTempMDBFile;
        private string m_strVariant;
        private frmTherm m_frmTherm;
        private string m_strInDir;
        private string m_strConn;
        private System.Data.DataTable m_dt;
        private string m_strFVSTreeIdIDBWorkTable;
        private string m_strFVSTreeIdFIAWorkTable;
        private string m_strTempMDBFileNewProcess;
        private string m_strTempDBFile;

        // Constants for site index equation table
        const String SI_DELIM = "|";
        const String SI_EMPTY = "@";

        //Down Woody Materials fields
        private int m_intDWMOption;
        public enum m_enumDWMOption
        {
            SKIP_FUEL_MODEL_AND_DWM_DATA,
            USE_FUEL_MODEL_OR_DWM_DATA,
            USE_FUEL_MODEL_ONLY,
            USE_DWM_DATA_ONLY,
        };
        private string m_strDwmCoarseWoodyDebrisTable;
        private string m_strDwmFineWoodyDebrisTable;
        private string m_strDwmTransectSegmentTable;
        private string m_strDwmDuffLitterTable;
        private string m_strRefForestTypeTable;
        private string m_strRefForestTypeGroupTable;
        private string m_strFiaTreeSpeciesRefTable;
        private string m_strMinSmallFwdTransectLengthTotal = "10";
        private string m_strMinLargeFwdTransectLengthTotal = "30";
        private string m_strMinCwdTransectLengthTotal = "48";
        private string m_strMinDuffLitterPitCount;
        private string m_strDuffExcludedYears;
        private string m_strLitterExcludedYears;
        private bool m_bUsePrevHt;
        private bool m_bUsePrevDia;

        //Added for FIA2FVS implementation
        private bool m_bIncludeSeedlings;
        private string m_strSourceFiaDb;
        private string m_strDataDirectory;
        private string m_strGroup;

        private string m_strDebugFile;

        public fvs_input(string p_strProjDir, frmTherm p_frmTherm)
        {
            m_strProjDir = p_strProjDir.Trim();
            InitializeDataSource();

            this.m_strPlotTable = this.m_DataSource.getValidDataSourceTableName("PLOT");
            this.m_strCondTable = this.m_DataSource.getValidDataSourceTableName("CONDITION");
            this.m_strTreeTable = this.m_DataSource.getValidDataSourceTableName("TREE");
            this.m_strTreeSpcTable = this.m_DataSource.getValidDataSourceTableName("TREE SPECIES");
            this.m_strSiteTreeTable = this.m_DataSource.getValidDataSourceTableName("SITE TREE");
            this.m_strFVSTreeSpcTable = this.m_DataSource.getValidDataSourceTableName("FVS TREE SPECIES");

            m_strDwmCoarseWoodyDebrisTable = frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName;
            m_strDwmFineWoodyDebrisTable = frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName;
            m_strDwmTransectSegmentTable = frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName;
            m_strDwmDuffLitterTable = frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName;
            m_strRefForestTypeTable = "REF_FOREST_TYPE";
            m_strRefForestTypeGroupTable = "REF_FOREST_TYPE_GROUP";
            m_strFiaTreeSpeciesRefTable = m_DataSource.getValidDataSourceTableName("FIA Tree Species Reference");

            if (this.m_strPlotTable.Trim().Length == 0)
            {
                MessageBox.Show("!!Could Not Locate Plot Table!!", "FVS Input",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                return;
            }

            if (this.m_strCondTable.Trim().Length == 0)
            {
                MessageBox.Show("!!Could Not Locate Condition Table!!", "FVS Input",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                return;
            }

            if (this.m_strTreeTable.Trim().Length == 0)
            {
                MessageBox.Show("!!Could Not Locate Tree Table!!", "FVS Input",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                return;
            }

            if (this.m_strTreeSpcTable.Trim().Length == 0)
            {
                MessageBox.Show("!!Could Not Locate Tree Species Table!!", "FVS Input",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                return;
            }

            if (this.m_strSiteTreeTable.Trim().Length == 0)
            {
                MessageBox.Show("!!Could Not Locate Site Tree Table!!", "FVS Input",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                return;
            }

            if (this.m_strFVSTreeSpcTable.Trim().Length == 0)
            {
                MessageBox.Show("!!Could Not Locate FVS Tree Species Table!!", "FVS Input",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                return;
            }

            //this.m_strTempMDBFile = this.m_DataSource.CreateMDBAndTableDataSourceLinks();
            this.m_dao = new dao_data_access();
            m_ado = new ado_data_access();
            //this.m_strConn = m_ado.getMDBConnString(this.m_strTempMDBFile, "", "");
            this.m_frmTherm = p_frmTherm;
        }

        private void InitializeDataSource()
        {
            m_DataSource = new Datasource();
            m_DataSource.LoadTableColumnNamesAndDataTypes = false;
            m_DataSource.LoadTableRecordCount = false;
            m_DataSource.m_strDataSourceMDBFile = m_strProjDir + "\\db\\project.mdb";
            m_DataSource.m_strDataSourceTableName = "datasource";
            m_DataSource.m_strScenarioId = "";
            m_DataSource.populate_datasource_array();
        }

        ~fvs_input()
        {
            try
            {
                if (m_ado.m_OleDbDataAdapter != null)
                {
                    m_ado.m_OleDbDataAdapter.Dispose();
                    m_ado.m_OleDbDataAdapter = null;
                }

                if (m_ado.m_OleDbConnection != null)
                {
                    m_ado.m_OleDbConnection.Close();
                    m_ado.m_OleDbConnection.Dispose();
                    m_ado.m_OleDbConnection = null;
                }

                m_ado = null;
            }
            catch
            {
            }
            m_ado = null;
            if (m_dao != null)
            {
                if (this.m_dao.m_DaoWorkspace != null)
                {
                    this.m_dao.m_DaoWorkspace.Close();
                    this.m_dao.m_DaoWorkspace = null;
                }
                this.m_dao = null;
            }
            this.m_DataSource = null;
        }


        public void Start(string p_strVariant, string p_strDebugFile)
        {
            try
            {
                this.m_strDebugFile = p_strDebugFile;
                DebugLogMessage("*****START*****" + System.DateTime.Now.ToString() + "\r\n");
                DebugLogMessage("//Start(" + m_strDataDirectory + "," + p_strVariant + ")\r\n", 1);

                this.m_intError = 0;
                this.m_strInDir = m_strDataDirectory + "\\" + p_strVariant.Trim();
                this.m_strVariant = p_strVariant.Trim();
                this.strFVSInDBFile = "FVSIn.accdb";

                CheckDir();
                DeleteFiles();

                CopyFVSBlankDatabaseToFVSInDir(this.m_strInDir);

                CreateTableLinks();

                //create work tables with similar schemas to the FVS Input tables
                CreateFVSWorkTables();

                //Create/append to database input files
                if (this.m_intError != 0) return;
                //CreateFVSInputDbLOC();
                //if (this.m_intError != 0) return;
                CreateFVSStandInit();
                if (this.m_intError != 0) return;
                CreateFVSTreeInit();
                if (this.m_intError != 0) return;


                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,
                        "*****END*****" + System.DateTime.Now.ToString() + "\r\n");
            }
            finally
            {
                if (m_ado.m_DataSet != null)
                {
                    m_ado.m_DataSet.Clear();
                    m_ado.m_DataSet.Dispose();
                    m_ado.m_DataSet = null;
                }

                if (m_ado.m_OleDbCommand != null)
                {
                    m_ado.m_OleDbCommand.Dispose();
                    m_ado.m_OleDbCommand = null;
                }

                if (m_ado.m_OleDbConnection != null)
                {
                    m_ado.m_OleDbConnection.Dispose();
                    m_ado.m_OleDbConnection = null;
                }

                if (m_ado.m_OleDbDataAdapter != null)
                {
                    m_ado.m_OleDbDataAdapter.Dispose();
                    m_ado.m_OleDbDataAdapter = null;
                }

                if (m_ado.m_OleDbDataReader != null)
                {
                    m_ado.m_OleDbDataReader.Dispose();
                    m_ado.m_OleDbDataReader = null;
                }
            }
        }


        private void CreateFVSWorkTables()
        {
            /* Sometimes when processing multiple variants during FVSIn workflow, 
             * processing will stop before these tables are deleted 
             * if there are no stands in the standlist dataset.
             */
            DebugLogMessage("//CreateFVSWorkTables()\r\n", 1);

            using (var conn = new OleDbConnection(m_strConn))
            {
                conn.Open();
                if (m_ado.TableExist(conn, "FVS_StandInit_WorkTable"))
                {
                    DebugLogSQL(Queries.FVS.FVSInput.StandInit.DeleteFvsStandInitWorkTable());
                    m_ado.SqlNonQuery(conn, Queries.FVS.FVSInput.StandInit.DeleteFvsStandInitWorkTable());
                }

                if (m_ado.TableExist(conn, "FVS_TreeInit_WorkTable"))
                {
                    DebugLogSQL(Queries.FVS.FVSInput.TreeInit.DeleteWorkTable());
                    m_ado.SqlNonQuery(conn, Queries.FVS.FVSInput.TreeInit.DeleteWorkTable());
                }

                DebugLogSQL(frmMain.g_oTables.m_oFvs.CreateFVSInputStandInitTableSQL("FVS_StandInit_WorkTable"));
                frmMain.g_oTables.m_oFvs.CreateFVSInputStandInitTable(m_ado, conn,
                    "FVS_StandInit_WorkTable");
                DebugLogSQL(frmMain.g_oTables.m_oFvs.CreateFVSInputTreeInitTableSQL("FVS_TreeInit_WorkTable"));
                frmMain.g_oTables.m_oFvs.CreateFVSInputTreeInitWorkTable(m_ado, conn,
                    "FVS_TreeInit_WorkTable");
            }
        }

        public void CopyFVSBlankDatabaseToFVSInDir(string strFVSInDir)
        {
            env p_env = new env();
            string strFVSInSourcePath = p_env.strAppDir + "\\db\\" + this.strFVSInDBFile;
            string strFVSInDestPath = strFVSInDir + "\\" + this.strFVSInDBFile;
            File.Copy(strFVSInSourcePath, strFVSInDestPath, true);
            string strFVSInConn = m_ado.getMDBConnString(strFVSInDestPath, "", "");

            using (var conn = new OleDbConnection(strFVSInConn))
            {
                conn.Open();
                string strSQL = Queries.FVS.FVSInput.GroupAddFilesAndKeywords.UpdateAllPlots(this.strFVSInDBFile);
                m_ado.SqlNonQuery(conn, strSQL);
                strSQL = Queries.FVS.FVSInput.GroupAddFilesAndKeywords.UpdateAllStands(this.strFVSInDBFile);
                m_ado.SqlNonQuery(conn, strSQL);
            }

            p_env = null;
        }

        public void CreateTableLinks()
        {
            DebugLogMessage("//CreateTableLinks()\r\n", 1);

            if (m_dao != null)
            {
                if (m_dao.m_DaoDatabase != null)
                {
                    m_dao.m_DaoDatabase.Close();
                    System.Threading.Thread.Sleep(5000);
                    m_dao.m_DaoDatabase = null;
                }

                if (m_dao.m_DaoWorkspace != null)
                {
                    m_dao.m_DaoWorkspace.Close();
                    System.Threading.Thread.Sleep(5000);
                    m_dao.m_DaoWorkspace = null;
                }

                m_dao = null;
            }

            if (m_dao == null)
            {
                m_dao = new dao_data_access();
            }

            //Destination FVSIn.accdb tables
            m_dao.CreateTableLink(this.m_strTempMDBFile, "FVS_StandInit", this.m_strInDir + "\\" + this.strFVSInDBFile,
                "FVS_StandInit", true);
            m_dao.CreateTableLink(this.m_strTempMDBFile, "FVS_TreeInit", this.m_strInDir + "\\" + this.strFVSInDBFile,
                "FVS_TreeInit", true);


            //Reference Tables in biosum_ref.accdb
            m_dao.CreateTableLink(this.m_strTempMDBFile, m_strRefForestTypeTable,
                frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceDbFile,
                m_strRefForestTypeTable, true);
            m_dao.CreateTableLink(this.m_strTempMDBFile, m_strRefForestTypeGroupTable,
                frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceDbFile,
                m_strRefForestTypeGroupTable, true);

            //DWM Source Tables
            m_dao.CreateTableLink(this.m_strTempMDBFile, m_strDwmCoarseWoodyDebrisTable, m_strProjDir + "\\db\\master_aux.accdb", m_strDwmCoarseWoodyDebrisTable, true);
            m_dao.CreateTableLink(this.m_strTempMDBFile, m_strDwmFineWoodyDebrisTable, m_strProjDir + "\\db\\master_aux.accdb", m_strDwmFineWoodyDebrisTable, true);
            m_dao.CreateTableLink(this.m_strTempMDBFile, m_strDwmDuffLitterTable, m_strProjDir + "\\db\\master_aux.accdb", m_strDwmDuffLitterTable, true);
            m_dao.CreateTableLink(this.m_strTempMDBFile, m_strDwmTransectSegmentTable, m_strProjDir + "\\db\\master_aux.accdb", m_strDwmTransectSegmentTable, true);

            if (m_dao.m_DaoDatabase != null)
            {
                m_dao.m_DaoDatabase.Close();
                System.Threading.Thread.Sleep(5000);
                m_dao.m_DaoDatabase = null;
            }

            if (m_dao.m_DaoWorkspace != null)
            {
                m_dao.m_DaoWorkspace.Close();
                System.Threading.Thread.Sleep(5000);
                m_dao.m_DaoWorkspace = null;
            }

        }
        public void CreateTablesLinksToFVSIn()
        {
            DebugLogMessage("//CreateTablesLinksToFVSIn()\r\n", 1);

            if (m_dao != null)
            {
                if (m_dao.m_DaoDatabase != null)
                {
                    m_dao.m_DaoDatabase.Close();
                    System.Threading.Thread.Sleep(5000);
                    m_dao.m_DaoDatabase = null;
                }

                if (m_dao.m_DaoWorkspace != null)
                {
                    m_dao.m_DaoWorkspace.Close();
                    System.Threading.Thread.Sleep(5000);
                    m_dao.m_DaoWorkspace = null;
                }

                m_dao = null;
            }

            if (m_dao == null)
            {
                m_dao = new dao_data_access();
            }

            m_dao.CreateTableLink(this.m_strTempMDBFile, "FVS_StandInit", this.m_strInDir + "\\" + this.strFVSInDBFile,
                "FVS_StandInit", true);
            m_dao.CreateTableLink(this.m_strTempMDBFile, "FVS_TreeInit", this.m_strInDir + "\\" + this.strFVSInDBFile,
                "FVS_TreeInit", true);
            if (m_dao.m_DaoDatabase != null)
            {
                m_dao.m_DaoDatabase.Close();
                System.Threading.Thread.Sleep(5000);
                m_dao.m_DaoDatabase = null;
            }

            if (m_dao.m_DaoWorkspace != null)
            {
                m_dao.m_DaoWorkspace.Close();
                System.Threading.Thread.Sleep(5000);
                m_dao.m_DaoWorkspace = null;
            }
        }

        private void CreateFVSInputDbLOC()
        {
            DebugLogMessage("//CreateFVSInputDbLOC()\r\n", 1);
            try
            {
                System.IO.FileStream p_fs = new System.IO.FileStream(
                    this.m_strInDir + "\\" + this.m_strVariant + ".loc", System.IO.FileMode.Create,
                    System.IO.FileAccess.Write);

                System.IO.StreamWriter p_sw = new System.IO.StreamWriter(p_fs);
                p_sw.WriteLine("{0} {1} {2}", "C",
                    '"' + "FVS Input Database for " + this.m_strVariant + " variant" + '"',
                    this.strFVSInDBFile);
                p_sw.Close();
                p_fs.Close();
                p_sw = null;
                p_fs = null;
            }
            catch (Exception e)
            {
                MessageBox.Show("!!Error!! \n" +
                                "Module - fvs_input:CreateFVSInputDbLOC  \n" +
                                "Err Msg - " + e.Message.ToString().Trim(),
                    "FVS Input", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
        }

        private void CreateFVSStandInit()
        {
            DebugLogMessage("//CreateFVSStandInit()\r\n", 1);
            try
            {
                string strStandInitWorkTable = "FVS_StandInit_WorkTable";
                string strStandInit = "FVS_StandInit";

                //Import Plot/Cond data that doesn't require much extra transformation
                if (m_intError == 0)
                {
                    using (var conn = new OleDbConnection(m_strConn))
                    {
                        conn.Open();
                        m_ado.m_strSQL = Queries.FVS.FVSInput.StandInit.BulkImportStandDataFromBioSumMaster(
                            m_strVariant,
                            strStandInitWorkTable, m_strCondTable, m_strPlotTable);
                        m_ado.SqlNonQuery(conn, m_ado.m_strSQL);
                        DebugLogSQL(m_ado.m_strSQL);

                        //if (bUseGrmCalibrationData)
                        //{
                        //    m_ado.m_strSQL = String.Format(
                        //        "UPDATE {0} fvs INNER JOIN {1} grm ON fvs.Stand_ID=grm.biosum_cond_id " +
                        //        "SET fvs.DG_Measure=grm.measurement_period, " +
                        //        "fvs.HTG_Measure=grm.measurement_period, " +
                        //        "fvs.Mort_Measure=grm.measurement_period;",
                        //        strStandInitWorkTable, m_strGrmStandTable);
                        //    m_ado.SqlNonQuery(conn, m_ado.m_strSQL);
                        //    DebugLogSQL(m_ado.m_strSQL);
                        //}
                    }
                }

                //Site index and site species sequential insertion into FVS_StandInit_WorkTable
                if (m_intError == 0)
                {
                    GenerateSiteIndexAndSiteSpeciesSQL_access();
                }

                //Create lookup table for DWM Fuelbed Type Code to convert to FVS format
                if (m_intError == 0 && new int[]
                {
                    (int) m_enumDWMOption.USE_FUEL_MODEL_OR_DWM_DATA,
                    (int) m_enumDWMOption.USE_FUEL_MODEL_ONLY
                }.Contains(m_intDWMOption))
                {
                    using (var conn = new OleDbConnection(m_strConn))
                    {
                        conn.Open();
                        CreateDwmFuelbedTypeCodeFvsConversionTable(conn);
                        m_ado.m_strSQL =
                            Queries.FVS.FVSInput.StandInit.UpdateFuelModel(strStandInitWorkTable, m_strCondTable);
                        m_ado.SqlNonQuery(conn, m_ado.m_strSQL);
                        m_ado.SqlNonQuery(conn, "DROP TABLE Ref_DWM_Fuelbed_Type_Codes;");
                    }
                }

                //Calculate DWM biomass (tons/acre)
                if (m_intError == 0 && new int[]
                {
                    (int) m_enumDWMOption.USE_FUEL_MODEL_OR_DWM_DATA,
                    (int) m_enumDWMOption.USE_DWM_DATA_ONLY
                }.Contains(m_intDWMOption))
                {
                    PopulateFuelColumns_access();
                }


                //Write the final results to /projectroot/fvs/data/variant/FVSIn.accdb and clean up
                if (m_intError == 0)
                {
                    using (OleDbConnection conn = new OleDbConnection(m_strConn))
                    {
                        conn.Open();
                        m_ado.m_strSQL = Queries.FVS.FVSInput.StandInit.TranslateWorkTableToStandInitTable(
                            strStandInitWorkTable,
                            strStandInit);
                        m_ado.SqlNonQuery(conn, m_ado.m_strSQL);
                        DebugLogSQL(m_ado.m_strSQL);
                        //Delete work table
                        m_ado.m_strSQL = Queries.FVS.FVSInput.StandInit.DeleteFvsStandInitWorkTable();
                        m_ado.SqlNonQuery(conn, m_ado.m_strSQL);
                        DebugLogSQL(m_ado.m_strSQL);
                    }

                    UpdateFvsInConfigurationTable();
                    CreateConfigurationTextFile();
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(
                    "!!Error!! \n" + "Module - fvs_input:CreateFVSStandInit  \n" + "Err Msg - " +
                    e.Message.ToString().Trim(), "FVS Input", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            finally
            {
                if (m_ado.m_OleDbConnection != null)
                {
                    while (m_ado.m_OleDbConnection.State == System.Data.ConnectionState.Open)
                    {
                        m_ado.CloseConnection(m_ado.m_OleDbConnection);
                        System.Threading.Thread.Sleep(1000);
                    }
                    m_ado.m_OleDbConnection.Dispose();
                    m_ado.m_OleDbConnection = null;
                }
            }
        }

        private void CreateConfigurationTextFile()
        {
            string logFile = m_strProjDir + "/fvs/data/" + m_strVariant + "/biosum_fvs_input_configurations.txt";
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine("========================================================");
            stringBuilder.AppendLine("FVSIn.accdb created: " + DateTime.Now.ToString());
            stringBuilder.AppendLine("========================================================");
            string choice =
                (intDWMOption == (int)m_enumDWMOption.USE_FUEL_MODEL_ONLY ||
                 intDWMOption == (int)m_enumDWMOption.USE_FUEL_MODEL_OR_DWM_DATA)
                    ? "Yes"
                    : "No";
            stringBuilder.AppendLine("DWM Scott and Burgan (2005) Surface Fuel Model used: " + choice);
            choice = (intDWMOption == (int)m_enumDWMOption.USE_DWM_DATA_ONLY ||
                 intDWMOption == (int)m_enumDWMOption.USE_FUEL_MODEL_OR_DWM_DATA)
                    ? "Yes"
                    : "No";
            stringBuilder.AppendLine("DWM Fuel Biomass (tons/acre) calculated: " + choice);
            stringBuilder.AppendLine("Minimum Small/Medium FWD Transect Length (ft): " + m_strMinSmallFwdTransectLengthTotal);
            stringBuilder.AppendLine("Minimum Large FWD Transect Length (ft): " + m_strMinLargeFwdTransectLengthTotal);
            stringBuilder.AppendLine("Minimum CWD Transect Length (ft): " + m_strMinCwdTransectLengthTotal);
            stringBuilder.AppendLine("Duff years excluded: " + m_strDuffExcludedYears);
            stringBuilder.AppendLine("Litter years excluded: " + m_strLitterExcludedYears);
            stringBuilder.AppendLine("Previous height calibration data used: " + m_bUsePrevHt.ToString());
            stringBuilder.AppendLine("Previous diameter calibration data used: " + m_bUsePrevDia.ToString());
            stringBuilder.AppendLine("Include seedlings: " + m_bIncludeSeedlings.ToString());
            stringBuilder.AppendLine("Source FIA Data Mart database: " + m_strSourceFiaDb);
            stringBuilder.AppendLine("Selected FVS group: " + m_strGroup);
            frmMain.g_oUtils.WriteText(logFile, stringBuilder.ToString());
        }

        /// <summary>
        /// Connect to FVSIn for the current variant, add a table with two columns to record configurations
        /// </summary>
        private void UpdateFvsInConfigurationTable()
        {
            string strFvsInPathFile = m_strProjDir + "/fvs/data/" + m_strVariant + "/" + strFVSInDBFile;
            using (var conn = new OleDbConnection(m_ado.getMDBConnString(strFvsInPathFile, "", "")))
            {
                conn.Open();
                m_ado.m_strSQL = "CREATE TABLE biosum_fvsin_configuration (Setting TEXT(255), `Value` TEXT(255));";
                m_ado.SqlNonQuery(conn, m_ado.m_strSQL);
                DebugLogSQL(m_ado.m_strSQL);

                List<string[]> configs = CreateFVSInConfigurationList();
                foreach (string[] pair in configs)
                {
                    m_ado.m_strSQL = String.Format("INSERT INTO biosum_fvsin_configuration (Setting, `Value`) " +
                                                   "VALUES ('{0}','{1}');", pair[0], pair[1]);
                    m_ado.SqlNonQuery(conn, m_ado.m_strSQL);
                    DebugLogSQL(m_ado.m_strSQL);
                }
            }
        }

        /// <summary>
        /// Connect to FVSIn for the current variant, add a table with two columns to record configurations
        /// </summary>
        private void UpdateFvsInSqliteConfigurationTable()
        {
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            string strFvsInPathFile = m_strProjDir + "/fvs/data/" + Tables.FIA2FVS.DefaultFvsInputFile;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strFvsInPathFile)))
            {
                conn.Open();
                if (!oDataMgr.TableExist(conn, "biosum_fvsin_configuration"))
                {
                    oDataMgr.m_strSQL = "CREATE TABLE biosum_fvsin_configuration (Setting CHAR(255), `Value` CHAR(255));";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    DebugLogSQL(oDataMgr.m_strSQL);

                }

                List<string[]> configs = CreateFVSInConfigurationList();
                foreach (string[] pair in configs)
                {
                    oDataMgr.m_strSQL = String.Format("INSERT INTO biosum_fvsin_configuration (Setting, `Value`) " +
                                                   "VALUES ('{0}','{1}');", pair[0], pair[1]);
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    DebugLogSQL(oDataMgr.m_strSQL);
                }
            }
        }

        private List<string[]> CreateFVSInConfigurationList()
        {
            List<string[]> rows = new List<string[]>();
            rows.Add(new string[] { "Created date and time", DateTime.Now.ToString() });
            string choice =
                (intDWMOption == (int)m_enumDWMOption.USE_FUEL_MODEL_ONLY ||
                 intDWMOption == (int)m_enumDWMOption.USE_FUEL_MODEL_OR_DWM_DATA)
                    ? "Yes"
                    : "No";
            rows.Add(new string[] { "Included DWM Scott and Burgan (2005) Surface Fuel Model", choice });
            choice = (intDWMOption == (int)m_enumDWMOption.USE_DWM_DATA_ONLY ||
                      intDWMOption == (int)m_enumDWMOption.USE_FUEL_MODEL_OR_DWM_DATA)
                ? "Yes"
                : "No";
            rows.Add(new string[] { "Included DWM Fuel Biomass calculations (tons/acre)", choice });
            rows.Add(new string[]
                {"Minimum Small/Medium FWD Transect Length (ft)", m_strMinSmallFwdTransectLengthTotal});
            rows.Add(new string[] { "Minimum Large FWD Transect Length (ft)", m_strMinLargeFwdTransectLengthTotal });
            rows.Add(new string[] { "Minimum CWD Transect Length (ft)", m_strMinCwdTransectLengthTotal });
            rows.Add(new string[] { "Duff years excluded", m_strDuffExcludedYears });
            rows.Add(new string[] { "Litter years excluded", m_strLitterExcludedYears });
            rows.Add(new string[] { "Previous height calibration data used", m_bUsePrevHt.ToString() });
            rows.Add(new string[] { "Previous diameter calibration data used", m_bUsePrevDia.ToString() });
            rows.Add(new string[] { "Include seedlings", m_bIncludeSeedlings.ToString() });
            rows.Add(new string[] { "Source FIA Data Mart database", m_strSourceFiaDb });
            rows.Add(new string[] { "Selected FVS group", m_strGroup });
            return rows;
        }

        private void CreateDwmFuelbedTypeCodeFvsConversionTable(OleDbConnection conn)
        {
            if (!m_ado.TableExist(conn, "Ref_DWM_Fuelbed_Type_Codes"))
            {
                DebugLogMessage("Creating Fuelbed Type Code lookup table\r\n", 1);
                DebugLogSQL(Queries.FVS.FVSInput.StandInit.CreateDWMFuelbedTypCdToFVSConversionTable());
                m_ado.SqlNonQuery(conn, Queries.FVS.FVSInput.StandInit.CreateDWMFuelbedTypCdToFVSConversionTable());
                m_intError = m_ado.m_intError;

                Dictionary<string, int> fuelModelNumbers = new Dictionary<string, int>
                {
                    {"GR1", 101},
                    {"GR2", 102},
                    {"GR3", 103},
                    {"GR4", 104},
                    {"GR5", 105},
                    {"GR6", 106},
                    {"GR7", 107},
                    {"GR8", 108},
                    {"GR9", 109},
                    {"GS1", 121},
                    {"GS2", 122},
                    {"GS3", 123},
                    {"GS4", 124},
                    {"SH1", 141},
                    {"SH2", 142},
                    {"SH3", 143},
                    {"SH4", 144},
                    {"SH5", 145},
                    {"SH6", 146},
                    {"SH7", 147},
                    {"SH8", 148},
                    {"SH9", 149},
                    {"TU1", 161},
                    {"TU2", 162},
                    {"TU3", 163},
                    {"TU4", 164},
                    {"TU5", 165},
                    {"TL1", 181},
                    {"TL2", 182},
                    {"TL3", 183},
                    {"TL4", 184},
                    {"TL5", 185},
                    {"TL6", 186},
                    {"TL7", 187},
                    {"TL8", 188},
                    {"TL9", 189},
                    {"SB1", 201},
                    {"SB2", 202},
                    {"SB3", 203},
                    {"SB4", 204}
                };

                foreach (string key in fuelModelNumbers.Keys)
                {
                    string strSQL = String.Format(
                        "INSERT INTO Ref_DWM_Fuelbed_Type_Codes " +
                        "(dwm_fuelbed_typcd, fuel_model) VALUES (\'{0}\', {1}) ",
                        key, fuelModelNumbers[key]);
                    DebugLogSQL(strSQL);
                    m_ado.SqlNonQuery(conn, strSQL);
                    m_intError = m_ado.m_intError;
                }
            }
        }

        private void CreateDwmFuelbedTypeCodeFvsConversionTableNew(SQLite.ADO.DataMgr oDataMgr, System.Data.SQLite.SQLiteConnection conn)
        {
            if (!oDataMgr.TableExist(conn, "Ref_DWM_Fuelbed_Type_Codes"))
            {
                DebugLogMessage("Creating Fuelbed Type Code lookup table\r\n", 1);
                DebugLogSQL(Queries.FVS.FVSInput.StandInit.CreateDWMFuelbedTypCdToFVSConversionTable());
                oDataMgr.SqlNonQuery(conn, Queries.FVS.FVSInput.StandInit.CreateDWMFuelbedTypCdToFVSConversionTable());
                m_intError = oDataMgr.m_intError;

                Dictionary<string, int> fuelModelNumbers = new Dictionary<string, int>
                {
                    {"GR1", 101},
                    {"GR2", 102},
                    {"GR3", 103},
                    {"GR4", 104},
                    {"GR5", 105},
                    {"GR6", 106},
                    {"GR7", 107},
                    {"GR8", 108},
                    {"GR9", 109},
                    {"GS1", 121},
                    {"GS2", 122},
                    {"GS3", 123},
                    {"GS4", 124},
                    {"SH1", 141},
                    {"SH2", 142},
                    {"SH3", 143},
                    {"SH4", 144},
                    {"SH5", 145},
                    {"SH6", 146},
                    {"SH7", 147},
                    {"SH8", 148},
                    {"SH9", 149},
                    {"TU1", 161},
                    {"TU2", 162},
                    {"TU3", 163},
                    {"TU4", 164},
                    {"TU5", 165},
                    {"TL1", 181},
                    {"TL2", 182},
                    {"TL3", 183},
                    {"TL4", 184},
                    {"TL5", 185},
                    {"TL6", 186},
                    {"TL7", 187},
                    {"TL8", 188},
                    {"TL9", 189},
                    {"SB1", 201},
                    {"SB2", 202},
                    {"SB3", 203},
                    {"SB4", 204}
                };

                foreach (string key in fuelModelNumbers.Keys)
                {
                    oDataMgr.m_strSQL = String.Format(
                        "INSERT INTO Ref_DWM_Fuelbed_Type_Codes " +
                        "(dwm_fuelbed_typcd, fuel_model) VALUES (\'{0}\', {1}) ",
                        key, fuelModelNumbers[key]);
                    DebugLogSQL(oDataMgr.m_strSQL);
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    m_intError = oDataMgr.m_intError;
                }
            }
        }

        private void GenerateSiteIndexAndSiteSpeciesSQL_access()
        {
            using (var conn = new OleDbConnection(m_strConn))
            {
                m_ado.m_OleDbConnection = conn;
                conn.Open();
                fvs_input.site_index oSiteIndex = new site_index();
                oSiteIndex.ado_data_access = m_ado;
                oSiteIndex.CondTable = this.m_strCondTable;
                oSiteIndex.PlotTable = this.m_strPlotTable;
                oSiteIndex.TreeTable = this.m_strTreeTable;
                oSiteIndex.SiteTreeTable = this.m_strSiteTreeTable;
                oSiteIndex.TreeSpeciesTable = this.m_strTreeSpcTable;
                oSiteIndex.FVSTreeSpeciesTable = this.m_strFVSTreeSpcTable;
                oSiteIndex.SiteIndexEquations = LoadSiteIndexEquations_access(this.m_strVariant.Trim().ToUpper());
                oSiteIndex.DebugFile = this.m_strDebugFile;

                m_ado.m_strSQL =
                    Queries.FVS.FVSInput.StandInit.CreateSiteIndexDataset(m_strVariant, m_strCondTable, m_strPlotTable);

                m_ado.CreateDataSet(conn, m_ado.m_strSQL, "standlist");
                if (m_ado.m_DataSet.Tables["standlist"].Rows.Count == 0)
                {
                    this.m_intError = -1;
                    MessageBox.Show("!!No standlist Records To Process!!", "FVS Input",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Exclamation);
                    m_ado.m_DataSet.Clear();
                    m_ado.m_DataSet.Dispose();
                    return;
                }

                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                    "Writing BioSum FVS_StandInit For Variant " + this.m_strVariant);
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Maximum",
                    m_ado.m_DataSet.Tables["standlist"].Rows.Count);
                this.m_dt = m_ado.m_DataSet.Tables["standlist"];

                DebugLogMessage("Inserting Site Index/Site Species\r\n", 1);
                for (int x = 0; x <= this.m_dt.Rows.Count - 1; x++)
                {
                    string strStand_ID = "null";
                    string strSite_Species = "null";
                    string strSite_Index = "null";

                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", x);
                    strStand_ID = "\'" + this.m_dt.Rows[x]["biosum_cond_id"].ToString().Trim() + "\'";
                    oSiteIndex.getSiteIndex_access(m_dt.Rows[x]);
                    strSite_Species = "\'" + oSiteIndex.SiteIndexSpeciesAlphaCode + "\'";
                    strSite_Index = oSiteIndex.SiteIndex;

                    if (strSite_Species.Contains("@"))
                    {
                        strSite_Species = "null";
                    }

                    if (strSite_Index.Contains("@"))
                    {
                        strSite_Index = "null";
                    }

                    if (strSite_Species != "null" && strSite_Index != "null")
                    {
                        m_ado.m_strSQL =
                            Queries.FVS.FVSInput.StandInit.InsertSiteIndexSpeciesRow(strStand_ID, strSite_Species,
                                strSite_Index);
                        DebugLogSQL(m_ado.m_strSQL);
                        m_ado.SqlNonQuery(conn, m_ado.m_strSQL);
                    }
                }

                m_ado.m_DataSet.Clear();
                m_ado.m_DataSet.Dispose();
                m_ado.m_DataSet = null;
            }
        }

        private void GenerateSiteIndexAndSiteSpeciesSQL(string strVariant)
        {
            using (OleDbConnection conn = new OleDbConnection(m_ado.getMDBConnString(m_strTempMDBFileNewProcess, "", "")))
            {
                conn.Open();

                fvs_input.site_index oSiteIndex = new site_index();
                oSiteIndex.ado_data_access = m_ado;
                oSiteIndex.CondTable = this.m_strCondTable;
                oSiteIndex.PlotTable = this.m_strPlotTable;
                oSiteIndex.TreeTable = this.m_strTreeTable;
                oSiteIndex.SiteTreeTable = this.m_strSiteTreeTable;
                oSiteIndex.TreeSpeciesTable = this.m_strTreeSpcTable;
                oSiteIndex.FVSTreeSpeciesTable = this.m_strFVSTreeSpcTable;
                oSiteIndex.SiteIndexEquations = LoadSiteIndexEquations(strVariant.Trim().ToUpper());
                oSiteIndex.DebugFile = this.m_strDebugFile;
                oSiteIndex.TempMDBFileNewProcess = this.m_strTempMDBFileNewProcess;
                oSiteIndex.ProjectDirectory = this.m_strProjDir;

                m_ado.m_strSQL =
                    Queries.FVS.FVSInput.StandInit.CreateSiteIndexDataset(strVariant, m_strCondTable, m_strPlotTable);
                m_ado.CreateDataSet(conn, m_ado.m_strSQL, "standlist");
                if (m_ado.m_DataSet.Tables["standlist"].Rows.Count == 0)
                {
                    this.m_intError = -1;
                    MessageBox.Show("!!No standlist Records To Process!!", "FVS Input",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Exclamation);
                    m_ado.m_DataSet.Clear();
                    m_ado.m_DataSet.Dispose();
                    return;
                }

                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                    "Writing BioSum FVS_StandInit For Variant " + strVariant);
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Maximum",
                    m_ado.m_DataSet.Tables["standlist"].Rows.Count);
                this.m_dt = m_ado.m_DataSet.Tables["standlist"];

                DebugLogMessage("Inserting Site Index/Site Species\r\n", 1);

                for (int x = 0; x <= this.m_dt.Rows.Count - 1; x++)
                {
                    string strStand_ID = "null";
                    string strSite_Species = "null";
                    string strSite_Index = "null";
                    string strBase_Age = "null";

                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", x);
                    strStand_ID = "\'" + this.m_dt.Rows[x]["biosum_cond_id"].ToString().Trim() + "\'";
                    oSiteIndex.getSiteIndex(m_dt.Rows[x]);
                    strSite_Species = "\'" + oSiteIndex.SiteIndexSpeciesAlphaCode + "\'";
                    strSite_Index = oSiteIndex.SiteIndex;
                    strBase_Age = oSiteIndex.BaseAge;

                    if (strSite_Species.Contains("@"))
                    {
                        strSite_Species = "null";
                    }

                    if (strSite_Index.Contains("@"))
                    {
                        strSite_Index = "null";
                    }

                    if (strBase_Age.Contains("@") || strBase_Age == "")
                    {
                        strBase_Age = "null";
                    }

                    if (strSite_Species != "null" && strSite_Index != "null")
                    {
                        m_ado.m_strSQL =
                            Queries.FVS.FVSInput.StandInit.InsertSiteIndexSpeciesRowNew(strStand_ID, strSite_Species,
                                strSite_Index, strBase_Age);
                        DebugLogSQL(m_ado.m_strSQL);
                        m_ado.SqlNonQuery(conn, m_ado.m_strSQL);
                    }
                }
            }
        }

        private void GenerateSiteIndexAndSiteSpeciesSQLNew(SQLite.ADO.DataMgr oDataMgr, System.Data.SQLite.SQLiteConnection conn, string strVariant)
        {
            fvs_input.site_index oSiteIndex = new site_index();
            oSiteIndex.CondTable = "master." + this.m_strCondTable;
            oSiteIndex.PlotTable = "master." + this.m_strPlotTable;
            oSiteIndex.TreeTable = "master." + this.m_strTreeTable;
            oSiteIndex.SiteTreeTable = "master." + this.m_strSiteTreeTable;
            oSiteIndex.SiteIndexEquations = LoadSiteIndexEquations(strVariant.Trim().ToUpper());
            oSiteIndex.DebugFile = this.m_strDebugFile;
            oSiteIndex.dataMgr = oDataMgr;
            oSiteIndex.tempConn = conn;
            oSiteIndex.ProjectDirectory = this.m_strProjDir;

            oDataMgr.m_strSQL =
                   Queries.FVS.FVSInput.StandInit.CreateSiteIndexDatasetNew(strVariant, "master." + m_strCondTable, "master." + m_strPlotTable);
            oDataMgr.CreateDataSet(conn, oDataMgr.m_strSQL, "standlist");
            if (oDataMgr.m_DataSet.Tables["standlist"].Rows.Count == 0)
            {
                this.m_intError = -1;
                MessageBox.Show("!!No standlist Records To Process!!", "FVS Input",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                oDataMgr.m_DataSet.Clear();
                oDataMgr.m_DataSet.Dispose();
                return;
            }

            frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                    "Writing BioSum FVS_StandInit For Variant " + strVariant);
            frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Minimum", 0);
            frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Maximum",
                oDataMgr.m_DataSet.Tables["standlist"].Rows.Count);
            m_dt = oDataMgr.m_DataSet.Tables["standlist"];

            DebugLogMessage("Inserting Site Index/Site Species\r\n", 1);

            for (int x = 0; x <= m_dt.Rows.Count - 1; x++)
            {
                string strStand_ID;
                string strSite_Species;
                string strSite_Index;
                string strBase_Age;

                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", x);
                strStand_ID = "\'" + m_dt.Rows[x]["biosum_cond_id"].ToString().Trim() + "\'";
                oSiteIndex.getSiteIndexNew(m_dt.Rows[x]);
                strSite_Species = "\'" + oSiteIndex.SiteIndexSpecies + "\'";
                strSite_Index = oSiteIndex.SiteIndex;
                strBase_Age = oSiteIndex.BaseAge;

                if (strSite_Species.Contains("@"))
                {
                    strSite_Species = "null";
                }

                if (strSite_Index.Contains("@"))
                {
                    strSite_Index = "null";
                }

                if (strBase_Age.Contains("@") || strBase_Age == "")
                {
                    strBase_Age = "null";
                }

                if (strSite_Species != "null" && strSite_Index != "null")
                {
                    oDataMgr.m_strSQL =
                        Queries.FVS.FVSInput.StandInit.InsertSiteIndexSpeciesRowNew(strStand_ID, strSite_Species,
                            strSite_Index, strBase_Age);
                    DebugLogSQL(oDataMgr.m_strSQL);
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
            }
        }

        private void PopulateFuelColumns_access()
        {
            //COARSE WOODY DEBRIS
            DebugLogMessage("Executing Coarse Woody Debris SQL (multiple steps)\r\n", 1);
            foreach (string strSQL in Queries.FVS.FVSInput.StandInit.CalculateCoarseWoodyDebrisBiomassTonsPerAcre(
                m_strDwmCoarseWoodyDebrisTable, m_strCondTable, m_strFiaTreeSpeciesRefTable,
                m_strDwmTransectSegmentTable, m_strMinCwdTransectLengthTotal))
            {
                if (!String.IsNullOrEmpty(strSQL) && m_intError == 0)
                {
                    using (OleDbConnection conn = new OleDbConnection(m_strConn))
                    {
                        conn.Open();
                        DebugLogSQL(strSQL);
                        m_ado.SqlNonQuery(conn, strSQL);
                        m_intError = m_ado.m_intError;
                    }
                }
            }

            //FINE WOODY DEBRIS
            DebugLogMessage("Executing Fine Woody Debris SQL (multiple steps)\r\n", 1);
            if (m_intError == 0)
            {
                foreach (string strSQL in Queries.FVS.FVSInput.StandInit.CalculateFineWoodyDebrisBiomassTonsPerAcre(
                    m_strDwmFineWoodyDebrisTable, m_strCondTable, m_strMinSmallFwdTransectLengthTotal,
                    m_strMinLargeFwdTransectLengthTotal))
                {
                    using (OleDbConnection conn = new OleDbConnection(m_strConn))
                    {
                        conn.Open();

                        if (!String.IsNullOrEmpty(strSQL) && m_intError == 0)
                        {
                            DebugLogSQL(strSQL);
                            m_ado.SqlNonQuery(conn, strSQL);
                            m_intError = m_ado.m_intError;
                            if (m_intError != 0)
                                break;
                        }
                    }
                }
            }

            //DUFF LITTER
            DebugLogMessage("Executing Duff Litter SQL (multiple steps)\r\n", 1);
            if (m_intError == 0)
            {
                using (OleDbConnection conn = new OleDbConnection(m_strConn))
                {
                    conn.Open();
                    string strSQL =
                        Queries.FVS.FVSInput.StandInit.CalculateDuffLitterBiomassTonsPerAcre(m_strDwmDuffLitterTable,
                            m_strCondTable);
                    DebugLogSQL(m_ado.m_strSQL);
                    m_ado.SqlNonQuery(conn, strSQL);
                    m_intError = m_ado.m_intError;

                    strSQL = Queries.FVS.FVSInput.StandInit.UpdateFvsStandInitDuffLitterColumns();
                    DebugLogSQL(strSQL);
                    m_ado.SqlNonQuery(conn, strSQL);
                    m_intError = m_ado.m_intError;

                    if (!string.IsNullOrEmpty(m_strDuffExcludedYears))
                    {
                        strSQL = Queries.FVS.FVSInput.StandInit.RemoveDuffYears(m_strDuffExcludedYears);
                        DebugLogSQL(strSQL);
                        m_ado.SqlNonQuery(conn, strSQL);
                        m_intError = m_ado.m_intError;
                    }
                    if (!string.IsNullOrEmpty(m_strLitterExcludedYears))
                    {
                        strSQL = Queries.FVS.FVSInput.StandInit.RemoveLitterYears(m_strLitterExcludedYears);
                        DebugLogSQL(strSQL);
                        m_ado.SqlNonQuery(conn, strSQL);
                        m_intError = m_ado.m_intError;
                    }
                }
            }

            //Clean up worktables
            DebugLogMessage("Deleting work tables\r\n", 1);
            if (m_intError == 0)
            {
                foreach (string strSQL in Queries.FVS.FVSInput.StandInit.DeleteDwmWorkTables())
                {
                    using (OleDbConnection conn = new OleDbConnection(m_strConn))
                    {
                        conn.Open();
                        DebugLogSQL(strSQL);
                        m_ado.SqlNonQuery(conn, strSQL);
                        m_intError = m_ado.m_intError;
                        if (m_intError != 0)
                            break;
                    }
                }
            }
        }

        
        private void PopulateFuelColumns(string strMDBConn, string strWorkDb, string strStandTable)
        {
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            string strDWMTablesDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\master_aux.db";
            string strBioSumRefDb = frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                        frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceSqliteFile;
            string strWorkDbConn = oDataMgr.GetConnectionString(strWorkDb);

            FIA_Biosum_Manager.Datasource oProjectDs = new Datasource();
            string strTreeSpeciesDb = oProjectDs.CreateTemporarySQLiteSpeciesRefTable();

            //COARSE WOODY DEBRIS
            DebugLogMessage("Executing Coarse Woody Debris SQL (multiple steps)\r\n", 1);
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strWorkDbConn))
            {
                conn.Open();

                oDataMgr.m_strSQL = "ATTACH DATABASE '" + strDWMTablesDb + "' AS dwmdb";
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                oDataMgr.m_strSQL = "ATTACH DATABASE '" + strTreeSpeciesDb + "' AS tree_species";
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                string[] strCwdSqlStmts = Queries.FVS.FVSInput.StandInit.CalculateCoarseWoodyDebrisBiomassTonsPerAcreSqlite(
                m_strDwmCoarseWoodyDebrisTable, m_strFiaTreeSpeciesRefTable,
                m_strDwmTransectSegmentTable, m_strMinCwdTransectLengthTotal, strStandTable);
                foreach (string strSQL in strCwdSqlStmts)
                {
                    if (!String.IsNullOrEmpty(strSQL) && m_intError == 0)
                    {
                            DebugLogSQL(strSQL);
                            oDataMgr.SqlNonQuery(conn, strSQL);
                            m_intError = oDataMgr.m_intError;
                    }
                }

                //FINE WOODY DEBRIS
                DebugLogMessage("Executing Fine Woody Debris SQL (multiple steps)\r\n", 1);
                oDataMgr.m_strSQL = "ATTACH DATABASE '" + strBioSumRefDb + "' AS biosumref";
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                if (m_intError == 0)
                {
                    string[] strFwdSqlStmts = Queries.FVS.FVSInput.StandInit.CalculateFineWoodyDebrisBiomassTonsPerAcreSqlite(
                        m_strDwmFineWoodyDebrisTable, m_strCondTable, m_strMinSmallFwdTransectLengthTotal,
                        m_strMinLargeFwdTransectLengthTotal, strStandTable, null, null);
                    foreach (string strSQL in strFwdSqlStmts)
                    {
                        if (Array.IndexOf(strFwdSqlStmts, strSQL) == 1)
                        {
                            using (OleDbConnection accessConn = new OleDbConnection(strMDBConn))
                            {
                                accessConn.Open();

                                if (!m_ado.TableExist(accessConn, "DWM_FWD_Aggregates_WorkTable"))
                                {
                                    m_dao.CreateSQLiteTableLink(m_strTempMDBFileNewProcess, "DWM_FWD_Aggregates_WorkTable", "DWM_FWD_Aggregates_WorkTable",
                                ODBCMgr.DSN_KEYS.Fia2FvsOutputDsnName, strWorkDb);
                                    int i = 0;
                                    do
                                    {
                                        // break out of loop if it runs too long
                                        if (i > 20)
                                        {
                                            break;
                                        }
                                        System.Threading.Thread.Sleep(1000);
                                        i++;
                                    }
                                    while (!m_ado.TableExist(accessConn, "DWM_FWD_Aggregates_WorkTable"));
                                }

                                DebugLogSQL(strSQL);
                                m_ado.SqlNonQuery(accessConn, strSQL);
                                m_intError = m_ado.m_intError;
                                if (m_intError != 0)
                                    break;
                            }
                            
                        }
                        else
                        {
                            DebugLogSQL(strSQL);
                            oDataMgr.SqlNonQuery(conn, strSQL);
                            m_intError = oDataMgr.m_intError;
                        }
                    }
                }

                //DUFF LITTER
                DebugLogMessage("Executing Duff Litter SQL (multiple steps)\r\n", 1);
                if (m_intError == 0)
                {
                    oDataMgr.m_strSQL = Queries.FVS.FVSInput.StandInit.CreateDuffLitterBiomassTonsPerAcreSqlite();
                    DebugLogSQL(oDataMgr.m_strSQL);
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    m_intError = oDataMgr.m_intError;
                    using (OleDbConnection accessConn = new OleDbConnection(strMDBConn))
                    {
                        accessConn.Open();

                        if (!m_ado.TableExist(accessConn, "DWM_DuffLitter_Aggregates_WorkTable"))
                        {
                            m_dao.CreateSQLiteTableLink(m_strTempMDBFileNewProcess, "DWM_DuffLitter_Aggregates_WorkTable", "DWM_DuffLitter_Aggregates_WorkTable",
                                ODBCMgr.DSN_KEYS.Fia2FvsOutputDsnName, strWorkDb);
                            int i = 0;
                            do
                            {
                                // break out of loop if it runs too long
                                if (i > 20)
                                {
                                    break;
                                }
                                System.Threading.Thread.Sleep(1000);
                                i++;
                            }
                            while (!m_ado.TableExist(accessConn, "DWM_DuffLitter_Aggregates_WorkTable"));
                        }

                        string strSQL =
                            Queries.FVS.FVSInput.StandInit.CalculateDuffLitterBiomassTonsPerAcreSqlite(m_strDwmDuffLitterTable,
                                m_strCondTable, strStandTable, null, null);
                        DebugLogSQL(m_ado.m_strSQL);
                        m_ado.SqlNonQuery(accessConn, strSQL);
                        m_intError = m_ado.m_intError;
                    }
                    oDataMgr.m_strSQL = Queries.FVS.FVSInput.StandInit.UpdateFvsStandInitDuffLitterColumnsSqlite(strStandTable);
                    DebugLogSQL(oDataMgr.m_strSQL);
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    m_intError = oDataMgr.m_intError;

                    if (!string.IsNullOrEmpty(m_strDuffExcludedYears))
                    {
                        oDataMgr.m_strSQL = Queries.FVS.FVSInput.StandInit.RemoveDuffYearsSqlite(m_strDuffExcludedYears, strStandTable);
                        DebugLogSQL(oDataMgr.m_strSQL);
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                        m_intError = oDataMgr.m_intError;
                    }
                    if (!string.IsNullOrEmpty(m_strLitterExcludedYears))
                    {
                        oDataMgr.m_strSQL = Queries.FVS.FVSInput.StandInit.RemoveLitterYearsSqlite(m_strLitterExcludedYears, strStandTable);
                        DebugLogSQL(oDataMgr.m_strSQL);
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                        m_intError = oDataMgr.m_intError;
                    }
                }

                //Clean up worktables
                DebugLogMessage("Deleting work tables\r\n", 1);
                if (m_intError == 0)
                {
                    string[] strDeleteStmts = Queries.FVS.FVSInput.StandInit.DeleteDwmWorkTables();
                    foreach (string strSQL in strDeleteStmts)
                    {
                        if (Array.IndexOf(strDeleteStmts, strSQL) == 3 ||
                            Array.IndexOf(strDeleteStmts, strSQL) == 4)
                        {
                            using (OleDbConnection accessConn = new OleDbConnection(strMDBConn))
                            {
                                accessConn.Open();
                                DebugLogSQL(strSQL);
                                m_ado.SqlNonQuery(accessConn, strSQL);
                                m_intError = m_ado.m_intError;
                                if (m_intError != 0)
                                    break;
                            }
                        }
                        DebugLogSQL(strSQL);
                        oDataMgr.SqlNonQuery(conn, strSQL);
                        m_intError = oDataMgr.m_intError;
                        if (m_intError != 0)
                            break;
                    }
                }
            }
        }

        private void PopulateFuelColumnsNew(SQLite.ADO.DataMgr oDataMgr, System.Data.SQLite.SQLiteConnection conn, string strStandTable)
        {
            string strDWMTablesDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\master_aux.db";
            FIA_Biosum_Manager.Datasource oProjectDs = new Datasource();
            string strBioSumRefDb = frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                        frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceSqliteFile;
            string strForestTypeTable = "ref.REF_FOREST_TYPE";
            string strForestTypeGroupTable = "ref.REF_FOREST_TYPE_GROUP";

            if (!oDataMgr.DatabaseAttached(conn, strDWMTablesDb))
            {
                oDataMgr.m_strSQL = "ATTACH DATABASE '" + strDWMTablesDb + "' AS dwmdb";
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
            }
            if (!oDataMgr.DatabaseAttached(conn, strBioSumRefDb))
            {
                oDataMgr.m_strSQL = "ATTACH DATABASE '" + strBioSumRefDb + "' AS ref";
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
            }

            //COARSE WOODY DEBRIS
            string[] strCwdSqlStmts = Queries.FVS.FVSInput.StandInit.CalculateCoarseWoodyDebrisBiomassTonsPerAcreSqlite(
                "dwmdb." + m_strDwmCoarseWoodyDebrisTable, "ref." + m_strFiaTreeSpeciesRefTable,
                "dwmdb." + m_strDwmTransectSegmentTable, m_strMinCwdTransectLengthTotal, strStandTable);
            foreach (string strSQL in strCwdSqlStmts)
            {
                if (!String.IsNullOrEmpty(strSQL) && m_intError == 0)
                {
                    DebugLogSQL(strSQL);
                    oDataMgr.SqlNonQuery(conn, strSQL);
                    m_intError = oDataMgr.m_intError;
                }
            }

            //FINE WOODY DEBRIS
            DebugLogMessage("Executing Fine Woody Debris SQL (multiple steps)\r\n", 1);
            if (m_intError == 0)
            {
                string[] strFwdSqlStmts = Queries.FVS.FVSInput.StandInit.CalculateFineWoodyDebrisBiomassTonsPerAcreSqlite(
                    "dwmdb." + m_strDwmFineWoodyDebrisTable, "master." + m_strCondTable, m_strMinSmallFwdTransectLengthTotal,
                    m_strMinLargeFwdTransectLengthTotal, strStandTable, strForestTypeTable, strForestTypeGroupTable);
                foreach (string strSQL in strFwdSqlStmts)
                {
                    DebugLogSQL(strSQL);
                    oDataMgr.SqlNonQuery(conn, strSQL);
                    m_intError = oDataMgr.m_intError;
                }
            }

            //DUFF LITTER
            DebugLogMessage("Executing Duff Litter SQL (multiple steps)\r\n", 1);
            if (m_intError == 0)
            {
                oDataMgr.m_strSQL = Queries.FVS.FVSInput.StandInit.CreateDuffLitterBiomassTonsPerAcreSqlite();
                DebugLogSQL(oDataMgr.m_strSQL);
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                m_intError = oDataMgr.m_intError;

                oDataMgr.m_strSQL = Queries.FVS.FVSInput.StandInit.CalculateDuffLitterBiomassTonsPerAcreSqlite("dwmdb." + m_strDwmDuffLitterTable,
                                "master." + m_strCondTable, strStandTable, strForestTypeTable, strForestTypeGroupTable);
                DebugLogSQL(oDataMgr.m_strSQL);
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                m_intError = oDataMgr.m_intError;

                oDataMgr.m_strSQL = Queries.FVS.FVSInput.StandInit.UpdateFvsStandInitDuffLitterColumnsSqlite(strStandTable);
                DebugLogSQL(oDataMgr.m_strSQL);
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                m_intError = oDataMgr.m_intError;

                if (!string.IsNullOrEmpty(m_strDuffExcludedYears))
                {
                    oDataMgr.m_strSQL = Queries.FVS.FVSInput.StandInit.RemoveDuffYearsSqlite(m_strDuffExcludedYears, strStandTable);
                    DebugLogSQL(oDataMgr.m_strSQL);
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    m_intError = oDataMgr.m_intError;
                }
                if (!string.IsNullOrEmpty(m_strLitterExcludedYears))
                {
                    oDataMgr.m_strSQL = Queries.FVS.FVSInput.StandInit.RemoveLitterYearsSqlite(m_strLitterExcludedYears, strStandTable);
                    DebugLogSQL(oDataMgr.m_strSQL);
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    m_intError = oDataMgr.m_intError;
                }
            }

            //Clean up worktables
            DebugLogMessage("Deleting work tables\r\n", 1);
            if (m_intError == 0)
            {
                string[] strDeleteStmts = Queries.FVS.FVSInput.StandInit.DeleteDwmWorkTables();
                foreach (string strSQL in strDeleteStmts)
                {
                    DebugLogSQL(strSQL);
                    oDataMgr.SqlNonQuery(conn, strSQL);
                    m_intError = oDataMgr.m_intError;
                    if (m_intError != 0)
                        break;
                }
            }
        }

        private void CreateFVSTreeInit()
        {
            using (var conn = new OleDbConnection(m_strConn))
            {
                try
                {
                    conn.Open();
                    int intProgressBarCounter = 0;
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                        "Writing BioSum FVS_TreeInit For Variant " + this.m_strVariant);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Minimum", 0);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Maximum", 31);

                    string strTreeInitWorkTable = "FVS_TreeInit_WorkTable";
                    string strTreeInit = "FVS_TreeInit";

                    //Insert records from Master.mdb
                    string strSQL = Queries.FVS.FVSInput.TreeInit.BulkImportTreeDataFromBioSumMaster(
                        m_strVariant, strTreeInitWorkTable, m_strCondTable, m_strPlotTable, m_strTreeTable);
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);

                    //Updating FIA Species Codes to FVS Species Codes
                    strSQL = Queries.FVS.FVSInput.TreeInit.CreateSpcdConversionTable(m_strCondTable, m_strPlotTable,
                        m_strTreeTable, m_strTreeSpcTable);
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);

                    //Execute the Species code update
                    strSQL = Queries.FVS.FVSInput.TreeInit.UpdateFVSSpeciesCodeColumn(m_strVariant,
                        strTreeInitWorkTable);
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);

                    //Dead trees don't have Compacted Crown Ratio logic in text file approach, so set CrRatio to null where history=9
                    strSQL = Queries.FVS.FVSInput.TreeInit.DeleteCrRatiosForDeadTrees(strTreeInitWorkTable);
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);

                    //Round Cr<10 to Crown Ratio Class 1 (so that FVS rounds it up to 5% to make it the middle of the threshold)
                    strSQL =
                        Queries.FVS.FVSInput.TreeInit.RoundSingleDigitPercentageCrRatiosDownTo1(strTreeInitWorkTable);
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);

                    //If Htcd not in {1,2,3,4} then set the Ht and HtTopK to 0
                    strSQL = Queries.FVS.FVSInput.TreeInit.DeleteHtAndHtTopKForUnknownHtcd(strTreeInitWorkTable);
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);

                    //Calculating Broken top using Ht > ActualHt (HtTopK) before setting HtTopK>=Ht to 0.
                    //The 0<HtTopK means we could execute this after setting HtTopK's to 0
                    //Broken tops can determine damage codes
                    strSQL = Queries.FVS.FVSInput.TreeInit.SetBrokenTopFlag(strTreeInitWorkTable);
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);

                    //If HtTopK >= Ht, set it to 0
                    strSQL = Queries.FVS.FVSInput.TreeInit.SetHtTopKToZeroIfGteHt(strTreeInitWorkTable);
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);

                    //Set Dbh to 0.1 if Tpa > 25 and dbh <= 0 and live tree (implies seedlings) 
                    strSQL = Queries.FVS.FVSInput.TreeInit.SetInferredSeedlingDbh(strTreeInitWorkTable);
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);

                    //Pad FVS_TreeInit.Species with 0 in case it's not 3-digits
                    strSQL = Queries.FVS.FVSInput.TreeInit.PadSpeciesWithZero(strTreeInitWorkTable);
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);

                    //Damage code section
                    string[] strDamageCodes = Queries.FVS.FVSInput.TreeInit.DamageCodes(strTreeInitWorkTable);
                    foreach (string strDamageCodeSQL in strDamageCodes)
                    {
                        m_ado.SqlNonQuery(conn, strDamageCodeSQL);
                        frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                            intProgressBarCounter++);
                    }

                    string[] strTreeValues = Queries.FVS.FVSInput.TreeInit.TreeValueClass(strTreeInitWorkTable);
                    foreach (string strTreeValueUpdate in strTreeValues)
                    {
                        m_ado.SqlNonQuery(conn, strTreeValueUpdate);
                        frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                            intProgressBarCounter++);
                    }

                    //Calibration Data
                    //if (bUseGrmCalibrationData)
                    //{
                    //    //Set Tree History: Died during observation if Micr_Component_Al_Forest='MORTALITY1'
                    //    m_ado.m_strSQL = String.Format(
                    //        "UPDATE {0} fvs INNER JOIN {1} grm " +
                    //        "ON fvs.Stand_ID=grm.biosum_cond_id AND fvs.Tree_ID=grm.fvs_tree_id " +
                    //        "SET fvs.History=7 WHERE grm.micr_component_al_forest='MORTALITY1';"
                    //        , strTreeInitWorkTable, m_strGrmTreeTable);
                    //    DebugLogSQL(m_ado.m_strSQL);
                    //    m_ado.SqlNonQuery(conn, m_ado.m_strSQL);

                    //    //Set DG and HTG to previous measurements if Micr_Component_Al_Forest='SURVIVOR'
                    //    m_ado.m_strSQL = String.Format(
                    //        "UPDATE {0} fvs INNER JOIN {1} grm " +
                    //        "ON fvs.Stand_ID=grm.biosum_cond_id AND fvs.Tree_ID=grm.fvs_tree_id " +
                    //        "SET fvs.DG=IIF(grm.dia_begin IS NOT NULL, grm.dia_begin, NULL), " +
                    //        "fvs.HTG=IIF(grm.ht_begin IS NOT NULL, grm.ht_begin, NULL) " +
                    //        "WHERE grm.micr_component_al_forest='SURVIVOR';", strTreeInitWorkTable, m_strGrmTreeTable);
                    //    DebugLogSQL(m_ado.m_strSQL);
                    //    m_ado.SqlNonQuery(conn, m_ado.m_strSQL);
                    //}

                    //Insert into linked FVS_TreeInit after doing intermediate work in the work table
                    strSQL = Queries.FVS.FVSInput.TreeInit.TranslateWorkTableToTreeInitTable(strTreeInitWorkTable,
                        strTreeInit);
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);

                    //Pad and trim the Species column again so FVS works with it properly
                    strSQL = Queries.FVS.FVSInput.TreeInit.PadSpeciesWithZero(strTreeInit);
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);

                    //Delete work tables
                    strSQL = Queries.FVS.FVSInput.TreeInit.DeleteWorkTable();
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);
                    strSQL = Queries.FVS.FVSInput.TreeInit.DeleteSpcdChangeWorkTable();
                    m_ado.SqlNonQuery(conn, strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value",
                        intProgressBarCounter++);
                }
                catch (Exception e)
                {
                    MessageBox.Show("!!Error!! \n" +
                                    "Module - fvs_input:CreateFVSTreeInit  \n" +
                                    "Err Msg - " + e.Message.Trim(),
                        "FVS Input", System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Exclamation);
                    this.m_intError = -1;
                }
            }
        }


        private void CheckDir()
        {
            try
            {

                if (!System.IO.Directory.Exists(this.m_strInDir))
                    System.IO.Directory.CreateDirectory(this.m_strInDir);
            }
            catch (Exception e)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - fvs_input:CheckDir  \n" +
                    "Err Msg - " + e.Message.Trim(),
                    "FVS Input", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
        }
        private string IDB_getInvYear(string p_strInvId)
        {
            switch (p_strInvId.Trim())
            {
                case "0010":
                    return "1993";
                case "0011":
                    return "1993";
                case "0012":
                    return "1997";
                case "0013":
                    return "1998";
                case "0014":
                    return "1995";
                case "0015":
                    return "1991";
                case "0016":
                    return "1991";
                case "0017":
                    return "1988";
                case "0018":
                    return "   @";
                case "0019":
                    return "   @";
            }
            return "   @";
        }
        private int IDB_getStandAge(int intStdAge)
        {
            switch (intStdAge)
            {
                case 1: return 5;
                case 2: return 15;
                case 3: return 25;
                case 4: return 35;
                case 5: return 45;
                case 6: return 55;
                case 7: return 65;
                case 8: return 75;
                case 9: return 85;
                case 10: return 95;
                case 11: return 105;
                case 12: return 115;
                case 13: return 125;
                case 14: return 135;
                case 15: return 145;
                case 16: return 155;
                case 17: return 165;
                case 18: return 175;
                case 19: return 185;
                case 20: return 195;
                case 21: return 250;
                case 22: return 350;
            }
            return 0;
        }
        private string IDB_getStateCd(string p_strStateCd)
        {
            switch (p_strStateCd)
            {
                case "41": return "OR";
                case "53": return "WA";
                case "6": return "CA";
                case "1": return "AL";
                case "2": return "AK";
                case "4": return "AZ";
                case "5": return "AR";
                case "8": return "CO";
                case "9": return "CT";
                case "10": return "DE";
                case "11": return "DC";
                case "12": return "FL";
                case "13": return "GA";
                case "15": return "HI";
                case "16": return "ID";
                case "17": return "IL";
                case "18": return "IN";
                case "19": return "IA";
                case "20": return "KS";
                case "21": return "KY";
                case "22": return "LA";
                case "23": return "ME";
                case "24": return "MD";
                case "25": return "MA";
                case "26": return "MI";
                case "27": return "MN";
                case "28": return "MS";
                case "29": return "MO";
                case "30": return "MT";
                case "31": return "NE";
                case "32": return "NV";
                case "33": return "NH";
                case "34": return "NJ";
                case "35": return "NM";
                case "36": return "NY";
                case "37": return "NC";
                case "38": return "ND";
                case "39": return "OH";
                case "40": return "OK";
                case "42": return "PA";
                case "44": return "RI";
                case "45": return "SC";
                case "46": return "SD";
                case "47": return "TN";
                case "48": return "TX";
                case "49": return "UT";
                case "50": return "VT";
                case "51": return "VA";
                case "54": return "WV";
                case "55": return "WI";
                case "56": return "WY";
                case "72": return "PR";
                case "78": return "VI";
            }
            return "";
        }
        /// <summary>
        /// delete loc, slf, and fvs files from the variant directory
        /// </summary>
        private void DeleteFiles()
        {

            string strCurrDir = System.IO.Directory.GetCurrentDirectory();
            System.IO.Directory.SetCurrentDirectory(this.m_strInDir);
            string[] strFiles = System.IO.Directory.GetFiles(this.m_strInDir, "*.fvs");
            foreach (string strFile in strFiles)
            {
                System.IO.File.Delete(strFile.Trim());
            }
            System.IO.File.Delete(this.m_strVariant.Trim() + ".loc");
            System.IO.File.Delete(this.m_strVariant.Trim() + ".slf");
            System.IO.File.Delete(this.m_strFVSInMDBFile);
            System.IO.Directory.SetCurrentDirectory(strCurrDir);

        }

        /// <summary>
        /// full directory path and file name to the fvsin mdb file
        /// </summary>
        /// <returns>string name of the Input MDB File</returns>
        public string strFVSInDBFile
        {
            set
            {
                this.m_strFVSInMDBFile = value;
            }
            get
            {
                return this.m_strFVSInMDBFile;
            }
        }


        public void DebugLogMessage(string strMessage)
        {
            if (frmMain.g_bDebug)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, strMessage);
            }
        }
        public void DebugLogMessage(string strMessage, int intDebugLevel)
        {

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > intDebugLevel)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, strMessage);
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
        }
        public void DebugLogSQL(string strSQL)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, strSQL + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
        }

        /// <summary>
        /// Set the DWM behavior in the FVS_StandInit creation process.
        /// Use the Enum 
        /// </summary>
        /// <returns>string name of the Input MDB File</returns>
        public int intDWMOption
        {
            set
            {
                m_intDWMOption = value;
            }
            get
            {
                return m_intDWMOption;
            }
        }

        public string strMinSmallFwdTransectLengthTotal
        {
            set { m_strMinSmallFwdTransectLengthTotal = value; }
            get { return m_strMinSmallFwdTransectLengthTotal; }
        }

        public string strMinLargeFwdTransectLengthTotal
        {
            set { m_strMinLargeFwdTransectLengthTotal = value; }
            get { return m_strMinLargeFwdTransectLengthTotal; }
        }

        public string strMinCwdTransectLengthTotal
        {
            set { m_strMinCwdTransectLengthTotal = value; }
            get { return m_strMinCwdTransectLengthTotal; }
        }

        public string strMinDuffLitterPitCount
        {
            set { m_strMinDuffLitterPitCount = value; }
            get { return m_strMinDuffLitterPitCount; }
        }

        public string strDuffExcludedYears
        {
            set { m_strDuffExcludedYears = value; }
            get { return m_strDuffExcludedYears; }
        }

        public string strLitterExcludedYears
        {
            set { m_strLitterExcludedYears = value; }
            get { return m_strLitterExcludedYears; }
        }

        public bool bUsePrevHtCalibrationData
        {
            set { m_bUsePrevHt = value; }
            get { return m_bUsePrevHt; }
        }

        public bool bUsePrevDiaCalibrationData
        {
            set { m_bUsePrevDia = value; }
            get { return m_bUsePrevDia; }
        }

        public bool bIncludeSeedlings
        {
            set { m_bIncludeSeedlings = value; }
            get { return m_bIncludeSeedlings; }
        }

        public string strSourceFiaDb
        {
            set { m_strSourceFiaDb = value; }
            get { return m_strSourceFiaDb; }
        }

        public string strDataDirectory
        {
            set { m_strDataDirectory = value; }
            get { return m_strDataDirectory; }
        }

        public string strGroup
        {
            set { m_strGroup = value; }
            get { return m_strGroup; }
        }

        /// <summary>
        /// Get the maximum dead tree id for a biosum_cond_id
        /// </summary>
        /// <param name="p_strBiosum_Cond_Id"></param>
        /// <returns></returns>
        private int GetMaximumDeadTreeId(string p_strBiosum_Cond_Id)
        {
            int intValue = (int)m_ado.getSingleDoubleValueFromSQLQuery(
               m_ado.m_OleDbConnection,
               "SELECT MAX(VAL(treeid)) FROM fvs_tree_id_work_table WHERE biosum_cond_id='" +
               p_strBiosum_Cond_Id.Trim() + "' AND VAL(treeid) > 499",
               "fvs_tree_id_work_table");
            if (intValue == null) intValue = -1;
            return intValue;
        }
        /// <summary>
        /// Get the maximum live tree id for a biosum_cond_id
        /// </summary>
        /// <param name="p_strBiosum_Cond_Id"></param>
        /// <returns></returns>
        private int GetMaximumLiveTreeId(string p_strBiosum_Cond_Id)
        {
            int intValue = (int)m_ado.getSingleDoubleValueFromSQLQuery(
               m_ado.m_OleDbConnection,
               "SELECT MAX(VAL(treeid)) FROM fvs_tree_id_work_table WHERE biosum_cond_id='" +
               p_strBiosum_Cond_Id.Trim() + "' AND VAL(treeid) < 500",
               "fvs_tree_id_work_table");
            if (intValue == null) intValue = -1;
            return intValue;
        }
        /// <summary>
        /// Get the maximum assigned plot id in order to assign the next new plot id
        /// </summary>
        /// <returns></returns>
        private int GetMaximumPlotId()
        {
            int intValue = (int)m_ado.getSingleDoubleValueFromSQLQuery(
                m_ado.m_OleDbConnection,
                "SELECT MAX(VAL(plotid))  as max_plotid FROM fvs_tree_id_work_table",
                "fvs_tree_id_work_table");
            if (intValue == null) intValue = -1;
            return intValue;
        }
        /// <summary>
        /// find the 1st unused plot id
        /// </summary>
        /// <returns></returns>
        private int GetAvailablePlotId()
        {
            int intValue = -1;
            string strPlotId = "";
            int x;
            for (x = 1; x <= 9999; x++)
            {
                strPlotId = Convert.ToString(x).PadLeft(4, '0');
                intValue = (int)m_ado.getSingleDoubleValueFromSQLQuery(m_ado.m_OleDbConnection,
                    "SELECT COUNT(*) as plotid_count FROM fvs_tree_id_work_table " +
                    "WHERE plotid='" + strPlotId.Trim() + "'", "temp");
                if (intValue == 0)
                {
                    break;
                }
            }
            intValue = x;
            return intValue;
        }

        /// <summary>
        /// find the 1st unused live tree id
        /// </summary>
        /// <returns></returns>
        private int GetAvailableLiveTreeId(string p_strBiosum_Cond_Id)
        {
            int intValue = -1;
            string strTreeId = "";
            int x;
            for (x = 1; x <= 499; x++)
            {
                strTreeId = Convert.ToString(x).PadLeft(3, '0');
                intValue = (int)m_ado.getSingleDoubleValueFromSQLQuery(m_ado.m_OleDbConnection,
                    "SELECT COUNT(*) as treeid_count FROM fvs_tree_id_work_table " +
                    "WHERE biosum_cond_id='" + p_strBiosum_Cond_Id + "' AND " +
                          "treeid='" + strTreeId.Trim() + "'", "temp");
                if (intValue == 0)
                {
                    break;
                }
            }
            intValue = x;
            return intValue;
        }
        /// <summary>
        /// find the 1st unused dead tree id
        /// </summary>
        /// <returns></returns>
        private int GetAvailableDeadTreeId(string p_strBiosum_Cond_Id)
        {
            int intValue = -1;
            string strTreeId = "";
            int x;
            for (x = 500; x <= 999; x++)
            {
                strTreeId = Convert.ToString(x).PadLeft(3, '0');
                intValue = (int)m_ado.getSingleDoubleValueFromSQLQuery(m_ado.m_OleDbConnection,
                    "SELECT COUNT(*) as treeid_count FROM fvs_tree_id_work_table " +
                    "WHERE biosum_cond_id='" + p_strBiosum_Cond_Id + "' AND " +
                          "treeid='" + strTreeId.Trim() + "'", "temp");
                if (intValue == 0)
                {
                    break;
                }
            }
            intValue = x;
            return intValue;
        }

        /// <summary>
        /// Site index functions beginning with "z" were programmed by Tara: these have been checked against the publications.
        /// Site index functions beginning with "q" were taken from Bruce Hiserote's Visual Basic program 4/2002: these should be fine.
        ///Other site index functions taken from FSVEG group - Kurt Campbell: these have not been checked for errors.
        ///Any changes from the sources are noted in the comments of each function.
        ///Other Site index equations come from the publication
        ///"Site Index Equations and Mean Annual Increment Equations
        ///for Pacific Northwest Research Station Forest Inventory and 
        ///Analysis Inventories, 1985-2001" 
        ///Authors: Erica Hanson, David Azuma, and Bruce Hiserote
        ///Research Note: PNW-RN533 December 2002
        /// </summary>
        private class site_index
        {
            double _dblCCTreeBasalAreaPerAcre;
            double _dblCCAvgDia;
            string _strStateCd = "";
            string _strCountyCd = "";
            string _strPlot = "";
            string _strCondId = "";
            string _strBiosumPlotId = "";
            ado_data_access _oAdo;
            string _strPlotTable = "";
            string _strTreeTable = "";
            string _strCondTable = "";
            string _strSiteTreeTable = "";
            string _strTreeSpeciesTable = "";
            string _strFVSTreeSpcTable = "";
            string _strFVSVariant = "";
            string _strSiteIndexSpecies = "";
            string _strSiteIndex = "";
            string _strBaseAge = "";
            string _strSiteIndexSpeciesAlphaCode = "";
            string _strCCHabitatTypeCd;
            IDictionary<String, String> _dictSiteIdxEq;
            string _strDebugFile;
            string _strTempMDBFileNewProcess;
            SQLite.ADO.DataMgr _oDataMgr;
            System.Data.SQLite.SQLiteConnection _tempConn;
            string _strProjDir = "";

            bool _bProcess = true;

            public site_index()
            {
            }

            public string StateCd
            {
                get { return _strStateCd; }
                set { _strStateCd = value; }
            }
            public string CountyCd
            {
                get { return _strCountyCd; }
                set { _strCountyCd = value; }
            }
            public string Plot
            {
                get { return _strPlot; }
                set { _strPlot = value; }
            }
            public string CondId
            {
                get { return _strCondId; }
                set { _strCondId = value; }
            }
            public string BiosumPlotId
            {
                get { return _strBiosumPlotId; }
                set { _strBiosumPlotId = value; }
            }
            public string FVSVariant
            {
                get { return _strFVSVariant; }
                set { _strFVSVariant = value; }
            }
            public string PlotTable
            {
                get { return _strPlotTable; }
                set { _strPlotTable = value; }
            }
            public string CondTable
            {
                get { return _strCondTable; }
                set { _strCondTable = value; }
            }
            public string TreeTable
            {
                get { return _strTreeTable; }
                set { _strTreeTable = value; }
            }
            public string SiteTreeTable
            {
                get { return _strSiteTreeTable; }
                set { _strSiteTreeTable = value; }
            }
            public string TreeSpeciesTable
            {
                get { return _strTreeSpeciesTable; }
                set { _strTreeSpeciesTable = value; }
            }
            public string FVSTreeSpeciesTable
            {
                get { return _strFVSTreeSpcTable; }
                set { _strFVSTreeSpcTable = value; }
            }
            public ado_data_access ado_data_access
            {
                set { this._oAdo = value; }
                get { return _oAdo; }
            }
            public bool Process
            {
                set { _bProcess = value; }
                get { return _bProcess; }
            }
            public string SiteIndexSpecies
            {
                get { return _strSiteIndexSpecies; }
                set { _strSiteIndexSpecies = value; }
            }
            public string SiteIndex
            {
                get { return _strSiteIndex; }
                set { _strSiteIndex = value; }
            }
            public string BaseAge
            {
                get { return _strBaseAge; }
                set { _strBaseAge = value; }
            }
            public string SiteIndexSpeciesAlphaCode
            {
                get { return _strSiteIndexSpeciesAlphaCode; }
                set { _strSiteIndexSpeciesAlphaCode = value; }
            }
            public double ConditionClassBasalAreaPerAcre
            {
                get { return _dblCCTreeBasalAreaPerAcre; }
                set { _dblCCTreeBasalAreaPerAcre = value; }
            }
            public double ConditionClassAverageDia
            {
                get { return _dblCCAvgDia; }
                set { _dblCCAvgDia = value; }
            }
            public string ConditionClassHabitatTypeCd
            {
                get { return _strCCHabitatTypeCd; }
                set { _strCCHabitatTypeCd = value; }
            }
            public IDictionary<String, String> SiteIndexEquations
            {
                set { _dictSiteIdxEq = value; }
            }
            public string DebugFile
            {
                set { _strDebugFile = value; }
            }
            public string TempMDBFileNewProcess
            {
                set { _strTempMDBFileNewProcess = value; }
            }
            public System.Data.SQLite.SQLiteConnection tempConn
            {
                set { _tempConn = value; }
                get { return _tempConn; }
            }
            public SQLite.ADO.DataMgr dataMgr
            {
                set { _oDataMgr = value; }
                get { return _oDataMgr; }
            }
            public string ProjectDirectory
            {
                get { return _strProjDir; }
                set { _strProjDir = value; }
            }

            public void getSiteIndex_access(System.Data.DataRow p_oRow)
            {
                //biosum plot id
                if (p_oRow["biosum_plot_id"] == System.DBNull.Value)
                    this.BiosumPlotId = "";
                else this.BiosumPlotId = p_oRow["biosum_plot_id"].ToString().Trim();
                //statecd
                if (p_oRow["statecd"] == System.DBNull.Value)
                    this.StateCd = "";
                else this.StateCd = Convert.ToString(p_oRow["statecd"]).Trim();
                //countycd
                if (p_oRow["countycd"] == System.DBNull.Value)
                    this.CountyCd = "";
                else this.CountyCd = Convert.ToString(p_oRow["countycd"]).Trim();
                //plot
                if (p_oRow["plot"] == System.DBNull.Value)
                    this.Plot = "";
                else this.Plot = Convert.ToString(p_oRow["plot"]).Trim();
                //fvs variant
                if (p_oRow["fvs_variant"] == System.DBNull.Value)
                    this.FVSVariant = "";
                else this.FVSVariant = Convert.ToString(p_oRow["fvs_variant"]).Trim();
                //cond id
                if (p_oRow["condid"] == System.DBNull.Value)
                    this.CondId = "";
                else this.CondId = Convert.ToString(p_oRow["condid"]).Trim();
                //habitat type code
                if (p_oRow["habtypcd1"] == System.DBNull.Value)
                    this.ConditionClassHabitatTypeCd = "";
                else this.ConditionClassHabitatTypeCd = Convert.ToString(p_oRow["habtypcd1"]).Trim();

                //15-JUN-2015: We now calculate this in a function rather than populate from COND table so we can control the parameters
                //tree basal area per acre on the condition
                //if (p_oRow["ba_ft2_ac"] != System.DBNull.Value)
                //    this.ConditionClassBasalAreaPerAcre=Convert.ToDouble(p_oRow["ba_ft2_ac"]);

                getSiteIndex_access();
            }
            public void getSiteIndex(System.Data.DataRow p_oRow)
            {
                //biosum plot id
                if (p_oRow["biosum_plot_id"] == System.DBNull.Value)
                    this.BiosumPlotId = "";
                else this.BiosumPlotId = p_oRow["biosum_plot_id"].ToString().Trim();
                //statecd
                if (p_oRow["statecd"] == System.DBNull.Value)
                    this.StateCd = "";
                else this.StateCd = Convert.ToString(p_oRow["statecd"]).Trim();
                //countycd
                if (p_oRow["countycd"] == System.DBNull.Value)
                    this.CountyCd = "";
                else this.CountyCd = Convert.ToString(p_oRow["countycd"]).Trim();
                //plot
                if (p_oRow["plot"] == System.DBNull.Value)
                    this.Plot = "";
                else this.Plot = Convert.ToString(p_oRow["plot"]).Trim();
                //fvs variant
                if (p_oRow["fvs_variant"] == System.DBNull.Value)
                    this.FVSVariant = "";
                else this.FVSVariant = Convert.ToString(p_oRow["fvs_variant"]).Trim();
                //cond id
                if (p_oRow["condid"] == System.DBNull.Value)
                    this.CondId = "";
                else this.CondId = Convert.ToString(p_oRow["condid"]).Trim();
                //habitat type code
                if (p_oRow["habtypcd1"] == System.DBNull.Value)
                    this.ConditionClassHabitatTypeCd = "";
                else this.ConditionClassHabitatTypeCd = Convert.ToString(p_oRow["habtypcd1"]).Trim();

                getSiteIndex();
            }

            public void getSiteIndexNew(System.Data.DataRow p_oRow)
            {
                //biosum plot id
                if (p_oRow["biosum_plot_id"] == System.DBNull.Value)
                    this.BiosumPlotId = "";
                else this.BiosumPlotId = p_oRow["biosum_plot_id"].ToString().Trim();
                //statecd
                if (p_oRow["statecd"] == System.DBNull.Value)
                    this.StateCd = "";
                else this.StateCd = Convert.ToString(p_oRow["statecd"]).Trim();
                //countycd
                if (p_oRow["countycd"] == System.DBNull.Value)
                    this.CountyCd = "";
                else this.CountyCd = Convert.ToString(p_oRow["countycd"]).Trim();
                //plot
                if (p_oRow["plot"] == System.DBNull.Value)
                    this.Plot = "";
                else this.Plot = Convert.ToString(p_oRow["plot"]).Trim();
                //fvs variant
                if (p_oRow["fvs_variant"] == System.DBNull.Value)
                    this.FVSVariant = "";
                else this.FVSVariant = Convert.ToString(p_oRow["fvs_variant"]).Trim();
                //cond id
                if (p_oRow["condid"] == System.DBNull.Value)
                    this.CondId = "";
                else this.CondId = Convert.ToString(p_oRow["condid"]).Trim();
                //habitat type code
                if (p_oRow["habtypcd1"] == System.DBNull.Value)
                    this.ConditionClassHabitatTypeCd = "";
                else this.ConditionClassHabitatTypeCd = Convert.ToString(p_oRow["habtypcd1"]).Trim();

                getSiteIndexNew();
            }

            private void getSiteIndex_access()
            {
                int x, y;
                int intCount;

                //MAX variables hold the values of the selected site index
                int intSICountMax;
                double dblSIAvgMax;
                int intSISpeciesMax;
                int intCurSIFVSSpecies;
                int intCurFIASpecies;
                int intCurHtFt;
                int intCurAgeDia;
                int intCondId;
                bool bFound;
                int intSiTree;

                //These arrays contain the values of all the site index trees on the plot
                int[] intSIFVSSpecies;
                int[] intSICount;
                double[] dblSISum;
                double[] dblSIAvg;

                this.SiteIndex = "@";
                this.SiteIndexSpecies = "@";
                this.SiteIndexSpeciesAlphaCode = "@";

                double dblSiteIndex = 0;


                //calculate site index for OR, WA, CA, ID, and MT
                if (StateCd == "41" || StateCd == "6" ||
                    StateCd == "53" || StateCd == "16" ||
                    StateCd == "30")
                {
                }
                else return;

                //get all the site index trees that are applied to the current plot+condition
                _oAdo.m_strSQL = "SELECT s.biosum_plot_id," +
                    "s.condid," +
                    "s.tree," +
                    "s.spcd," +
                    "s.dia," +
                    "s.ht," +
                    "s.agedia," +
                    "s.subp," +
                    "s.method," +
                    "s.validcd, " +
                    "s.sitree " +
                    "FROM " + this.SiteTreeTable + " s " +
                    "WHERE s.biosum_plot_id = '" + this.BiosumPlotId + "' " +
                    "AND s.condid = " + this.CondId +
                    "AND s.validcd <> 0";
                x = Convert.ToInt32(_oAdo.getRecordCount(_oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM (" + _oAdo.m_strSQL + ")", "cond"));
                if (x > 0)
                {
                    _oAdo.AddSQLQueryToDataSet(_oAdo.m_OleDbConnection, ref _oAdo.m_OleDbDataAdapter, ref _oAdo.m_DataSet, _oAdo.m_strSQL, "GetSiteIndex");
                    if (_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows.Count > 0)
                    {
                        intSIFVSSpecies = new int[x];
                        intSICount = new int[x];
                        dblSISum = new double[x];
                        intCount = 0;
                        intSICountMax = 0;
                        for (y = 0; y <= _oAdo.m_DataSet.Tables["GetSiteIndex"].Rows.Count - 1; y++)
                        {
                            intCurFIASpecies = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["spcd"]);
                            intCurSIFVSSpecies = 0;
                            intCurAgeDia = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["agedia"]);
                            intCurHtFt = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["ht"]);
                            intCondId = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["condid"]);
                            intSiTree = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["sitree"]);

                            //***************************************************
                            //**if no age then bypass site index tree
                            //**************************************************
                            if (intCurAgeDia > 0)
                            {
                                LoadSiteIndexValues_access(intCondId, intCurFIASpecies,
                                    intCurAgeDia,
                                    intCurHtFt,
                                    ref intCurSIFVSSpecies,
                                    ref dblSiteIndex,
                                    intSiTree);
                                //*************************************************
                                //**if the site index = 0, write it to the log, we want to know
                                //**how often this occurs
                                //*************************************************
                                if (dblSiteIndex == 0)
                                {
                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    {
                                        string logEntry = "//variant: " + this.FVSVariant +
                                                          " plot id: " + this.BiosumPlotId +
                                                          " cond id: " + this.CondId +
                                                          " spec cd: " + intCurSIFVSSpecies + "\r\n";
                                        frmMain.g_oUtils.WriteText(_strDebugFile, "\r\n//\r\n");
                                        frmMain.g_oUtils.WriteText(_strDebugFile, "//Site_Index_getSiteIndex\r\n");
                                        frmMain.g_oUtils.WriteText(_strDebugFile, "//Site index equation returned 0    \r\n");
                                        frmMain.g_oUtils.WriteText(_strDebugFile, logEntry);
                                        frmMain.g_oUtils.WriteText(_strDebugFile, "//\r\n");
                                    }
                                }
                                //*************************************************
                                //**lets find the current SI species in the array
                                //*************************************************
                                if (intCount == 0)
                                {
                                    intCount = intCount + 1;
                                    intSIFVSSpecies[intCount - 1] = intCurSIFVSSpecies;
                                    intSICount[intCount - 1] = intSICount[intCount - 1] + 1;
                                    dblSISum[intCount - 1] = dblSISum[intCount - 1] + dblSiteIndex;


                                }
                                else if (intSIFVSSpecies[intCount - 1] == intCurSIFVSSpecies)
                                {
                                    intSICount[intCount - 1] = intSICount[intCount - 1] + 1;
                                    dblSISum[intCount - 1] = dblSISum[intCount - 1] + dblSiteIndex;
                                }
                                else
                                {
                                    bFound = false;
                                    for (x = 0; x <= intCount - 1; x++)
                                    {
                                        if (intSIFVSSpecies[x] == intCurSIFVSSpecies)
                                        {
                                            bFound = true;
                                            break;
                                        }
                                    }
                                    if (bFound)
                                    {
                                        intSICount[x] = intSICount[x] + 1;
                                        dblSISum[x] = dblSISum[x] + dblSiteIndex;
                                    }
                                    else
                                    {
                                        intCount = intCount + 1;
                                        intSIFVSSpecies[intCount - 1] = intCurSIFVSSpecies;
                                        intSICount[intCount - 1] = intSICount[intCount - 1] + 1;
                                        dblSISum[intCount - 1] = dblSISum[intCount - 1] + dblSiteIndex;
                                    }
                                }
                            }
                        }
                        //***************************************************************
                        //**get the most frequently occuring site index species
                        //***************************************************************
                        intSICountMax = 0;
                        dblSIAvgMax = 0;
                        intSISpeciesMax = 0;
                        dblSIAvg = new double[intCount];
                        for (x = 0; x <= intCount - 1; x++)
                        {
                            dblSIAvg[x] = dblSISum[x] / intSICount[x];
                            if (intSICount[x] > intSICountMax)
                            {
                                intSICountMax = intSICount[x];
                                dblSIAvgMax = dblSIAvg[x];
                                intSISpeciesMax = intSIFVSSpecies[x];
                            }
                            else if (intSICount[x] == intSICountMax)
                            {
                                if (dblSIAvg[x] > dblSIAvgMax)
                                {
                                    dblSIAvgMax = dblSIAvg[x];
                                    intSISpeciesMax = intSIFVSSpecies[x];
                                }
                            }
                        }
                        if (dblSIAvgMax <= 0 && intSISpeciesMax > 0 && intSISpeciesMax != 999)
                        {
                            MessageBox.Show("Warning: Site index tree species " + Convert.ToString(intSISpeciesMax) + " has an invalid  site index value of " + Convert.ToString(Math.Round(dblSIAvgMax, 6)).Trim() + ". Both the SI species and SI height will be given a value of @");
                            this.SiteIndexSpecies = "@";
                            this.SiteIndex = "@";

                        }
                        else
                        {
                            this.SiteIndexSpecies = intSISpeciesMax.ToString().Trim();
                            this.SiteIndex = Convert.ToString(Math.Round(dblSIAvgMax, 0)).Trim();
                        }
                        if (this.SiteIndexSpecies != "@" && this.SiteIndexSpecies.Trim().Length > 0)
                            GetSiteIndexSpeciesAlphaCode_access();
                    }
                    _oAdo.m_DataSet.Tables.Remove("GetSiteIndex");

                }
            }

            private void getSiteIndex()
            {
                int x, y;
                int intCount;

                //MAX variables hold the values of the selected site index
                int intSICountMax;
                double dblSIAvgMax;
                int intSISpeciesMax;
                int intCurSIFVSSpecies;
                int intCurFIASpecies;
                int intCurHtFt;
                int intCurAgeDia;
                int intCondId;
                bool bFound;
                int intSiTree;

                //These arrays contain the values of all the site index trees on the plot
                int[] intSIFVSSpecies;
                int[] intSICount;
                double[] dblSISum;
                double[] dblSIAvg;

                this.SiteIndex = "@";
                this.SiteIndexSpecies = "@";
                this.SiteIndexSpeciesAlphaCode = "@";
                this.BaseAge = "@";

                double dblSiteIndex = 0;

                SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
                bool variantFound = false;

                // Calculate site index for variants in site_index_equations
                string strEqnConn = oDataMgr.GetConnectionString(frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceSqliteFile);
                using (System.Data.SQLite.SQLiteConnection eqnConn = new System.Data.SQLite.SQLiteConnection(strEqnConn))
                {
                    eqnConn.Open();

                    oDataMgr.m_strSQL = "SELECT DISTINCT FVS_VARIANT FROM " + Tables.Reference.DefaultSiteIndexEquationsTable;
                    oDataMgr.SqlQueryReader(eqnConn, oDataMgr.m_strSQL);
                    if (oDataMgr.m_DataReader.HasRows)
                    {
                        while (oDataMgr.m_DataReader.Read())
                        {
                            if (Convert.ToString(oDataMgr.m_DataReader["FVS_VARIANT"]) == FVSVariant)
                            {
                                variantFound = true;
                                break;
                            }
                        }
                    }
                }
                if (!variantFound)
                {
                    return;
                }

                using (OleDbConnection conn = new OleDbConnection(_oAdo.getMDBConnString(_strTempMDBFileNewProcess, "", "")))
                {
                    conn.Open();
                    //get all the site index trees that are applied to the current plot+condition
                    _oAdo.m_strSQL = "SELECT s.biosum_plot_id," +
                        "s.condid," +
                        "s.tree," +
                        "s.spcd," +
                        "s.dia," +
                        "s.ht," +
                        "s.agedia," +
                        "s.subp," +
                        "s.method," +
                        "s.validcd, " +
                        "s.sitree " +
                        "FROM " + this.SiteTreeTable + " s " +
                        "WHERE s.biosum_plot_id = '" + this.BiosumPlotId + "' " +
                        "AND s.condid = " + this.CondId +
                        " AND s.validcd <> 0";
                    x = Convert.ToInt32(_oAdo.getRecordCount(conn, "SELECT COUNT(*) FROM (" + _oAdo.m_strSQL + ")", "cond"));
                    if (x > 0)
                    {
                        _oAdo.AddSQLQueryToDataSet(conn, ref _oAdo.m_OleDbDataAdapter, ref _oAdo.m_DataSet, _oAdo.m_strSQL, "GetSiteIndex");
                        if (_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows.Count > 0)
                        {
                            intSIFVSSpecies = new int[x];
                            intSICount = new int[x];
                            dblSISum = new double[x];
                            intCount = 0;
                            intSICountMax = 0;
                            for (y = 0; y <= _oAdo.m_DataSet.Tables["GetSiteIndex"].Rows.Count - 1; y++)
                            {
                                intCurFIASpecies = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["spcd"]);
                                intCurSIFVSSpecies = 0;
                                intCurAgeDia = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["agedia"]);
                                intCurHtFt = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["ht"]);
                                intCondId = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["condid"]);
                                intSiTree = Convert.ToInt32(_oAdo.m_DataSet.Tables["GetSiteIndex"].Rows[y]["sitree"]);

                                //***************************************************
                                //**if no age then bypass site index tree
                                //**************************************************
                                if (intCurAgeDia > 0)
                                {
                                    LoadSiteIndexValues(intCondId, intCurFIASpecies,
                                        intCurAgeDia,
                                        intCurHtFt,
                                        ref intCurSIFVSSpecies,
                                        ref dblSiteIndex,
                                        intSiTree);
                                    //*************************************************
                                    //**if the site index = 0, write it to the log, we want to know
                                    //**how often this occurs
                                    //*************************************************
                                    if (dblSiteIndex == 0)
                                    {
                                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                        {
                                            string logEntry = "//variant: " + this.FVSVariant +
                                                              " plot id: " + this.BiosumPlotId +
                                                              " cond id: " + this.CondId +
                                                              " spec cd: " + intCurSIFVSSpecies + "\r\n";
                                            frmMain.g_oUtils.WriteText(_strDebugFile, "\r\n//\r\n");
                                            frmMain.g_oUtils.WriteText(_strDebugFile, "//Site_Index_getSiteIndex\r\n");
                                            frmMain.g_oUtils.WriteText(_strDebugFile, "//Site index equation returned 0    \r\n");
                                            frmMain.g_oUtils.WriteText(_strDebugFile, logEntry);
                                            frmMain.g_oUtils.WriteText(_strDebugFile, "//\r\n");
                                        }
                                    }
                                    //*************************************************
                                    //**lets find the current SI species in the array
                                    //*************************************************
                                    if (intCount == 0)
                                    {
                                        intCount = intCount + 1;
                                        intSIFVSSpecies[intCount - 1] = intCurSIFVSSpecies;
                                        intSICount[intCount - 1] = intSICount[intCount - 1] + 1;
                                        dblSISum[intCount - 1] = dblSISum[intCount - 1] + dblSiteIndex;


                                    }
                                    else if (intSIFVSSpecies[intCount - 1] == intCurSIFVSSpecies)
                                    {
                                        intSICount[intCount - 1] = intSICount[intCount - 1] + 1;
                                        dblSISum[intCount - 1] = dblSISum[intCount - 1] + dblSiteIndex;
                                    }
                                    else
                                    {
                                        bFound = false;
                                        for (x = 0; x <= intCount - 1; x++)
                                        {
                                            if (intSIFVSSpecies[x] == intCurSIFVSSpecies)
                                            {
                                                bFound = true;
                                                break;
                                            }
                                        }
                                        if (bFound)
                                        {
                                            intSICount[x] = intSICount[x] + 1;
                                            dblSISum[x] = dblSISum[x] + dblSiteIndex;
                                        }
                                        else
                                        {
                                            intCount = intCount + 1;
                                            intSIFVSSpecies[intCount - 1] = intCurSIFVSSpecies;
                                            intSICount[intCount - 1] = intSICount[intCount - 1] + 1;
                                            dblSISum[intCount - 1] = dblSISum[intCount - 1] + dblSiteIndex;
                                        }
                                    }
                                }
                            }
                            //***************************************************************
                            //**get the most frequently occuring site index species
                            //***************************************************************
                            intSICountMax = 0;
                            dblSIAvgMax = 0;
                            intSISpeciesMax = 0;
                            dblSIAvg = new double[intCount];
                            for (x = 0; x <= intCount - 1; x++)
                            {
                                dblSIAvg[x] = dblSISum[x] / intSICount[x];
                                if (intSICount[x] > intSICountMax)
                                {
                                    intSICountMax = intSICount[x];
                                    dblSIAvgMax = dblSIAvg[x];
                                    intSISpeciesMax = intSIFVSSpecies[x];
                                }
                                else if (intSICount[x] == intSICountMax)
                                {
                                    if (dblSIAvg[x] > dblSIAvgMax)
                                    {
                                        dblSIAvgMax = dblSIAvg[x];
                                        intSISpeciesMax = intSIFVSSpecies[x];
                                    }
                                }
                            }
                            if (dblSIAvgMax <= 0 && intSISpeciesMax > 0 && intSISpeciesMax != 999)
                            {
                                MessageBox.Show("Warning: Site index tree species " + Convert.ToString(intSISpeciesMax) + " has an invalid  site index value of " + Convert.ToString(Math.Round(dblSIAvgMax, 6)).Trim() + ". Both the SI species and SI height will be given a value of @");
                                this.SiteIndexSpecies = "@";
                                this.SiteIndex = "@";

                            }
                            else
                            {
                                this.SiteIndexSpecies = intSISpeciesMax.ToString().Trim();
                                this.SiteIndex = Convert.ToString(Math.Round(dblSIAvgMax, 0)).Trim();
                            }
                            if (this.SiteIndexSpecies != "@" && this.SiteIndexSpecies.Trim().Length > 0)
                                GetSiteIndexSpeciesAlphaCode();
                        }
                        _oAdo.m_DataSet.Tables.Remove("GetSiteIndex");

                    }
                }

                
            }
            private void getSiteIndexNew()
            {
                int x, y;
                int intCount;

                //MAX variables hold the values of the selected site index
                int intSICountMax;
                double dblSIAvgMax;
                int intSISpeciesMax;
                int intCurSIFVSSpecies;
                int intCurFIASpecies;
                int intCurHtFt;
                int intCurAgeDia;
                int intCondId;
                bool bFound;
                int intSiTree;

                //These arrays contain the values of all the site index trees on the plot
                int[] intSIFVSSpecies;
                int[] intSICount;
                double[] dblSISum;
                double[] dblSIAvg;

                this.SiteIndex = "@";
                this.SiteIndexSpecies = "@";
                this.SiteIndexSpeciesAlphaCode = "@";
                this.BaseAge = "@";

                double dblSiteIndex = 0;
                bool variantFound = false;

                string strRefDb = frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceSqliteFile;

                if (!dataMgr.DatabaseAttached(tempConn, strRefDb))
                {
                    dataMgr.m_strSQL = "ATTACH DATABASE '" + strRefDb + "' AS ref";
                    dataMgr.SqlNonQuery(tempConn, dataMgr.m_strSQL);
                }

                dataMgr.m_strSQL = "SELECT DISTINCT FVS_VARIANT FROM " + Tables.Reference.DefaultSiteIndexEquationsTable;
                dataMgr.SqlQueryReader(tempConn, dataMgr.m_strSQL);
                if (dataMgr.m_DataReader.HasRows)
                {
                    while (dataMgr.m_DataReader.Read())
                    {
                        if (Convert.ToString(dataMgr.m_DataReader["FVS_VARIANT"]) == FVSVariant)
                        {
                            variantFound = true;
                            break;
                        }
                    }
                    dataMgr.m_DataReader.Close();
                }

                if (!variantFound)
                {
                    return;
                }

                //get all the site index trees that are applied to the current plot+condition
                dataMgr.m_strSQL = "SELECT s.biosum_plot_id," +
                    "s.condid," +
                    "s.tree," +
                    "s.spcd," +
                    "s.dia," +
                    "s.ht," +
                    "s.agedia," +
                    "s.subp," +
                    "s.method," +
                    "s.validcd, " +
                    "s.sitree " +
                    "FROM " + SiteTreeTable + " s " +
                    "WHERE s.biosum_plot_id = '" + BiosumPlotId + "' " +
                    "AND s.condid = " + CondId +
                    " AND s.validcd <> 0";
                x = Convert.ToInt32(dataMgr.getRecordCount(tempConn, "SELECT COUNT(*) FROM (" + dataMgr.m_strSQL + ")", "cond"));

                if (x > 0)
                {
                    dataMgr.AddSQLQueryToDataSet(tempConn, ref dataMgr.m_DataAdapter, ref dataMgr.m_DataSet, dataMgr.m_strSQL, "GetSiteIndex");

                    if (dataMgr.m_DataSet.Tables["GetSiteIndex"].Rows.Count > 0)
                    {
                        intSIFVSSpecies = new int[x];
                        intSICount = new int[x];
                        dblSISum = new double[x];
                        intCount = 0;
                        intSICountMax = 0;

                        for (y = 0; y <= dataMgr.m_DataSet.Tables["GetSiteIndex"].Rows.Count - 1; y++)
                        {
                            intCurFIASpecies = Convert.ToInt32(dataMgr.m_DataSet.Tables["GetSiteIndex"].Rows[y]["spcd"]);
                            intCurSIFVSSpecies = 0;
                            intCurAgeDia = Convert.ToInt32(dataMgr.m_DataSet.Tables["GetSiteIndex"].Rows[y]["agedia"]);
                            intCurHtFt = Convert.ToInt32(dataMgr.m_DataSet.Tables["GetSiteIndex"].Rows[y]["ht"]);
                            intCondId = Convert.ToInt32(dataMgr.m_DataSet.Tables["GetSiteIndex"].Rows[y]["condid"]);
                            intSiTree = Convert.ToInt32(dataMgr.m_DataSet.Tables["GetSiteIndex"].Rows[y]["sitree"]);

                            //***************************************************
                            //**if no age then bypass site index tree
                            //**************************************************
                            if (intCurAgeDia > 0)
                            {
                                LoadSiteIndexValues(intCondId, intCurFIASpecies,
                                        intCurAgeDia,
                                        intCurHtFt,
                                        ref intCurSIFVSSpecies,
                                        ref dblSiteIndex,
                                        intSiTree);

                                //*************************************************
                                //**if the site index = 0, write it to the log, we want to know
                                //**how often this occurs
                                //*************************************************
                                if (dblSiteIndex == 0)
                                {
                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    {
                                        string logEntry = "//variant: " + this.FVSVariant +
                                                          " plot id: " + this.BiosumPlotId +
                                                          " cond id: " + this.CondId +
                                                          " spec cd: " + intCurSIFVSSpecies + "\r\n";
                                        frmMain.g_oUtils.WriteText(_strDebugFile, "\r\n//\r\n");
                                        frmMain.g_oUtils.WriteText(_strDebugFile, "//Site_Index_getSiteIndex\r\n");
                                        frmMain.g_oUtils.WriteText(_strDebugFile, "//Site index equation returned 0    \r\n");
                                        frmMain.g_oUtils.WriteText(_strDebugFile, logEntry);
                                        frmMain.g_oUtils.WriteText(_strDebugFile, "//\r\n");
                                    }
                                }

                                if (intCount == 0)
                                {
                                    intCount = intCount + 1;
                                    intSIFVSSpecies[intCount - 1] = intCurSIFVSSpecies;
                                    intSICount[intCount - 1] = intSICount[intCount - 1] + 1;
                                    dblSISum[intCount - 1] = dblSISum[intCount - 1] + dblSiteIndex;


                                }
                                else if (intSIFVSSpecies[intCount - 1] == intCurSIFVSSpecies)
                                {
                                    intSICount[intCount - 1] = intSICount[intCount - 1] + 1;
                                    dblSISum[intCount - 1] = dblSISum[intCount - 1] + dblSiteIndex;
                                }
                                else
                                {
                                    bFound = false;
                                    for (x = 0; x <= intCount - 1; x++)
                                    {
                                        if (intSIFVSSpecies[x] == intCurSIFVSSpecies)
                                        {
                                            bFound = true;
                                            break;
                                        }
                                    }
                                    if (bFound)
                                    {
                                        intSICount[x] = intSICount[x] + 1;
                                        dblSISum[x] = dblSISum[x] + dblSiteIndex;
                                    }
                                    else
                                    {
                                        intCount = intCount + 1;
                                        intSIFVSSpecies[intCount - 1] = intCurSIFVSSpecies;
                                        intSICount[intCount - 1] = intSICount[intCount - 1] + 1;
                                        dblSISum[intCount - 1] = dblSISum[intCount - 1] + dblSiteIndex;
                                    }
                                }
                            }
                        }

                        //***************************************************************
                        //**get the most frequently occuring site index species
                        //***************************************************************
                        intSICountMax = 0;
                        dblSIAvgMax = 0;
                        intSISpeciesMax = 0;
                        dblSIAvg = new double[intCount];
                        for (x = 0; x <= intCount - 1; x++)
                        {
                            dblSIAvg[x] = dblSISum[x] / intSICount[x];
                            if (intSICount[x] > intSICountMax)
                            {
                                intSICountMax = intSICount[x];
                                dblSIAvgMax = dblSIAvg[x];
                                intSISpeciesMax = intSIFVSSpecies[x];
                            }
                            else if (intSICount[x] == intSICountMax)
                            {
                                if (dblSIAvg[x] > dblSIAvgMax)
                                {
                                    dblSIAvgMax = dblSIAvg[x];
                                    intSISpeciesMax = intSIFVSSpecies[x];
                                }
                            }
                        }

                        if (dblSIAvgMax <= 0 && intSISpeciesMax > 0 && intSISpeciesMax != 999)
                        {
                            MessageBox.Show("Warning: Site index tree species " + Convert.ToString(intSISpeciesMax) + " has an invalid  site index value of " + Convert.ToString(Math.Round(dblSIAvgMax, 6)).Trim() + ". Both the SI species and SI height will be given a value of @");
                            this.SiteIndexSpecies = "@";
                            this.SiteIndex = "@";

                        }
                        else
                        {
                            this.SiteIndexSpecies = intSISpeciesMax.ToString().Trim();
                            this.SiteIndex = Convert.ToString(Math.Round(dblSIAvgMax, 0)).Trim();
                        }
                    }

                    dataMgr.m_DataSet.Tables.Remove("GetSiteIndex");
                }
            }
            private void GetSiteIndexSpeciesAlphaCode_access()
            {
                _oAdo.m_strSQL = "SELECT fvs_species FROM " + this.FVSTreeSpeciesTable + " f " +
                                 "WHERE fvs_variant = '" + this.FVSVariant + "' AND " +
                                       "spcd = " + this.SiteIndexSpecies;

                this.SiteIndexSpeciesAlphaCode = _oAdo.getSingleStringValueFromSQLQuery(this._oAdo.m_OleDbConnection, _oAdo.m_strSQL, this.FVSTreeSpeciesTable);
                if (this.SiteIndexSpeciesAlphaCode.Trim().Length == 0)
                    this.SiteIndexSpeciesAlphaCode = "@";



            }
            private void GetSiteIndexSpeciesAlphaCode()
            {
                _oAdo.m_strSQL = "SELECT fvs_species FROM " + this.FVSTreeSpeciesTable + " f " +
                                 "WHERE fvs_variant = '" + this.FVSVariant + "' AND " +
                                       "spcd = " + this.SiteIndexSpecies;

                using (OleDbConnection conn = new OleDbConnection(_oAdo.getMDBConnString(_strTempMDBFileNewProcess, "", "")))
                {
                    conn.Open();

                    this.SiteIndexSpeciesAlphaCode = _oAdo.getSingleStringValueFromSQLQuery(conn, _oAdo.m_strSQL, this.FVSTreeSpeciesTable);
                    
                }
                if (this.SiteIndexSpeciesAlphaCode.Trim().Length == 0)
                    this.SiteIndexSpeciesAlphaCode = "@";
            }
            private void GetSiteIndexSpeciesAlphaCodeNew()
            {
                dataMgr.m_strSQL = "SELECT fvs_species FROM " + FVSTreeSpeciesTable +
                     " WHERE fvs_variant = '" + FVSVariant + "' AND " +
                     "spcd = " + SiteIndexSpecies;
                SiteIndexSpeciesAlphaCode = dataMgr.getSingleStringValueFromSQLQuery(tempConn, dataMgr.m_strSQL, FVSTreeSpeciesTable);
                if (SiteIndexSpeciesAlphaCode.Trim().Length == 0)
                {
                    SiteIndexSpeciesAlphaCode = "@";
                }
            }

            private void LoadSiteIndexValues_access(int p_intSICondId,
                int p_intSISpCd,
                int p_intSIAgeDia,
                int p_intSIHtFt,
                ref int p_intSIFVSSpecies,
                ref double p_dblSiteIndex,
                int p_intSiTree)
            {

                // added base age as parameter for several equations
                // input 0 for that paramter for the old method

                //
                //Western Cascades variant
                //
                if (this.FVSVariant == "WC")
                {

                    if (p_intSISpCd == 11) //pacific silver fir
                    {
                        p_dblSiteIndex = ABAM2(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 11;
                    }
                    else if (p_intSISpCd == 17 || p_intSISpCd == 15)  //grand fir or white fir
                    {
                        p_dblSiteIndex = ABGR1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 15;
                    }
                    else if (p_intSISpCd == 19 || p_intSISpCd == 93) //subalpine fir, englemann spruce
                    {
                        p_dblSiteIndex = PIEN3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 93;
                    }
                    else if (p_intSISpCd == 20 || p_intSISpCd == 21)  //CA red fir, Shasta red fir
                    {
                        p_dblSiteIndex = zABMA2(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 20;
                    }
                    else if (p_intSISpCd == 22) //noble fir
                    {
                        p_dblSiteIndex = ABPR1(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 22;
                    }
                    else if (p_intSISpCd == 42 ||
                        p_intSISpCd == 119 ||
                        p_intSISpCd == 202 ||
                        p_intSISpCd == 242)  //Alaska cedar, western white pine, Douglas-fir, red cedar
                    {
                        p_dblSiteIndex = qPSME13(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 202;
                    }
                    else if (p_intSISpCd == 108) //lodgepole
                    {
                        getAvgDbhAndBasalArea_access(p_intSICondId);
                        p_dblSiteIndex = zPICO3(p_intSIAgeDia,
                            p_intSIHtFt,
                            this.ConditionClassBasalAreaPerAcre,
                            this.ConditionClassAverageDia);
                        p_intSIFVSSpecies = 108;
                    }
                    else if (p_intSISpCd == 119) //western white pine
                    {
                        p_dblSiteIndex = PIMO3(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 119;
                    }
                    else if (p_intSISpCd == 122)  //ponderosa pine
                    {
                        p_dblSiteIndex = PIPO3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 122;
                    }
                    else if (p_intSISpCd == 263) //western hemlock
                    {
                        p_dblSiteIndex = TSHE1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 263;
                    }
                    else if (p_intSISpCd == 264) //mountain hemlock
                    {
                        p_dblSiteIndex = zTSME(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 264;
                    }
                    else if (p_intSISpCd == 351) //red alder
                    {
                        p_dblSiteIndex = ALRU2(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 351;
                    }
                    else if (p_intSISpCd == 72 || p_intSISpCd == 73) //subalpine larch (larix lyallii)
                    {
                        p_dblSiteIndex = LAOC1_OR(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 73;
                    }
                    else if (p_intSISpCd == 211 ||
                        p_intSISpCd == 312 ||
                        p_intSISpCd == 321 ||
                        p_intSISpCd == 475 ||
                        p_intSISpCd == 352 ||
                        p_intSISpCd == 375 ||
                        p_intSISpCd == 431 ||
                        p_intSISpCd == 746 ||
                        p_intSISpCd == 747 ||
                        p_intSISpCd == 815 ||
                        p_intSISpCd == 64 ||
                        p_intSISpCd == 101 ||
                        p_intSISpCd == 103 ||
                        p_intSISpCd == 231 ||
                        p_intSISpCd == 492 ||
                        p_intSISpCd == 500 ||
                        p_intSISpCd == 920)
                    {
                        //redwood, maple, maple, maple, white alder,
                        //paper birch, golden chink, quaking aspen
                        //black cottonwood, white oak, juniper, whitebark pine,
                        //knobcone pine, pacific yew
                        //pacific dogwood, hawthorne, bitter cherry, willow
                        p_dblSiteIndex = qPSME13(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 202;
                    }
                    else
                    {
                        p_dblSiteIndex = 0;
                        p_intSIFVSSpecies = 999;
                    }
                    return;
                }
                //
                //Eastern Cascades variant
                //
                if (this.FVSVariant == "EC")
                {
                    if (p_intSISpCd == 119) //western white pine
                    {
                        p_dblSiteIndex = PIMO2(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 119;
                    }
                    else if (p_intSISpCd == 73)  //western larch
                    {
                        p_dblSiteIndex = LAOC1_OR(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 73;
                    }
                    else if (p_intSISpCd == 202) //Douglas-fir
                    {
                        p_dblSiteIndex = qPSME12(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 202;
                    }
                    else if (p_intSISpCd == 11 ||
                        p_intSISpCd == 17 ||
                        p_intSISpCd == 15)    //Pacific silver fir and grand fir
                    {
                        p_dblSiteIndex = ABGR1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 15;
                    }
                    else if (p_intSISpCd == 108) //lodgepole
                    {
                        getAvgDbhAndBasalArea_access(p_intSICondId);
                        p_dblSiteIndex = zPICO3(p_intSIAgeDia,
                            p_intSIHtFt,
                            this.ConditionClassBasalAreaPerAcre,
                            this.ConditionClassAverageDia);
                        p_intSIFVSSpecies = 108;
                    }
                    else if (p_intSISpCd == 93) //englemann spruce
                    {
                        p_dblSiteIndex = PIEN3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 93;
                    }
                    else if (p_intSISpCd == 19 || p_intSISpCd == 22) //subalpine fir,noble fir
                    {
                        p_dblSiteIndex = zABPR2(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 22;
                    }
                    else if (p_intSISpCd == 122)  //ponderosa pine
                    {
                        p_dblSiteIndex = PIPO3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 122;
                    }
                    else if (p_intSISpCd == 264) //mountain hemlock
                    {
                        p_dblSiteIndex = zTSME(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 264;
                    }

                    else
                    {
                        p_dblSiteIndex = 0;
                        p_intSIFVSSpecies = 999;
                        //	MessageBox.Show("No site index equation found for species " + p_intSISpCd.ToString() + " of variant " + this.FVSVariant);
                    }

                }
                //
                //Pacific Northwest Coast variant
                //
                if (this.FVSVariant == "PN")
                {
                    if (p_intSISpCd == 11) //pacific silver fir
                    {
                        p_dblSiteIndex = ABAM2(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 11;
                    }
                    else if (p_intSISpCd == 17 || p_intSISpCd == 15)  //grand fir or white fir
                    {
                        p_dblSiteIndex = ABGR1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 15;
                    }
                    else if (p_intSISpCd == 19 || p_intSISpCd == 93) //subalpine fir, englemann spruce
                    {
                        p_dblSiteIndex = PIEN3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 93;
                    }
                    else if (p_intSISpCd == 20 || p_intSISpCd == 21)  //CA red fir, Shasta red fir
                    {
                        p_dblSiteIndex = zABMA2(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 20;
                    }
                    else if (p_intSISpCd == 98) //Sitka spruce
                    {
                        p_dblSiteIndex = PISI1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 98;
                    }
                    else if (p_intSISpCd == 22) //noble fir
                    {
                        p_dblSiteIndex = ABPR1(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 22;
                    }
                    else if (p_intSISpCd == 108) //lodgepole
                    {
                        getAvgDbhAndBasalArea_access(p_intSICondId);
                        p_dblSiteIndex = zPICO3(p_intSIAgeDia,
                            p_intSIHtFt,
                            this.ConditionClassBasalAreaPerAcre,
                            this.ConditionClassAverageDia);
                        p_intSIFVSSpecies = 108;
                    }
                    else if (p_intSISpCd == 119) //western white pine
                    {
                        p_dblSiteIndex = PIMO3(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 119;
                    }
                    else if (p_intSISpCd == 122)  //ponderosa pine
                    {
                        p_dblSiteIndex = PIPO3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 122;
                    }
                    else if (p_intSISpCd == 202) //Douglas-fir
                    {
                        p_dblSiteIndex = qPSME1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 202;
                    }
                    else if (p_intSISpCd == 263) //western hemlock
                    {
                        p_dblSiteIndex = TSHE1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 263;
                    }
                    else if (p_intSISpCd == 264) //mountain hemlock
                    {
                        p_dblSiteIndex = zTSME(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 264;
                    }
                    else if (p_intSISpCd == 73 || p_intSISpCd == 72)  //western larch
                    {
                        p_dblSiteIndex = LAOC1_OR(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 73;
                    }
                    else if (p_intSISpCd == 42 ||
                        p_intSISpCd == 73 ||
                        p_intSISpCd == 211 ||
                        p_intSISpCd == 312 ||
                        p_intSISpCd == 321 ||
                        p_intSISpCd == 475 ||
                        p_intSISpCd == 352 ||
                        p_intSISpCd == 375 ||
                        p_intSISpCd == 431 ||
                        p_intSISpCd == 746 ||
                        p_intSISpCd == 747 ||
                        p_intSISpCd == 815 ||
                        p_intSISpCd == 64 ||
                        p_intSISpCd == 101 ||
                        p_intSISpCd == 103 ||
                        p_intSISpCd == 231 ||
                        p_intSISpCd == 492 ||
                        p_intSISpCd == 500 ||
                        p_intSISpCd == 768 ||
                        p_intSISpCd == 920)
                    {
                        //Alaska cedar,western larch, redwood, maple, maple, maple, white alder,
                        //paper birch, golden chink, quaking aspen
                        // black cottonwood, white oak, juniper, whitebark pine, knobcone pine, pacific yew
                        //pacific dogwood, hawthorne, bitter cherry, willow
                        p_dblSiteIndex = qPSME13(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 202;

                    }
                    else
                    {
                        p_dblSiteIndex = 0;
                        p_intSIFVSSpecies = 999;
                    }
                }
                //
                //Blue Mountains variant
                //
                if (this.FVSVariant == "BM")
                {
                    if (p_intSISpCd == 119) //western white pine
                    {
                        p_dblSiteIndex = PIMO2(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 119;
                    }
                    else if (p_intSISpCd == 73)  //western larch
                    {
                        p_dblSiteIndex = LAOC1_OR(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 73;
                    }
                    else if (p_intSISpCd == 202) //Douglas-fir
                    {
                        p_dblSiteIndex = qPSME12(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 202;
                    }
                    else if (p_intSISpCd == 17 ||
                             p_intSISpCd == 15) //grand fir and white fir
                    {
                        p_dblSiteIndex = ABGR1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 17;
                    }
                    else if (p_intSISpCd == 108) //lodgepole
                    {
                        getAvgDbhAndBasalArea_access(p_intSICondId);
                        p_dblSiteIndex = zPICO3(p_intSIAgeDia,
                            p_intSIHtFt,
                            this.ConditionClassBasalAreaPerAcre,
                            this.ConditionClassAverageDia);
                        p_intSIFVSSpecies = 108;
                    }
                    else if (p_intSISpCd == 93) //englemann spruce
                    {
                        p_dblSiteIndex = PIEN3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 93;
                    }
                    else if (p_intSISpCd == 19 || p_intSISpCd == 22) //subalpine fir,noble fir
                    {
                        p_dblSiteIndex = zABPR2(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 22;
                    }
                    else if (p_intSISpCd == 122 || p_intSISpCd == 116)  //ponderosa pine
                    {
                        p_dblSiteIndex = PIPO3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 122;
                    }
                    else if (p_intSISpCd == 264) //mountain hemlock
                    {
                        p_dblSiteIndex = zTSME(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 264;
                    }
                    else
                    {
                        p_dblSiteIndex = 0;
                        p_intSIFVSSpecies = 999;
                    }

                }
                //
                //Klamath Mountains variant
                //
                if (this.FVSVariant == "NC")
                {
                    if (p_intSISpCd == 202 || //Douglas-fir
                        p_intSISpCd == 211 || //other softwoods
                        p_intSISpCd == 98 ||
                        p_intSISpCd == 103 ||
                        p_intSISpCd == 127 ||
                        p_intSISpCd == 201 ||
                        p_intSISpCd == 264 ||
                        p_intSISpCd == 263 || //Douglas-fir
                        p_intSISpCd == 41)
                    {
                        p_dblSiteIndex = qPSME1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 202;
                    }
                    else if (p_intSISpCd == 20 ||
                        p_intSISpCd == 21 ||
                        p_intSISpCd == 15 ||
                        p_intSISpCd == 81 ||
                        p_intSISpCd == 17)  //red firs, white firs, incense cedar
                    {
                        p_dblSiteIndex = zABCO2(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 15;
                    }
                    else if (p_intSISpCd == 361) //Madrone
                    {
                        p_dblSiteIndex = zARME1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 361;
                    }
                    else if (p_intSISpCd == 818) //California black oak
                    {
                        p_dblSiteIndex = zQUKE(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 818;
                    }
                    else if (p_intSISpCd == 631) //tan oak
                    {
                        p_dblSiteIndex = zLIDE1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 631;
                    }
                    else if (p_intSISpCd == 117 ||    //Sugar pine 
                             p_intSISpCd == 122 ||    //Ponderosa pine
                             p_intSISpCd == 116 ||
                             p_intSISpCd == 124 ||
                             p_intSISpCd == 108 ||    //Ponderosa pine
                             p_intSISpCd == 119)      //Sugar pine     	    

                    {
                        p_dblSiteIndex = zPIPO8(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 122;
                    }
                    else
                    {
                        p_dblSiteIndex = 0;
                        p_intSIFVSSpecies = 999;
                    }

                }
                //
                //South Central Oregon / Northeast California variant
                //
                if (this.FVSVariant == "SO")
                {
                    if (p_intSISpCd == 119) //western white pine
                    {
                        p_dblSiteIndex = PIMO2(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 119;
                    }
                    else if (p_intSISpCd == 117)  //sugar pine
                    {
                        p_dblSiteIndex = PIPO3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 122;
                    }
                    else if (p_intSISpCd == 202) //Douglas-fir
                    {
                        p_dblSiteIndex = qPSME12(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 202;
                    }
                    else if (p_intSISpCd == 15)   //white fir
                    {
                        p_dblSiteIndex = ABGR1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 15;
                    }
                    else if (p_intSISpCd == 264) //mountain hemlock
                    {
                        p_dblSiteIndex = zTSME(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 264;
                    }
                    else if (p_intSISpCd == 81)   //incense cedar
                    {
                        p_dblSiteIndex = ABGR1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 15;
                    }
                    else if (p_intSISpCd == 108) //lodgepole
                    {
                        getAvgDbhAndBasalArea_access(p_intSICondId);
                        p_dblSiteIndex = zPICO3(p_intSIAgeDia,
                            p_intSIHtFt,
                            this.ConditionClassBasalAreaPerAcre,
                            this.ConditionClassAverageDia);
                        p_intSIFVSSpecies = 108;
                    }
                    else if (p_intSISpCd == 93) //englemann spruce
                    {
                        p_dblSiteIndex = PIEN3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 93;
                    }
                    else if (p_intSISpCd == 20 || p_intSISpCd == 21) //red fir
                    {
                        p_dblSiteIndex = zABPR2(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 22;
                    }
                    else if (p_intSISpCd == 122 || p_intSISpCd == 116)  //ponderosa pine,jeffrey pine
                    {
                        p_dblSiteIndex = PIPO3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 122;
                    }
                    else if (p_intSISpCd == 64)  //western juniper
                    {
                        getAvgDbhAndBasalArea_access(p_intSICondId);
                        p_dblSiteIndex = SI_LP5(p_intSIAgeDia, p_intSIHtFt,
                            this.ConditionClassBasalAreaPerAcre,
                            this.ConditionClassAverageDia);
                        p_intSIFVSSpecies = 64;
                    }
                    else
                    {
                        p_dblSiteIndex = 0;
                        p_intSIFVSSpecies = 999;
                    }
                }
                //
                //West-side Sierra Nevada
                //
                if (this.FVSVariant == "WS")
                {
                    //Adjustment factors of Dunning by species are from FVS Wessin documentation
                    //Dunning adjustment factors also listed on table 3.4.1.2 in FVS WS Variant Overview from 11/2015
                    if (p_intSISpCd == 119)		//Western white pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 119;
                    }
                    else if (p_intSISpCd == 117) //Sugar pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 117;
                    }
                    else if (p_intSISpCd == 202) //Douglas-fir
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 202;
                    }
                    else if (p_intSISpCd == 15)  //White fir
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 15;
                    }
                    else if (p_intSISpCd == 212)  //Giant sequoia
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 212;
                    }
                    else if (p_intSISpCd == 81)   //Incense cedar
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.76;
                        p_intSIFVSSpecies = 81;
                    }
                    else if (p_intSISpCd == 116)   //Jeffrey Pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 116;
                    }
                    else if (p_intSISpCd == 122)   //Ponderosa Pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 122;
                    }
                    else if (p_intSISpCd == 64)   //Western Juniper
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 64;
                    }
                    else if (p_intSISpCd == 101)   //whitebark pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 101;
                    }
                    else if (p_intSISpCd == 108)   //lodgepole pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 108;
                    }
                    else if (p_intSISpCd == 109)   //coulter pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 109;
                    }
                    else if (p_intSISpCd == 113)   //limber pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 113;
                    }
                    else if (p_intSISpCd == 120)   //bishop pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 120;
                    }
                    else if (p_intSISpCd == 127)   //California foothill pine
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 127;
                    }
                    else if (p_intSISpCd == 201)   //bigcone Douglas-fir
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 201;
                    }
                    else if (p_intSISpCd == 264)   //mountain hemlock
                    {
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt) * 0.9;
                        p_intSIFVSSpecies = 264;
                    }
                    else if (p_intSISpCd == 818)  //Black oak
                    {
                        p_dblSiteIndex = this.qPSME1(p_intSIAgeDia, p_intSIHtFt); //uses King's Douglas-fir
                        p_intSIFVSSpecies = 202;
                    }
                    else if (p_intSISpCd == 20)  //red fir
                    {
                        p_dblSiteIndex = this.zABMA2(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 20;
                    }
                    else if (p_intSISpCd == 631)  //tan oak
                    {
                        p_dblSiteIndex = this.qPSME1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 202;
                    }
                    else if (p_intSISpCd == 103)  //knobcone pine to great basin bristlecone pine
                    {                             //KP maps to GB for this variant according to FVS
                        p_dblSiteIndex = this.PIEN3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 103;
                    }
                    else
                    {
                        p_dblSiteIndex = 0;
                        p_intSIFVSSpecies = 999;
                    }
                }

                if (this.FVSVariant == "CA")
                {
                    //Note: this crosswalk came from a worksheet from CAheight_ref.xls sent
                    //by Chad, FVS-Fort Collins group, 4/30/2002

                    if (p_intSISpCd == 41 || p_intSISpCd == 81 ||
                        p_intSISpCd == 242 || p_intSISpCd == 15 ||
                        p_intSISpCd == 202 || p_intSISpCd == 263 ||
                        p_intSISpCd == 117 || p_intSISpCd == 92 ||
                        p_intSISpCd == 212 || p_intSISpCd == 120 ||
                        p_intSISpCd == 17)
                    {
                        //Port Orford cedar, incense cedar, western redcedar, 
                        //white-fir, Douglas-fir, western hemlock
                        //sugar pine (!), brewer spruce, giant sequoia, bishop pine

                        p_dblSiteIndex = zPSME14(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 202;
                    }
                    else if (p_intSISpCd == 20 || p_intSISpCd == 21 ||
                        p_intSISpCd == 264)  //CA red fir, Shasta red fir,mountain hemlock
                    {
                        p_dblSiteIndex = zABMA2(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 20;
                    }
                    else if (p_intSISpCd == 108 || p_intSISpCd == 101 ||
                        p_intSISpCd == 103 || p_intSISpCd == 109 ||
                        p_intSISpCd == 113 || p_intSISpCd == 64)
                    {
                        //Whitebark pine, knobcone pine, lodgepole pine,
                        //Coulter pine, Limber pine, Western juniper
                        getAvgDbhAndBasalArea_access(p_intSICondId);
                        p_dblSiteIndex = zPICO3(p_intSIAgeDia,
                            p_intSIHtFt,
                            this.ConditionClassBasalAreaPerAcre,
                            this.ConditionClassAverageDia);
                        p_intSIFVSSpecies = 108;
                    }
                    else if (p_intSISpCd == 116 || p_intSISpCd == 119 ||
                        p_intSISpCd == 122 || p_intSISpCd == 124 ||
                        p_intSISpCd == 127)
                    {
                        //Jeffrey pine, western white pine, ponderosa pine,
                        //Monterey pine, gray pine
                        p_dblSiteIndex = zPIPO9(p_intSIAgeDia, p_intSIHtFt, 0);
                        p_intSIFVSSpecies = 122;
                    }
                    else if (p_intSISpCd == 818 || p_intSISpCd == 231 ||
                        p_intSISpCd == 815 || p_intSISpCd == 821 ||
                        p_intSISpCd == 330 || p_intSISpCd == 492 ||
                        p_intSISpCd == 542)
                    {
                        //Pacific yew, Oregon white oak, California black oak, valley white oak,
                        //California buckeye, Pacific dogwood, Oregon ash
                        p_dblSiteIndex = zQUKE(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 818;
                    }
                    else if (p_intSISpCd == 361 || p_intSISpCd == 801 ||
                        p_intSISpCd == 805 || p_intSISpCd == 807 ||
                        p_intSISpCd == 811 || p_intSISpCd == 839 ||
                        p_intSISpCd == 431)
                    {
                        //Coast live oak, canyon live oak, blue oak, engelmann oak,
                        //interior live oak, Pacific madrone, golden chinkapin
                        p_dblSiteIndex = zARME1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 361;
                    }
                    else if (p_intSISpCd == 631) //tan oak
                    {
                        p_dblSiteIndex = zLIDE1(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 631;
                    }
                    else if (p_intSISpCd == 312 || p_intSISpCd == 351 ||
                        p_intSISpCd == 600 || p_intSISpCd == 603 ||
                        p_intSISpCd == 747 || p_intSISpCd == 981)
                    {
                        //Bigleaf maple, red alder, walnuts, black cottonwood, CA laurel
                        p_dblSiteIndex = zALRU3(p_intSIAgeDia, p_intSIHtFt);
                        p_intSIFVSSpecies = 351;
                    }
                    else
                    {
                        p_dblSiteIndex = 0;
                        p_intSIFVSSpecies = 999;
                    }
                }

                // Variants for ID and MT that were implemented using a site
                // index equation database
                if (this.FVSVariant == "CI" || this.FVSVariant == "EM"
                    || this.FVSVariant == "IE" || this.FVSVariant == "TT")
                {
                    // The compound key for the dictionary is the variant + species code
                    string strKey = this.FVSVariant + SI_DELIM + p_intSISpCd;
                    // Initialize values to blank in case the key is not found
                    string strEquation = "";
                    string strSlfSpCd = "";
                    string strRegion = "";
                    if (_dictSiteIdxEq.ContainsKey(strKey))
                    {
                        // If the key is found extract the values from the delimited string
                        string[] arrValues = _dictSiteIdxEq[strKey].Split(Convert.ToChar(SI_DELIM));
                        strEquation = arrValues[0];
                        strSlfSpCd = arrValues[1];
                        strRegion = arrValues[2];
                    }
                    // Reset site index and species code, in case they aren't found in database
                    p_dblSiteIndex = 0;
                    // Return the numeric FVS-FIA species code; 
                    p_intSIFVSSpecies = p_intSISpCd;
                    // Calculate the site index for the equation from the database
                    switch (strEquation)
                    {
                        case "ABGR1":
                            p_dblSiteIndex = ABGR1(p_intSIAgeDia, p_intSIHtFt);
                            break;
                        case "LAOC1_OR":
                            p_dblSiteIndex = LAOC1_OR(p_intSIAgeDia, p_intSIHtFt, 0);
                            break;
                        case "PIEN3":
                            p_dblSiteIndex = PIEN3(p_intSIAgeDia, p_intSIHtFt);
                            break;
                        case "PSME11":
                            p_dblSiteIndex = PSME11(p_intSIAgeDia, p_intSIHtFt);
                            break;
                        case "SI_AS1":
                            p_dblSiteIndex = SI_AS1(p_intSIAgeDia, p_intSIHtFt);
                            // substitute FIA site index if 0 returned; Uses same site index equation
                            if (p_dblSiteIndex == 0)
                            {
                                p_dblSiteIndex = Convert.ToDouble(p_intSiTree);
                            }
                            break;
                        case "SI_DF2":
                            p_dblSiteIndex = SI_DF2(p_intSIAgeDia, p_intSIHtFt, this.ConditionClassHabitatTypeCd);
                            break;
                        case "SI_LP5":
                            getAvgDbhAndBasalArea_access(p_intSICondId);
                            p_dblSiteIndex = SI_LP5(p_intSIAgeDia, p_intSIHtFt, this.ConditionClassBasalAreaPerAcre,
                                this.ConditionClassAverageDia);
                            break;
                        case "SI_PP6":
                            p_dblSiteIndex = SI_PP6(p_intSIAgeDia, p_intSIHtFt);
                            break;
                        default:
                            p_intSIFVSSpecies = 999;
                            break;
                    }
                    // If there is a cross-reference slf species code use it, otherwise use input species code
                    int intSpCd = -1;
                    bool boolSlfSpecies = Int32.TryParse(strSlfSpCd, out intSpCd);
                    if (boolSlfSpecies) p_intSIFVSSpecies = intSpCd;
                }
            }

            private void LoadSiteIndexValues(int p_intSICondId,
                int p_intSISpCd,
                int p_intSIAgeDia,
                int p_intSIHtFt,
                ref int p_intSIFVSSpecies,
                ref double p_dblSiteIndex,
                int p_intSiTree)
            {
                // The compound key for the dictionary is the variant + species code
                string strKey = this.FVSVariant + SI_DELIM + p_intSISpCd;
                // Initialize values to blank in case the key is not found
                string strEquation = "";
                string strSlfSpCd = "";
                string strRegion = "";
                string strBaseAge = "";
                if (_dictSiteIdxEq.ContainsKey(strKey))
                {
                    // If the key is found extract the values from the delimited string
                    string[] arrValues = _dictSiteIdxEq[strKey].Split(Convert.ToChar(SI_DELIM));
                    strEquation = arrValues[0];
                    strSlfSpCd = arrValues[1];
                    strRegion = arrValues[2];
                    strBaseAge = arrValues[3];
                }
                if (_strBaseAge != strBaseAge)
                {
                    _strBaseAge = strBaseAge;
                }
                
                // Reset site index and species code, in case they aren't found in database
                p_dblSiteIndex = 0;

                int intBaseAge = 0;
                if (!strBaseAge.Contains("@") && strBaseAge != "")
                {
                    intBaseAge = Convert.ToInt32(strBaseAge);
                }
                // Return the numeric FVS-FIA species code; 
                p_intSIFVSSpecies = p_intSISpCd;
                // Calculate the site index for the equation from the database
                switch (strEquation)
                {
                    case "ABAM2":
                        p_dblSiteIndex = ABAM2(p_intSIAgeDia, p_intSIHtFt, intBaseAge);
                        break;
                    case "ABGR1":
                        p_dblSiteIndex = ABGR1(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "ABPR1":
                        p_dblSiteIndex = ABPR1(p_intSIAgeDia, p_intSIHtFt, intBaseAge);
                        break;
                    case "ALRU1":
                        p_dblSiteIndex = ALRU1(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "ALRU2":
                        p_dblSiteIndex = ALRU2(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "LAOC1_OR":
                        p_dblSiteIndex = LAOC1_OR(p_intSIAgeDia, p_intSIHtFt, intBaseAge);
                        break;
                    case "LAOC1_WA":
                        p_dblSiteIndex = LAOC1_WA(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "PIEN3":
                        p_dblSiteIndex = PIEN3(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "PIMO2":
                        p_dblSiteIndex = PIMO2(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "PIMO3":
                        p_dblSiteIndex = PIMO3(p_intSIAgeDia, p_intSIHtFt, intBaseAge);
                        break;
                    case "PIPO3":
                        p_dblSiteIndex = PIPO3(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "PISI1":
                        p_dblSiteIndex = PISI1(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "PSME11":
                        p_dblSiteIndex = PSME11(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "SI_AS1":
                        p_dblSiteIndex = SI_AS1(p_intSIAgeDia, p_intSIHtFt);
                        // substitute FIA site index if 0 returned; Uses same site index equation
                        if (p_dblSiteIndex == 0)
                        {
                            p_dblSiteIndex = Convert.ToDouble(p_intSiTree);
                        }
                        break;
                    case "SI_DF2":
                        p_dblSiteIndex = SI_DF2(p_intSIAgeDia, p_intSIHtFt, this.ConditionClassHabitatTypeCd);
                        break;
                    case "SI_LP5":
                        getAvgDbhAndBasalArea(p_intSICondId);
                        p_dblSiteIndex = SI_LP5(p_intSIAgeDia, p_intSIHtFt, this.ConditionClassBasalAreaPerAcre,
                            this.ConditionClassAverageDia);
                        break;
                    case "SI_PP6":
                        p_dblSiteIndex = SI_PP6(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "TSHE1":
                        p_dblSiteIndex = TSHE1(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "qPSME1":
                        p_dblSiteIndex = qPSME1(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "qPSME12":
                        p_dblSiteIndex = qPSME12(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "qPSME13":
                        p_dblSiteIndex = qPSME13(p_intSIAgeDia, p_intSIHtFt, intBaseAge);
                        break;
                    case "zABCO2":
                        p_dblSiteIndex = zABCO2(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "zABMA2":
                        p_dblSiteIndex = zABMA2(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "zABPR2":
                        p_dblSiteIndex = zABPR2(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "zALRU3":
                        p_dblSiteIndex = zALRU3(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "zARME1":
                        p_dblSiteIndex = zARME1(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "zDunning":
                        p_dblSiteIndex = zDunning(p_intSIAgeDia, p_intSIHtFt);
                        if (p_intSISpCd == 119 || p_intSISpCd == 117 ||
                            p_intSISpCd == 64 || p_intSISpCd == 101 ||
                            p_intSISpCd == 108 || p_intSISpCd == 109 ||
                            p_intSISpCd == 113 || p_intSISpCd == 120 ||
                            p_intSISpCd == 127 || p_intSISpCd == 264)
                        {
                            p_dblSiteIndex = p_dblSiteIndex * 0.9;
                        }
                        else if (p_intSISpCd == 81)
                        {
                            p_dblSiteIndex = p_dblSiteIndex * 0.76;
                        }
                        break;
                    case "zLIDE1":
                        p_dblSiteIndex = zLIDE1(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "zPICO3":
                        getAvgDbhAndBasalArea(p_intSICondId);
                        p_dblSiteIndex = zPICO3(p_intSIAgeDia,
                            p_intSIHtFt,
                            this.ConditionClassBasalAreaPerAcre,
                            this.ConditionClassAverageDia);
                        break;
                    case "zPIPO8":
                        p_dblSiteIndex = zPIPO8(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "zPIPO9":
                        p_dblSiteIndex = zPIPO9(p_intSIAgeDia, p_intSIHtFt, intBaseAge);
                        break;
                    case "zPSME14":
                        p_dblSiteIndex = zPSME14(p_intSIAgeDia, p_intSIHtFt, intBaseAge);
                        break;
                    case "zQUKE":
                        p_dblSiteIndex = zQUKE(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    case "zTSME":
                        p_dblSiteIndex = zTSME(p_intSIAgeDia, p_intSIHtFt);
                        break;
                    default:
                        p_intSIFVSSpecies = 999;
                        break;
                }
                // If there is a cross-reference slf species code use it, otherwise use input species code
                int intSpCd = -1;
                bool boolSlfSpecies = Int32.TryParse(strSlfSpCd, out intSpCd);
                if (boolSlfSpecies) p_intSIFVSSpecies = intSpCd;
            }

            /// <summary>
            ///-- SITE INDEX FOR PACIFIC SILVER FIR - Hoyer
            ///-- Height-age and Site Index Curves for Pacific
            ///-- Silver Fir if in the Pacific Northwest
            ///-- Research Paper:  PNW-RP-418 Hoyer and Herman, Sept. 1989
            ///-- Curves for high elevation ABAM in West Cascades
            ///-- Site index at breast high age of 100
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double ABAM2(int p_intSIDiaAge, int p_intSIHtFt, int intBaseAge)
            {
                double dblSI;
                int intHtFt;
                int intAgeDia;

                intAgeDia = p_intSIDiaAge;
                intHtFt = p_intSIHtFt;
                dblSI = intHtFt * Math.Exp((-0.0268797) * (intAgeDia - intBaseAge) / intAgeDia + 0.0046259 * Math.Pow((double)(intAgeDia - intBaseAge), 2) / 100
                    - 0.0015862 * Math.Pow((double)(intAgeDia - intBaseAge), 3) / 10000
                    - 0.0761453 * (intAgeDia - intBaseAge) / (Math.Pow((double)intHtFt, 0.5))
                    + 0.0891105 * (intAgeDia - intBaseAge) / intHtFt);
                return dblSI;
            }
            /// <summary>
            ///-- SITE INDEX FOR GRAND AND WHITE FIR - Cochran
            ///-- Site Index and Height Growth Curves for Managed Stands of
            ///-- White or Grand fir East of the Cascades in Oregon and Washinton
            ///-- Research Paper:  PNW-252,  April 1979
            ///-- Site index at breast high age of 50
            /// tmb note: checked some values against graph in publication
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double ABGR1(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double a, b, logTreeAge;
                double dblSI;
                if (p_intSIDiaAge > 125) logTreeAge = Math.Log(125);
                else logTreeAge = Math.Log(p_intSIDiaAge);

                a = 3.886 - 1.8017 * logTreeAge + 0.2105 * (logTreeAge * logTreeAge)
                    - 0.0000002885 * Math.Pow(logTreeAge, 9) + 1.187E-18
                    * Math.Pow(logTreeAge, 24);

                b = -0.30935 + 1.2383 * logTreeAge + 0.001762 * Math.Pow(logTreeAge, 4)
                    - 5.4E-06 * Math.Pow(logTreeAge, 9) + 2.046E-07
                    * Math.Pow(logTreeAge, 11) - 4.04E-13 * Math.Pow(logTreeAge, 18);

                dblSI = (p_intSIHtFt - 4.5) * Math.Exp(a) - Math.Exp(a) * Math.Exp(b)
                    + 89.43;
                return dblSI;

            }
            /// <summary>
            ///-- SITE INDEX FOR ENGELMANN SPRUCE - Clendenen/Alexander
            ///-- Base-Age Conversion and Site Index Equations for Englemann
            ///-- Spruce Stands in the Central and Southern Rocky Mountians
            ///-- Research Note:  INT-223  1977
            ///-- Clendenen equations for Alexander SI converted from 100 to 50 yr.
            ///-- age = Total tree Age
            ///-- si = Site index at total age of 100
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double PIEN3(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI;
                int intHtFt;
                double dblSI50;
                double a = 0, b = 0;
                int intTotalAge;

                if (p_intSIDiaAge + 15 > 300) intTotalAge = 300;
                else intTotalAge = p_intSIDiaAge + 15;   //Arbitrary adjustment by Tara
                intHtFt = p_intSIHtFt;
                if (intTotalAge < 50)
                {
                    a = 7.3214 - 0.08797 * (Math.Pow((double)(intTotalAge - 20), 1.3));
                    b = 2.2366 - 0.43083 * (Math.Pow((double)(intTotalAge - 20), 0.31));
                }
                //				else if (intTotalAge > 291) dblSI = -999;
                else
                {
                    a = -25.4094 + (0.00001477) * (Math.Pow((double)(300 - intTotalAge), 2.6));
                    b = 0.7121 + (7.4576E-17) * (Math.Pow((double)(300 - intTotalAge), 6.5));
                }
                dblSI50 = a + b * intHtFt;
                dblSI = 1.2764 * dblSI50 + 14.1943;
                return dblSI;
            }
            /// <summary>
            /// calculate red fir - Dolph 1991 PSW-RP-206 "Polymorphic site index curves for red fir 
            /// in California and southern Oregon"
            /// tmb 4/17/02  Note: Dolph uses an equation where SI can not be solved for explicitly.
            /// Hence the brute force method used below.
            /// si = site index at breast high age of 50
            /// Note: this equation as published in Dolph is missing a minus sign before 
            /// the b in the test_ht equation, and has an extra parenthesis.  
            /// Notified Martin Richie at PSW 4/16/02.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double zABMA2(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                double dblTest_SI;
                double b, b50, dblTest_Ht;
                double b1, b2, b3, b4, b5;
                bool bDone;

                b1 = 1.51744;
                b2 = 0.00000141512;
                b3 = -0.0440853;
                b4 = -3049510;
                b5 = 0.000572474;

                dblTest_SI = 10;
                bDone = false;
                while (!bDone)
                {
                    b = p_intSIDiaAge * Math.Exp(p_intSIDiaAge * b3) * b2 * dblTest_SI + Math.Pow((double)(p_intSIDiaAge * Math.Exp(p_intSIDiaAge * b3) * b2), 2) * b4 + b5;
                    b50 = 50 * Math.Exp(50 * b3) * b2 * dblTest_SI + Math.Pow((double)(50 * Math.Exp(50 * b3) * b2), 2) * b4 + b5;
                    dblTest_Ht = ((dblTest_SI - 4.5) * (1 - Math.Exp(-b * Math.Pow((double)p_intSIDiaAge, b1)))) / (1 - Math.Exp(-b50 * Math.Pow((double)50, b1))) + 4.5;
                    if (dblTest_Ht > p_intSIHtFt)
                    {
                        bDone = true;
                        dblSI = dblTest_SI - 0.5;
                    }
                    else
                    {
                        dblTest_SI = dblTest_SI + 0.5;
                    }
                }
                return dblSI;
            }

            /// <summary>
            ///-- CALCULATE NOBLE FIR SITE INDEX - Herman, Curtis and DeMars
            ///-- Height Growth and Site Index Estimates in High Elevation
            ///-- Forests of the Oregon Washington Cascades   (PNW-243)
            ///-- bh_age = Ring count at breast height
            ///-- si = Site index at breast high age of 100
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double ABPR1(int p_intSIDiaAge, int p_intSIHtFt, int intBaseAge)
            {
                double dblSI = 0;
                double a, b, c;
                if (p_intSIDiaAge <= intBaseAge)
                {
                    dblSI = 4.5
                        + 0.2145
                        * (intBaseAge - p_intSIDiaAge)
                        + 0.0089
                        * ((intBaseAge - p_intSIDiaAge) * (intBaseAge - p_intSIDiaAge))
                        + (p_intSIHtFt - 4.5)
                        * (1.0 + 0.00386 * (intBaseAge - p_intSIDiaAge) + 1.2518
                        * Math.Pow((intBaseAge - p_intSIDiaAge), 5) / Math.Pow(10.0, 10));
                }
                else
                {
                    a = -62.755 + 672.55 * Math.Pow((1.0 / p_intSIDiaAge), 0.5);
                    b = 0.9484 + 516.49 * ((1.0 / p_intSIDiaAge) * (1.0 / p_intSIDiaAge));
                    c = -0.00144 + 0.1442 * (1.0 / p_intSIDiaAge);
                    dblSI = a + b * (p_intSIHtFt - 4.5) + c
                        * ((p_intSIHtFt - 4.5) * (p_intSIHtFt - 4.5));
                }
                return dblSI;
            }

            /// <summary>
            /// qPSME13 is taken from Bruce's site index routine.  Source:  Curtis, RO, Herman, FR, DeMars, DJ.
            ///1974.  Height growth and site index for Douglas-fir in high elevation forests of the Oregon-
            /// Washington cascades.  Forest Science 20(4):307-316
            /// si = breast height age at 100 years
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double qPSME13(int p_intSIDiaAge, int p_intSIHtFt, int intBaseAge)
            {
                double dblSI = 0;
                double a, b;
                if (p_intSIDiaAge <= intBaseAge)
                {
                    a = .010006 * ((intBaseAge - p_intSIDiaAge) * (intBaseAge - p_intSIDiaAge));
                    b = 1 + .00549779 * (intBaseAge - p_intSIDiaAge) + (1.46842 * Math.Pow(10, -14))
                        * (Math.Pow(intBaseAge - p_intSIDiaAge, 7));
                }
                else
                {
                    a = 7.66772 * (Math.Exp(-0.95 * Math.Pow(100.0 / (p_intSIDiaAge - intBaseAge), 2)));
                    b = 1.0 - 0.730948 * Math.Pow(Math.Log10((double)p_intSIDiaAge) - 2, 0.8);

                }
                dblSI = a + b * (p_intSIHtFt - 4.5) + 4.5;
                return dblSI;
            }

            ///<summary>Some site index equations require the Canopy Crown Factor (CCF)
            ///which is a function of avg dbh and basal area per plot. This function queries
            ///the TREE table to calculate these two factors for a given condition.
            ///This function needs to be called prior to running an equation that uses CCF
            ///</summary>
            ///<param name="p_intCondId">The id of the condition</param>  
            private void getAvgDbhAndBasalArea_access(int p_intCondId)
            {

                this.ConditionClassAverageDia = 0;
                this.ConditionClassBasalAreaPerAcre = 0;

                string strSQL = "SELECT t.tpacurr, t.dia " +
                "FROM " + this.TreeTable + " t " +
                "WHERE t.biosum_cond_id = '" +
                this.BiosumPlotId + Convert.ToString(p_intCondId).Trim() +
                "' AND t.statuscd=1 AND t.tpacurr IS NOT NULL and t.dia IS NOT NULL AND t.dia >= 1";

                _oAdo.SqlQueryReader(_oAdo.m_OleDbConnection, strSQL);

                //Variables to accumulate the data from the tree table
                double dblCount = 0;
                double dblDia = 0;
                double dblBasalArea = 0;
                if (_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (_oAdo.m_OleDbDataReader.Read())
                    {
                        double tempTpa = Convert.ToDouble(_oAdo.m_OleDbDataReader["tpacurr"]);
                        double tempDia = Convert.ToDouble(_oAdo.m_OleDbDataReader["dia"]);
                        dblCount = dblCount + tempTpa;
                        dblDia = dblDia + tempDia * tempTpa;
                        dblBasalArea = dblBasalArea + (Math.Pow(tempDia, 2) * 0.00545415) * tempTpa;
                    }
                }
                _oAdo.m_OleDbDataReader.Dispose();

                if (dblCount > 0)
                {
                    this.ConditionClassAverageDia = dblDia / dblCount;
                }
                this.ConditionClassBasalAreaPerAcre = dblBasalArea;
            }

            private void getAvgDbhAndBasalArea(int p_intCondId)
            {
                this.ConditionClassAverageDia = 0;
                this.ConditionClassBasalAreaPerAcre = 0;

                //Variables to accumulate the data from the tree table
                double dblCount = 0;
                double dblDia = 0;
                double dblBasalArea = 0;

                using (OleDbConnection conn = new OleDbConnection(_oAdo.getMDBConnString(_strTempMDBFileNewProcess, "", "")))
                {
                    conn.Open();

                    string strSQL = "SELECT t.tpacurr, t.dia " +
                        "FROM " + this.TreeTable + " t " +
                        "WHERE t.biosum_cond_id = '" +
                        this.BiosumPlotId + Convert.ToString(p_intCondId).Trim() +
                        "' AND t.statuscd=1 AND t.tpacurr IS NOT NULL and t.dia IS NOT NULL AND t.dia >= 1";

                    _oAdo.SqlQueryReader(conn, strSQL);

                    if (_oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (_oAdo.m_OleDbDataReader.Read())
                        {
                            double tempTpa = Convert.ToDouble(_oAdo.m_OleDbDataReader["tpacurr"]);
                            double tempDia = Convert.ToDouble(_oAdo.m_OleDbDataReader["dia"]);
                            dblCount = dblCount + tempTpa;
                            dblDia = dblDia + tempDia * tempTpa;
                            dblBasalArea = dblBasalArea + (Math.Pow(tempDia, 2) * 0.00545415) * tempTpa;
                        }
                    }
                    _oAdo.m_OleDbDataReader.Dispose();
                }

                if (dblCount > 0)
                {
                    this.ConditionClassAverageDia = dblDia / dblCount;
                }
                this.ConditionClassBasalAreaPerAcre = dblBasalArea;
            }

            /// <summary>
            ///-- SITE INDEX FOR LOGDGEPOLE PINE - DAHMS
            ///-- Gross Yield of Central Oregon Lodgepole Pine
            ///-- Research Paper: PNW-8  1964
            ///-- bh_age = Ring count at breast height
            ///-- si = Site index at breast high age of 100
            ///Tara's note: This is the same function as PICO2,
            /// but with the addition of a density correction factor from
            /// the Dahm's 1975 publication. If the p_dblBasalArea field or p_dblAvgDbh field
            /// is null or results in a CCF less than 125, no correction factor is applied.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <param name="p_dblBasalArea"></param>
            /// <param name="p_dblAvgDbh"></param>
            /// <returns></returns>
            private double zPICO3(int p_intSIDiaAge, int p_intSIHtFt, double p_dblBasalArea, double p_dblAvgDbh)
            {
                double dblSI;
                int intHtFt;
                int intTtlAge;
                double a, b, CCF;

                intTtlAge = p_intSIDiaAge + 10;
                intHtFt = p_intSIHtFt;
                a = 68.18 - 8.8 * (Math.Pow((double)intTtlAge, 0.45));
                b = 2.2614 - 1.26489 * (Math.Pow((double)(1 - Math.Exp(-0.08333 * intTtlAge)), 5));
                dblSI = a + 4.5 + b * (intHtFt - 4.5);
                if (p_dblBasalArea == 0 || p_dblAvgDbh == 0)
                {
                }
                else
                {
                    CCF = Math.Exp(1.25021 + 0.97834 * Math.Log10(p_dblBasalArea) -
                        0.524548 * Math.Log10(p_dblAvgDbh)); //p.218 of Dahms(1975)
                    if ((CCF > 125) && (p_intSIDiaAge > 41))
                        dblSI = dblSI / (1 - 0.00082 * (CCF - 125)); //corrected for density

                }
                return dblSI;
            }
            /// <summary>
            ///-- SITE INDEX FOR WESTERN WHITE PINE - Curtis
            ///-- PNW-RP-423 (Cascades - Mt Hood and Gifford Pinchot NF's)
            ///-- Height Growth and Site Index for Western White Pine in the
            ///-- Cascade Range of Washington and Oregon
            ///-- IAGE = Ring count at breast height
            ///-- si = Site index at breast high age of 100
            /// tara note: checked against figures in publication
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double PIMO3(int p_intSIDiaAge, int p_intSIHtFt, int intBaseAge)
            {
                double dblSI;
                int intHtFt;
                int intDiaAge;
                double c1;
                double a, b;

                intDiaAge = p_intSIDiaAge;
                intHtFt = p_intSIHtFt;
                c1 = intDiaAge - intBaseAge;
                a = Math.Exp(0.37072 * (Math.Log((double)intDiaAge) - Math.Log((double)100)) - 0.03745 * c1 + (0.00021645) * Math.Pow((double)c1, 2));
                b = (Math.Pow((double)intHtFt, (1 + 0.005936 * c1 - (0.00003879) * Math.Pow((double)c1, 2))));
                dblSI = a * b;
                return dblSI;
            }
            /// <summary>
            ///-- SITE INDEX FOR PONDEROSA PINE - Barrett
            ///-- Height Growth and Site Index Curves for Managed, Even-aged
            ///-- Stands of Ponderosa Pine in the Pacific Northwest
            ///-- Reseach Paper:  PNW-232
            ///-- bh_age = ring count at 4.5 feet
            ///-- si = site index at breast high age of 100
            /// Tara note: checked some points against figure in publication - seems o.k.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double PIPO3(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                double a, b, c;
                if (p_intSIDiaAge <= 130)
                {
                    a = 100.43;
                    b = 1.198632 - (0.00283073 * p_intSIDiaAge) + (8.44441 / p_intSIDiaAge);
                    c = 128.8952205 * (Math.Pow(1.0 - (Math.Exp(-0.016959 * p_intSIDiaAge)),
                        1.23114));
                    dblSI = a - (b * c) + (b * (p_intSIHtFt - 4.5)) + 4.5;
                }
                else
                {
                    dblSI = ((5.328 * (Math.Pow(p_intSIDiaAge, -0.1)) - 2.378) * (p_intSIHtFt - 4.5)) + 4.5;
                }
                return dblSI;
            }
            /// <summary>
            ///-- SITE INDEX FOR WESTERN HEMLOCK - Wiley
            ///-- WEYERHAUSER FORESTRY PAPER NO.17 - PAGE 10, 1978
            ///-- Site Index Tables for Western Hemlock in the Pacific Northwest
            ///-- age = Ring count at breast height
            ///-- si = Site index at breast high age of 50
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double TSHE1(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double denom;
                double dblSI = 0;

                if (p_intSIDiaAge <= 120)
                {
                    denom = ((p_intSIDiaAge * p_intSIDiaAge) - (p_intSIHtFt - 4.5)
                        * (-1.7307 - 0.0616 * p_intSIDiaAge + 0.00192 * (p_intSIDiaAge * p_intSIDiaAge)));
                    if (denom != 0)
                    {
                        dblSI = 2500 * (((p_intSIHtFt - 4.5) * (0.1394 + 0.0137 * p_intSIDiaAge + 0.00007 * (p_intSIDiaAge * p_intSIDiaAge))) / denom) + 4.5;
                    }

                }
                else
                {
                    dblSI = 4.5 + 22.6 * Math.Exp((0.014482 - 0.001162 * Math.Log(p_intSIDiaAge))
                        * (p_intSIHtFt - 4.5));
                }
                return dblSI;
            }
            /// <summary>
            ///calculate Means mountain hemlock site index.  From "Means, Joseph E.,
            /// Mary H. Campbell, and Gregory P. Johnson.  1988.  FIR report Vol. 10, No. 1,
            /// p8-9."  Also unpublished draft manuscript - has more detail.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double zTSME(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI;
                double a, b, c, d, e, f, g;
                double dblMetricHeight;
                a = 17.22;
                b = 0.583224;
                c = 99.1275;
                d = -1.189892;
                e = 47.9256;
                f = -0.00574787;
                g = 1.241599;

                dblMetricHeight = p_intSIHtFt * 0.3048; //Convert from feet to metric
                dblSI = 1.37 + a + (b + c * Math.Pow(p_intSIDiaAge, d)) * (dblMetricHeight - 1.37 - e * Math.Pow((double)(1 - Math.Exp(f * p_intSIDiaAge)), g));

                return dblSI;
            }
            /// <summary>
            ///-- CALCULATE RED ALDER - Worthington
            ///-- (PNW-36) - Normal Yeild Tables for Red Alder - 1960
            ///-- bh_age = Ring count at breast height
            ///-- si = Site index at breast high age of 50
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double ALRU2(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = (0.60924 + 19.538 / p_intSIDiaAge) * p_intSIHtFt;
                return dblSI;
            }

            /// <summary>
            /// - SITE INDEX FOR WESTERN LARCH - COCHRAN
            ///-- Site Index, Height Growth, Normal Yields, and Stocking
            ///-- Levels for Larch in Oregon and Washinton
            ///-- Research Note: PNW-424, May 1985
            ///-- bh_age = Ring count at breast height
            ///-- tht = Total tree height
            ///-- si = Site index at breast high age of 50
            /// </summary>
            /// <param name="p_intSIDiaAge">ring count at breast height</param>
            /// <param name="p_intSIHtFt">total tree height in feet</param>
            /// <returns></returns>
            private double LAOC1_OR(int p_intSIDiaAge, int p_intSIHtFt, int intBaseAge)
            {
                int intSIDiaAge = 0;
                if (p_intSIDiaAge > (intBaseAge * 2)) intSIDiaAge = (intBaseAge * 2);
                else intSIDiaAge = p_intSIDiaAge;

                double dblSI = 78.07
                    + (p_intSIHtFt - 4.5)
                    * (3.51412 - 0.125483 * intSIDiaAge + 0.0023559 * Math.Pow(intSIDiaAge, 2)
                    - 0.00002028 * Math.Pow(intSIDiaAge, 3) + 0.000000064782 * Math.Pow(
                    intSIDiaAge, 4))
                    - (3.51412 - 0.125483 * intSIDiaAge + 0.0023559 * Math.Pow(intSIDiaAge, 2)
                    - 0.00002028 * Math.Pow(intSIDiaAge, 3) + 0.000000064782 * Math.Pow(
                    intSIDiaAge, 4))
                    * (1.46897 * intSIDiaAge + 0.0092466 * Math.Pow(intSIDiaAge, 2) - 0.00023957
                    * Math.Pow(intSIDiaAge, 3) + 0.0000011122 * Math.Pow(intSIDiaAge, 4));
                return dblSI;
            }
            /// <summary>
            /// western larch in western washington and eastern washington;
            /// and western white pine in eastern washington (Brickell 1970)
            /// This equation replaced Cochran (1985) equation for the same areas and species
            /// beginning in 1990
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double LAOC1_WA(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0.37956 * p_intSIHtFt * Math.Exp(48.4372 / (p_intSIDiaAge + 8));
                return dblSI;
            }
            /// <summary>
            ///-- SITE INDEX FOR WESTERN WHITE PINE - Brickell and Haig
            ///-- Equations and computer subroutines for Estimating Site
            ///-- Quality of Eight Rocky Mountian Species
            ///-- Reaserch Paper: INT-75  1970
            ///-- age = Age of dominant trees in even-aged stand
            ///-- tht = Average height of dominant and codominant trees
            ///-- si = Site index at base age of 50
            /// Tara notes: Apparently developed from Haig, Irvine T. 1932.  Second-growth yield, stand and volume
            /// tables for the western white pine type.  USDA Tech. Bull. 323.
            /// Probably total age?  Note non-standard selection of "height" (average) - this function operates on
            /// individual trees.
            /// Changed coefficients to match Brickell publication
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double PIMO2(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI;
                int intHtFt;
                int intDiaAge;
                intDiaAge = p_intSIDiaAge + 11; //Average reported in Haig 1932.
                intHtFt = p_intSIHtFt;
                dblSI = 0.37504453 * intHtFt * (Math.Pow((double)(1 - 0.92503 * Math.Exp(-0.0207959 * p_intSIDiaAge)), -2.4881068));
                return dblSI;
            }
            /// <summary>
            ///calculate Cochran Douglas-fir index.  From "Cochran, PH. 1979.
            ///Site index and height growth curves for
            ///managed, even-aged stands of Douglas-fir east of the Cascades in Oregon and Washington.
            ///RP-PNW-251."
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double qPSME12(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI;
                double a, b, logTreeAge;

                if (p_intSIDiaAge > 100) logTreeAge = Math.Log(100);
                else logTreeAge = Math.Log(p_intSIDiaAge);
                a = Math.Exp(-0.37496 + 1.36164 * logTreeAge - 0.00243434
                    * Math.Pow((double)logTreeAge, 4));
                b = 0.52032 - 0.0013194 * p_intSIDiaAge + 27.2823 / p_intSIDiaAge;
                dblSI = 84.47 - a * b + (p_intSIHtFt - 4.5) * b;
                return dblSI;

            }

            /// <summary>
            ///-- Site Index Noble fir - DeMars, D.J., Herman, F.R., and J.F. Bell.  1970.  
            ///Preliminary site index curves for noble fir from stem analysis data.
            ///-- PNW Research Note PNW-119 May 1970.  si = 100 years at breast height age, total height
            ///-- Based on tallest dominant
            ///-- Tara's note: the original publication contained graphs and a table,
            /// developed from 4 polymorphic curves
            ///-- This program is my adaption of the method used.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double zABPR2(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                double ht_A, ht_B, ht_C, ht_D;
                double propDiff;
                int intAgeDia;
                intAgeDia = p_intSIDiaAge;

                ht_A = ((intAgeDia * intAgeDia) / (13.64781 + (0.19937 * intAgeDia) + (0.00416 * (intAgeDia * intAgeDia)))) + 4.5;
                ht_B = ((intAgeDia * intAgeDia) / (10.11508 + (0.40115 * intAgeDia) + (0.00426 * (intAgeDia * intAgeDia)))) + 4.5;
                ht_C = ((intAgeDia * intAgeDia) / (19.10187 + (0.50996 * intAgeDia) + (0.0049 * (intAgeDia * intAgeDia)))) + 4.5;
                ht_D = ((intAgeDia * intAgeDia) / (12.68825 + (1.05408 * intAgeDia) + (0.00501 * (intAgeDia * intAgeDia)))) + 4.5;

                if (p_intSIHtFt >= ht_A)
                    dblSI = (p_intSIHtFt / ht_A) * 137.5056;  //137 is the SI for polymorphic curve A

                if ((p_intSIHtFt >= ht_B) && (p_intSIHtFt < ht_A))
                {
                    propDiff = (p_intSIHtFt - ht_B) / (ht_A - ht_B);
                    dblSI = (propDiff * (137.5056 - 112.2237)) + 112.2237; //112 is the SI for polymorphic curve B
                }
                if ((p_intSIHtFt >= ht_C) && (p_intSIHtFt < ht_B))
                {
                    propDiff = (p_intSIHtFt - ht_C) / (ht_B - ht_C);
                    dblSI = (propDiff * (112.2237 - 88.46456)) + 88.46456;  //88 is the SI for polymorphic curve B
                }
                if ((p_intSIHtFt >= ht_D) && (p_intSIHtFt < ht_C))
                {
                    propDiff = (p_intSIHtFt - ht_D) / (ht_C - ht_D);
                    dblSI = (propDiff * (88.46456 - 63.95436)) + 63.95436;  //63 is the SI for polymorphic curve B
                }
                if (p_intSIHtFt < ht_D)
                    dblSI = (p_intSIHtFt / ht_D) * 63.95436;

                return dblSI;
            }
            /// <summary>
            ///-- SITKA SPRUCE SITE INDEX - Farr
            ///-- (PNW-326) Site Index and Height Growth Curves for Unmanaged Even-
            ///-- Aged Stannds of Western Hemlock and Sitka Spruce In Southeast Alaska
            ///-- Research Paper: PNW-326,  OCT. 1984
            ///-- intAgeDia = Ring count at breast height
            ///-- si = Site index at breast high age of 50
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double PISI1(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI;
                int intHtFt, intDiaAge;
                double c1, c2;

                intDiaAge = p_intSIDiaAge;
                intHtFt = p_intSIHtFt;
                c1 = 6.396816 - 4.098921 * Math.Log(intDiaAge) + 0.76287 * (Math.Pow((double)(Math.Log(intDiaAge)), 2))
                    - (0.00244688) * (Math.Pow((double)(Math.Log(intDiaAge)), 5))
                    + (0.000000244581) * (Math.Pow((double)(Math.Log(intDiaAge)), 10))
                    - (2.02215E-22) * (Math.Pow((double)(Math.Log(intDiaAge)), 30));
                c2 = -0.20505 + 1.4496 * Math.Log(intDiaAge) - 0.01781 * (Math.Pow((double)(Math.Log(intDiaAge)), 3))
                    + (0.0000651975) * (Math.Pow((double)(Math.Log(intDiaAge)), 5))
                    - (1.09559E-23) * (Math.Pow((double)(Math.Log(intDiaAge)), 30));
                dblSI = 90.93 - Math.Exp(c1) * Math.Exp(c2) + Math.Exp(c1) * (intHtFt - 4.5);
                return dblSI;
            }
            /// <summary>
            ///-- CALCULATE DOUGLAS FIR SITE INDEX - King
            ///-- Site Index for Douglas Fir in the Pacific Northwest
            ///-- WEYERHAUSER FORESTRY PAPER NO. 8,  July 1966
            ///-- site for breast height age 50
            ///-- Taras note:  This equation is listed incorrectly in Ericas publication PNW-RN-533.  However,
            ///-- the version that appears here looks correct.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double qPSME1(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                double z = 0, denom;
                denom = (0.109757 + (0.00792236 * p_intSIDiaAge) + 0.000197693 * (p_intSIDiaAge * p_intSIDiaAge));
                if (denom != 0)
                {
                    z = ((p_intSIDiaAge * p_intSIDiaAge) / (p_intSIHtFt - 4.5) + 0.954038
                        - (0.0558178 * p_intSIDiaAge) + 0.000733819 * (p_intSIDiaAge * p_intSIDiaAge))
                        / denom;

                }
                if (z != 0)
                {
                    dblSI = (2500 / z) + 4.5;
                }
                return dblSI;
            }
            /// <summary>
            ///-- SITE INDEX RED ALDER - Harrington
            ///-- Height Growth and Site Index Curves for Red Alder
            ///-- Research Paper:  PNW-326 1986
            ///-- bh_age = Ring count at breast height
            ///-- si = Site index at breast high age of 20
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double ALRU1(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI;
                int intHtFt, intDiaAge;
                intDiaAge = p_intSIDiaAge;
                intHtFt = p_intSIHtFt;
                dblSI = 54.185 - 4.61694 * intDiaAge + 0.11065 * intDiaAge * intDiaAge
                    - 0.0007633 * intDiaAge * intDiaAge * intDiaAge
                    + (1.25934 - 0.012989 * intDiaAge + 3.522 * (Math.Pow((double)(1 / intDiaAge), 3))) * intHtFt;

                return dblSI;
            }
            /// <summary>
            /// This function was taken from Dolph, Leroy K.  1987.  Site index curves for
            //  young-growth California white fir on the western slope of the Sierra Nevada.
            //  Res. Paper PSW-185.  Berkeley, CA: PSW Forest and Range Experiment Station.
            // uses breast height age and total tree height
            //  Checked the function against tables in the publication - o.k.
            //  si = total height at 50 years bhage
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double zABCO2(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI;
                dblSI = 69.91 - (38.020235 * (Math.Pow((double)p_intSIDiaAge, -1.052133)) *
                    (Math.Exp(p_intSIDiaAge * 0.009557)) *
                    101.842894 * (1 - Math.Exp(-0.001442 * (Math.Pow((double)p_intSIDiaAge, 1.679259))))) +
                    (p_intSIHtFt - 4.5) * (38.020235 * (Math.Pow((double)p_intSIDiaAge, -1.052133)) *
                    (Math.Exp(0.0095557 * p_intSIDiaAge)));
                return dblSI;
            }
            /// <summary>
            ///This equation for Madrone is from "Porter, D.R. and Wiant, H.V. 1965. 
            ///  Site index equations for tanoak, Pacific
            ///madrone, and red alder in the redwod region of humboldt county, California.  Journal
            ///of Forestry, April, 1965,p.286-287."  SI is total age at 50 years. Correction from bhage
            /// to total age is based on average values from the publication.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double zARME1(int p_intSIDiaAge, int p_intSIHtFt)
            {

                double dblSI;
                double Total_Age;

                Total_Age = p_intSIDiaAge + 2.8; //Average value from study
                dblSI = (0.375 + (31.233 / Total_Age)) * p_intSIHtFt;
                return dblSI;
            }
            /// <summary>
            ///Taken from "Powers, Robert F.  1972.  Site index curves for unmanaged stands of
            /// California black oak.  Research Note PSW-262.  5 p.
            /// Uses total height, breast height age, tree selection "Dominants", development method
            /// no sectioning: stratification into 5 ponderosa pine site classes, separate fitting for each
            /// resulting in 4 curves listed below.  Interpolation is my addition.  One curve discarded
            /// (Ponderosa pine SI group 110-115).  Equation given in publication doesn't match graph, appears
            /// to be incorrect (?).  Don't use for age less than 40 (curves cross).  
            /// Assumes anamorphic curves for
            /// everything that is not between si 42 and si 56.  Resulting function is slightly different than
            /// what is shown in publication figure (5.5 ft' higher at 100 yrs for SI 70, 2.5' lower
            /// at 100 years for SI 30)
            /// si = total height at breast height age of 50.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double zQUKE(int p_intSIDiaAge, int p_intSIHtFt)
            {

                double dblSI = 0;
                double htPPSI_65, htPPSI_75, htPPSI_85, htPPSI_95;
                double propDiff;

                htPPSI_65 = (Math.Sqrt(p_intSIDiaAge) - 0.442) / 0.158;  // si at 50 = 41.956
                htPPSI_75 = (Math.Sqrt(p_intSIDiaAge) - 1.888) / 0.117;  // si at 50 = 44.300
                htPPSI_85 = (Math.Sqrt(p_intSIDiaAge) - 2.267) / 0.093;  // si at 50 = 51.657
                htPPSI_95 = (Math.Sqrt(p_intSIDiaAge) - 2.08) / 0.088;   // si at 50 = 56.717

                if (p_intSIHtFt >= htPPSI_95)
                    dblSI = (p_intSIHtFt / htPPSI_95) * 56.717;

                if ((p_intSIHtFt < htPPSI_95) && (p_intSIHtFt >= htPPSI_85))
                {
                    propDiff = (p_intSIHtFt - htPPSI_85) / (htPPSI_95 - htPPSI_85);
                    dblSI = propDiff * (56.717 - 51.657) + 51.657;
                }
                if ((p_intSIHtFt < htPPSI_85) && (p_intSIHtFt >= htPPSI_75))
                {
                    propDiff = (p_intSIHtFt - htPPSI_75) / (htPPSI_85 - htPPSI_75);
                    dblSI = propDiff * (51.657 - 44.3) + 44.3;
                }
                if ((p_intSIHtFt < htPPSI_75) && (p_intSIHtFt >= htPPSI_65))
                {
                    propDiff = (p_intSIHtFt - htPPSI_65) / (htPPSI_75 - htPPSI_65);
                    dblSI = propDiff * (44.3 - 41.956) + 41.956;
                }
                if (p_intSIHtFt < htPPSI_65)
                    dblSI = (p_intSIHtFt / htPPSI_65) * 41.956;

                return dblSI;

            }
            /// <summary>
            ///This equation for Tanoak is from "Porter, D.R. and Wiant, H.V. 1965. 
            ///Site index equations for tanoak, Pacific
            ///madrone, and red alder in the redwod region of humboldt county, California.  Journal
            ///of Forestry, April, 1965,p.286-287."  SI is total age at 50 years. Correction from bhage
            /// to total age is based on average values from the publication.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double zLIDE1(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI;
                double Total_Age;

                Total_Age = p_intSIDiaAge + 3.2; //Average value from study
                dblSI = (0.204 + (39.787 / Total_Age)) * p_intSIHtFt;
                return dblSI;
            }
            /// <summary>
            ///From Powers, Robert F. and William W. Oliver.  1978.  Site classification of 
            ///ponderosa pine stands
            ///stocking control in California. USDA Forest Service research paper PSW-128.
            ///Powers & Oliver used total age - added 9 years to make the adjustment from bhage.
            ///si = total height at 50 years total age
            ///Tested function against graphed values.  Do not use for total age less than 10 years 
            ///or greater than 80 yrs.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double zPIPO8(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                double test_Ht, test_SI;
                bool bDone;
                int Total_Age;

                Total_Age = p_intSIDiaAge + 9; //My arbitrary choice to adjust age
                test_SI = 10;
                bDone = false;
                while (!bDone)
                {

                    test_Ht = ((1.88 * test_SI) - 7.178) *
                        (Math.Pow((double)(1 - Math.Exp(-0.025 * Total_Age)), (0.001 * test_SI + 1.64)));
                    if (test_Ht > p_intSIHtFt)
                    {
                        dblSI = test_SI - 0.5;
                        bDone = true;
                    }
                    else
                        test_SI = test_SI + 0.5;

                }
                return dblSI;
            }
            /// <summary>
            ///Tara note:  The following function was taken from the Region 5 "FIA User's guide" for the
            ///1990s inventory on FS land.  It is described as an adaption of "Dunning, Duncan. 1942 (amended
            ///1958).  Forest Research note 28, California Forest and Range Experiment Station."
            ///Per Mike Landram, this adaption was done by Jack Levitan in the 1970s and has been historically
            ///used by R5 since then.  For the Wessin variant development, they added 5 years to adjust for
            /// Dunning's use of stump age.  I chose to include this 5 year adaption in the function below.  The
            /// results are different than the original Dunning.  Other slight alteration from the
            /// R5 code:  SI values are calculated using interpolation, rather than assigning a site class 
            /// to each tree.
            /// This function assumes zage is breast height age
            /// si is total height at 50 years bhage.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double zDunning(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double Ht_Site_0, Ht_Site_1, Ht_Site_2, Ht_Site_3, Ht_Site_4, Ht_Site_5;
                double dblSI = 0;
                double totAge;
                double propDiff;

                totAge = p_intSIDiaAge + 5;

                Ht_Site_0 = -88.9 + 49.7067 * Math.Log(totAge);  //si @ 50 is 110.3
                Ht_Site_1 = -82.2 + 44.1147 * Math.Log(totAge);  //si @ 50 is 94.6
                Ht_Site_2 = -78.3 + 39.1441 * Math.Log(totAge);  //si @ 50 is 78.6
                Ht_Site_3 = -82.1 + 35.416 * Math.Log(totAge);   //si @ 50 is 59.8
                Ht_Site_4 = -56 + 26.7173 * Math.Log(totAge);    //si @ 50 is 51.1
                Ht_Site_5 = -33.8 + 18.64 * Math.Log(totAge);    //si @ 50 is 40.9

                if (p_intSIHtFt >= Ht_Site_0)
                    dblSI = (p_intSIHtFt / Ht_Site_0) * 110.3;

                if ((p_intSIHtFt < Ht_Site_0) && (p_intSIHtFt >= Ht_Site_1))
                {
                    propDiff = (p_intSIHtFt - Ht_Site_1) / (Ht_Site_0 - Ht_Site_1);
                    dblSI = propDiff * (110.3 - 94.6) + 94.6;
                }
                if ((p_intSIHtFt < Ht_Site_1) && (p_intSIHtFt >= Ht_Site_2))
                {
                    propDiff = (p_intSIHtFt - Ht_Site_2) / (Ht_Site_1 - Ht_Site_2);
                    dblSI = propDiff * (94.6 - 78.6) + 78.6;
                }
                if ((p_intSIHtFt < Ht_Site_2) && (p_intSIHtFt >= Ht_Site_3))
                {
                    propDiff = (p_intSIHtFt - Ht_Site_3) / (Ht_Site_2 - Ht_Site_3);
                    dblSI = propDiff * (78.6 - 59.8) + 59.8;
                }
                if ((p_intSIHtFt < Ht_Site_3) && (p_intSIHtFt >= Ht_Site_4))
                {
                    propDiff = (p_intSIHtFt - Ht_Site_4) / (Ht_Site_3 - Ht_Site_4);
                    dblSI = propDiff * (59.8 - 51.1) + 51.1;
                }
                if ((p_intSIHtFt < Ht_Site_4) && (p_intSIHtFt >= Ht_Site_5))
                {
                    propDiff = (p_intSIHtFt - Ht_Site_5) / (Ht_Site_4 - Ht_Site_5);
                    dblSI = propDiff * (51.1 - 40.9) + 40.9;
                }
                if (p_intSIHtFt < Ht_Site_5)
                {
                    dblSI = (p_intSIHtFt / Ht_Site_5) * 40.9;
                }
                return dblSI;
            }
            /// <summary>
            ///From Hann, David W. and John A. Scrivani.  1987. Dominant-height-growth and site-index
            /// equations for Douglas-fir and ponderosa pine in Southwest Oregon.  Forest research
            /// lab, OSU, Research bulletin 59, February 1987.
            /// Program is just a duplicate of their Appendix 2.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double zPSME14(int p_intSIDiaAge, int p_intSIHtFt, int intBaseAge)
            {
                double dblSI;
                double temp;
                double a0, a1, a2, b1, b2, b3, b4;
                double theTest, theB;
                a0 = -6.21693;
                a1 = 0.281176;
                a2 = 1.14354;
                b1 = -0.0521778;
                b2 = 0.000715141;
                b3 = 0.00797252;
                b4 = -0.000133377;

                temp = b1 * (p_intSIDiaAge - intBaseAge) + b2 *
                    Math.Pow((double)(p_intSIDiaAge - intBaseAge), 2) + b3 * Math.Log(p_intSIHtFt - 4.5) *
                    (p_intSIDiaAge - intBaseAge) + b4 * Math.Log(p_intSIHtFt - 4.5) * Math.Pow((double)(p_intSIDiaAge - intBaseAge), 2);
                dblSI = 4.5 + (p_intSIHtFt - 4.5) * Math.Exp(temp);
                theTest = 0;
                while (theTest < 0.999)
                {
                    temp = dblSI;
                    theB = 1 - Math.Exp(-Math.Exp(a0 + a1 * Math.Log(dblSI - 4.5) + a2 * 3.912023));
                    theB = theB / (1 - Math.Exp(-Math.Exp(a0 + a1 * Math.Log(dblSI - 4.5) + a2 * Math.Log(p_intSIDiaAge))));
                    dblSI = 4.5 + (p_intSIHtFt - 4.5) * theB;
                    theTest = 1 - Math.Abs((dblSI - temp) / temp);
                }
                return dblSI;
            }
            /// <summary>
            ///From Hann, David W. and John A. Scrivani.  1987. Dominant-height-growth and site-index
            /// equations for Douglas-fir and ponderosa pine in Southwest Oregon.  Forest research
            /// lab, OSU, Research bulletin 59, February 1987.
            /// Program is just a duplicate of their Appendix 2.
            /// Note:  the publication is missing a "-EXP" in the Appendix 2 program for PP.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double zPIPO9(int p_intSIDiaAge, int p_intSIHtFt, int intBaseAge)
            {
                double dblSI;
                double temp;
                double a0, a1, a2, b1, b2, b3, b4;
                double theTest, theB;
                a0 = -6.54707;
                a1 = 0.288169;
                a2 = 1.21297;
                b1 = -0.069934;
                b2 = 0.000359644;
                b3 = 0.0120483;
                b4 = -0.0000718058;

                temp = b1 * (p_intSIDiaAge - intBaseAge) + b2 *
                    Math.Pow((double)(p_intSIDiaAge - intBaseAge), 2) + b3 *
                    Math.Log(p_intSIHtFt - 4.5) *
                    (p_intSIDiaAge - intBaseAge) + b4 *
                    Math.Log(p_intSIHtFt - 4.5) *
                    Math.Pow((double)(p_intSIDiaAge - intBaseAge), 2);

                dblSI = 4.5 + (p_intSIHtFt - 4.5) * Math.Exp(temp);
                theTest = 0;
                while (theTest < 0.999)
                {
                    temp = dblSI;
                    theB = 1 - Math.Exp(-Math.Exp(a0 + a1 * Math.Log(dblSI - 4.5) + a2 * 3.912023));
                    theB = theB / (1 - Math.Exp(-Math.Exp(a0 + a1 * Math.Log(dblSI - 4.5) + a2 * Math.Log(p_intSIDiaAge))));
                    dblSI = 4.5 + (p_intSIHtFt - 4.5) * theB;
                    theTest = 1 - Math.Abs((dblSI - temp) / temp);
                }

                return dblSI;
            }
            /// <summary>
            /// This equation for red alder is from "Porter, D.R. and Wiant, H.V. 1965. 
            /// Site index equations for tanoak, Pacific
            ///madrone, and red alder in the redwod region of humboldt county, California.  Journal
            ///of Forestry, April, 1965,p.286-287."  SI is total age at 50 years. Correction from bhage
            /// to total age is based on average values from the publication.
            /// </summary>
            /// <param name="p_intSIDiaAge"></param>
            /// <param name="p_intSIHtFt"></param>
            /// <returns></returns>
            private double zALRU3(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI;
                double Total_Age;
                Total_Age = p_intSIDiaAge + 1.2;   //Average value from study
                dblSI = (0.649 + (17.556 / Total_Age)) * p_intSIHtFt;
                return dblSI;
            }

            /// <summary>
            /// This equation for Douglas Fir is from Brickell and Haig
            ///-- Equations and computer subroutines for Estimating Site
            ///-- Quality of Eight Rocky Mountain Species
            ///-- Research Paper: INT-75  1970. p_intSIDiaAge is total age
            ///---Derived from Kurt's PL/SQL
            ///---Curves validated by jsfried
            ///<param name="p_intSIDiaAge"></param>
            ///<param name="p_intSIHtFt"></param>
            /// </summary>
            private double PSME11(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                double b;
                int Total_Age = p_intSIDiaAge;
                if (p_intSIDiaAge < 20)
                {
                    return dblSI;
                }
                else if (p_intSIDiaAge > 200)
                {
                    Total_Age = 200;
                }
                else
                {
                    Total_Age = p_intSIDiaAge;
                }

                b = 40.984664 * (Math.Pow(Total_Age, -0.5) - Math.Pow(50, -0.5))
                    + 4521.1527 * (Math.Pow(Total_Age, -2.5) - Math.Pow(50, -2.5))
                    + 123059.38 * (Math.Pow(Total_Age, -3.5) - Math.Pow(50, -3.5))
                    - 0.53328682E-08 * (Math.Pow(Total_Age, 4) - Math.Pow(50, 4))
                    + 0.37808033E-10 * (Math.Pow(Total_Age, 5) - Math.Pow(50, 5))
                    + 216.64152 * p_intSIHtFt * (Math.Pow(Total_Age, -1.5) - Math.Pow(50, -1.5))
                    - 158121.49 * p_intSIHtFt * (Math.Pow(Total_Age, -4) - Math.Pow(50, -4))
                    + 1894030.8 * p_intSIHtFt * (Math.Pow(Total_Age, -5) - Math.Pow(50, -5))
                    - 0.10230592E-09 * p_intSIHtFt * (Math.Pow(Total_Age, 4) - Math.Pow(50, 4))
                    - 6.0686119 * p_intSIHtFt * p_intSIHtFt * (Math.Pow(Total_Age, -2) - Math.Pow(50, -2))
                    - 25351.090 * p_intSIHtFt * p_intSIHtFt * (Math.Pow(Total_Age, -5) - Math.Pow(50, -5))
                    + 0.33512858E-04 * p_intSIHtFt * p_intSIHtFt * (Total_Age - 50)
                    + 0.17024711E-02 * Math.Pow(p_intSIHtFt, 3) * (Math.Pow(Total_Age, -1) - Math.Pow(50, -1))
                    + 398.36720 * Math.Pow(p_intSIHtFt, 3) * (Math.Pow(Total_Age, -5) - Math.Pow(50, -5))
                    - 0.88665409E-08 * Math.Pow(p_intSIHtFt, 3) * (Math.Pow(Total_Age, 1.5) - Math.Pow(50, 1.5))
                    + 0.40019102E-14 * Math.Pow(p_intSIHtFt, 3) * (Math.Pow(Total_Age, 4) - Math.Pow(50, 4))
                    - 0.46929245E-08 * Math.Pow(p_intSIHtFt, 5) * (Math.Pow(Total_Age, -0.5) - Math.Pow(50, -0.5))
                    - 0.16640659E-20 * Math.Pow(p_intSIHtFt, 5) * (Math.Pow(Total_Age, 4.5) - Math.Pow(50, 4.5));
                dblSI = p_intSIHtFt + b;
                if (dblSI < 10)
                {
                    dblSI = 10;
                }
                else if (dblSI > 110)
                {
                    dblSI = 110;
                }
                return dblSI;
            }

            /// <summary>
            /// SITE INDEX FOR LODGEPOLE PINE - Alexander, R. R., Trackle, D. and Dahms, W. G. 1967
            /// Site indexes for lodgepole pine with corrections for stand density: methodology
            /// USDA Forest Service, Res. Pap. RM-29
            /// Site index at total age of 100
            /// Derived from VBA source code by Don Vandendriese for FIA2FVS from RMRS
            /// Curves validated by jsfried
            /// </summary>
            /// <param name="p_intSIDiaAge">Age of site tree (Ring count at breast height)</param>
            /// <param name="p_intSIHtFt">Diameter of site tree</param>
            /// <param name="p_dblAvgDbh">Average DBH per condition</param>
            /// <param name="p_dblBasalArea">Basal area per acre for the condition</param>
            /// <returns>Site Index at breast high age of 100</returns>
            private double SI_LP5(int p_intSIDiaAge, int p_intSIHtFt, double p_dblBasalArea, double p_dblAvgDbh)
            {
                double dblSI = 0;
                int Total_Age = p_intSIDiaAge;
                if (Total_Age > 0)
                {
                    Total_Age = Total_Age + 12;
                }
                else
                {
                    return dblSI;
                }

                double dblCCF = 0;
                if (p_dblAvgDbh > 0 && p_dblBasalArea > 0)
                    dblCCF = 50.58 + 5.25 * (p_dblBasalArea / p_dblAvgDbh);
                if (dblCCF > 125)
                    dblCCF = dblCCF - 125;

                dblSI = (p_intSIHtFt - 9.89331 + 0.19177 * Total_Age - 0.00124 * Math.Pow(Total_Age, 2)) / (-0.00082 * dblCCF + 0.01387 * Total_Age - 0.0000455 * Math.Pow(Total_Age, 2));
                return dblSI;
            }

            /// <summary>
            /// SITE INDEX FOR ASPEN - Edminister
            /// Site Index Curves for Aspen in the Central Rocky Mountains
            /// Research Note: RM-453
            /// Derived from VBA source code by Don Vandendriese for FIA2FVS from RMRS
            /// Base age is 80 years
            /// Curves validated by jsfried
            /// </summary>
            /// <param name="p_intSIDiaAge">Age of site tree (Ring count at breast height)</param>
            /// <param name="p_intSIHtFt">Diameter of site tree</param>
            private double SI_AS1(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                dblSI = 4.5 + 0.48274 * (p_intSIHtFt - 4.5) * Math.Pow((1 - Math.Exp(-0.007719 * p_intSIDiaAge)), -0.93972);
                return dblSI;
            }


            /// <summary>
            /// SITE INDEX FOR PONDEROSA PINE - MEYER
		    /// Forest Service Inventory procedures for Eastern Washington,
            /// (Approximates Meyer in USDA Tech. bull. 630)
            /// Derived from VBA source code by Don Vandendriese for FIA2FVS from RMRS
            /// Base age is 100 years
            /// Site index curves validated by jsfried
            /// </summary>
            /// <param name="p_intSIDiaAge">Age of site tree (Ring count at breast height)</param>
            /// <param name="p_intSIHtFt">Diameter of site tree</param>
            private double SI_PP6(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                int Total_Age = p_intSIDiaAge;
                // Adjustment because this equation uses total tree age
                if (Total_Age > 0) Total_Age = Total_Age + 9;
                dblSI = (5.328 * (Math.Pow(Total_Age, -0.1)) - 2.378) * (p_intSIHtFt - 4.5) + 4.5;
                dblSI = dblSI + 0.49;
                return dblSI;
            }

            /// <summary>
            /// SITE INDEX FOR DOUGLAS-FIR - Monserud
            /// Applying Height Growth and Site Index Curves for Inland DOUGLAS-FIR
            /// Research Paper:  INT-347  1985
            /// Derived from Kurt's PL/SQL
            /// Base age is 50 years
            /// site index curves validated by jsfried

            /// </summary>
            /// <param name="p_intSIDiaAge">Age of site tree (Ring count at breast height)</param>
            /// <param name="p_intSIHtFt">Diameter of site tree</param>
            /// <param name="p_strHabTypeCd">Habitat type code for condition</param>
            private double SI_DF2(int p_intSIDiaAge, int p_intSIHtFt, string p_strHabTypeCd)
            {
                double dblSI = 0;
                int intHabTypeCd = 0;
                // habTypeCd is stored as text in the cond table; Safely try to parse to an int
                bool isHabTypeAnInt = int.TryParse(p_strHabTypeCd, out intHabTypeCd);
                if (!isHabTypeAnInt)
                {
                    // if habTypeCd is not an int, set it to the middle value
                    intHabTypeCd = 500;
                }
                double dblC1 = 0;
                double dblC2 = 0;
                double dblC3 = 0;
                if (intHabTypeCd <= 400)
                {
                    dblC1 = 1.0;
                }
                else if (intHabTypeCd < 530)
                {
                    dblC2 = 1.0;
                }
                else if (intHabTypeCd >= 530)
                {
                    dblC3 = 1.0;
                }
                dblSI = (38.787 - (2.805 * (Math.Log(p_intSIDiaAge) * Math.Log(p_intSIDiaAge))) +
                        (0.0216 * p_intSIDiaAge * Math.Log(p_intSIDiaAge)) +
                        (0.4948 * dblC1 + 0.4305 * dblC2 + 0.3964 * dblC3) * p_intSIHtFt +
                        (25.315 * dblC1 + 28.415 * dblC2 + 30.008 * dblC3) * p_intSIHtFt / p_intSIDiaAge) + 4.5;

                return dblSI;
            }

            /// <summary>
            /// This equation for Interior Western Red Cedar is from Hegyi and others
            /// Site Index Equations and curves for the Major Tree Species in B.C.
            /// -- B.C. Min of Forest Inventory Report 189(1) Nov 1979, Rev Sept 1981, p.6
            /// -- Site index at total tree age of 100 -- Equation #272
            /// ---Derived from Kurt's PL/SQL
            ///OUTPUT HAS NOT BEEN VALIDATED; METHOD NOT CURRENTLY IN USE
            ///<param name="p_intSIDiaAge"></param>
            ///<param name="p_intSIHtFt"></param>
            /// </summary>
            private double THPL03(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                // PL/SQL:  HT/(1.3283*(1-EXP(-0.0174*AGE))**1.4711);
                dblSI = p_intSIHtFt / (1.3283 * Math.Pow(1 - Math.Exp(-0.0174 * p_intSIDiaAge), 1.4711));
                return dblSI;
            }

            /// <summary>
            ///--- SITE INDEX RED ALDER - Harrington
		    ///    Height Growth and Site Index Curves for Red Alder
		    ///    Research Paper:  PNW-358 1986 p.5
            /// ---Derived from VBA source code by Don Vandendriese for FIA2FVS from RMRS
            ///OUTPUT HAS NOT BEEN VALIDATED; METHOD NOT CURRENTLY IN USE
            ///<param name="p_intSIDiaAge"></param>
            ///<param name="p_intSIHtFt"></param>
            /// </summary>
            private double SI_RA1(int p_intSIDiaAge, int p_intSIHtFt)
            {
                double dblSI = 0;
                double a = 54.185 - 4.61694 * p_intSIDiaAge + 0.11065 * Math.Pow(p_intSIDiaAge, 2) - 0.0007633 * Math.Pow(p_intSIDiaAge, 3);
                double b = 1.25934 - (0.012989) * p_intSIDiaAge + 3.522 * Math.Pow((1.0 / p_intSIDiaAge), 3);
                dblSI = a + b * p_intSIHtFt;
                dblSI = dblSI + 0.49;
                return dblSI;
            }


        }

        private IDictionary<String, String> LoadSiteIndexEquations_access(string strVariant)
        {
            //instantiate the dictionary so we can add equation records
            IDictionary<String, String> _dictSiteIdxEq = new Dictionary<String, String>();
            ado_data_access oAdo = new ado_data_access();
            //create env object so we can get the appDir
            env pEnv = new env();
            //open the project db file; db name is hard-coded
            oAdo.OpenConnection(oAdo.getMDBConnString(m_strProjDir + "\\db\\ref_master.mdb", "", ""));
            string strSQL = "select * from site_index_equations where FVS_VARIANT = '" + strVariant + "'";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    if (oAdo.m_OleDbDataReader["FVS_VARIANT"] == System.DBNull.Value ||
                        oAdo.m_OleDbDataReader["FIA_SPCD"] == System.DBNull.Value ||
                        oAdo.m_OleDbDataReader["EQUATION"] == System.DBNull.Value)
                    {
                        //If either variant, spcd, or equation is null, we don't add because we can't use
                    }
                    else
                    {
                        string strFvsVariant = Convert.ToString(oAdo.m_OleDbDataReader["FVS_VARIANT"]);
                        string strFiaSpCd = Convert.ToString(oAdo.m_OleDbDataReader["FIA_SPCD"]);
                        string strRegion = SI_EMPTY;
                        if (oAdo.m_OleDbDataReader["REGION"] != System.DBNull.Value)
                        {
                            strRegion = Convert.ToString(oAdo.m_OleDbDataReader["REGION"]);
                        }
                        string strSlfSpcd = SI_EMPTY;
                        if (oAdo.m_OleDbDataReader["SLF_SPCD"] != System.DBNull.Value)
                        {
                            strSlfSpcd = Convert.ToString(oAdo.m_OleDbDataReader["SLF_SPCD"]);
                        }
                        string strEquation = Convert.ToString(oAdo.m_OleDbDataReader["EQUATION"]);
                        string strValue = strEquation + SI_DELIM + strSlfSpcd + SI_DELIM + strRegion;
                        _dictSiteIdxEq.Add(strFvsVariant + SI_DELIM + strFiaSpCd, strValue);
                    }
                }
            }
            // Always close the connection
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
            return _dictSiteIdxEq;
        }

        private IDictionary<String, String> LoadSiteIndexEquations(string strVariant)
        {
            //instantiate the dictionary so we can add equation records
            IDictionary<String, String> _dictSiteIdxEq = new Dictionary<String, String>();
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            //create env object so we can get the appDir
            env pEnv = new env();
            //open the reference db file
            string strCon = oDataMgr.GetConnectionString(frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceSqliteFile);
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strCon))
            {
                con.Open();

                oDataMgr.m_strSQL = "SELECT * FROM " + Tables.Reference.DefaultSiteIndexEquationsTable + 
                    " WHERE FVS_VARIANT = '" + strVariant + "'";
                oDataMgr.SqlQueryReader(con, oDataMgr.m_strSQL);

                if (oDataMgr.m_DataReader.HasRows)
                {
                    while (oDataMgr.m_DataReader.Read())
                    {
                        if (oDataMgr.m_DataReader["FVS_VARIANT"] == System.DBNull.Value ||
                            oDataMgr.m_DataReader["FIA_SPCD"] == System.DBNull.Value ||
                            oDataMgr.m_DataReader["EQUATION"] == System.DBNull.Value)
                        {
                            //If either variant, spcd, or equation is null, we don't add because we can't use
                        }
                        else
                        {
                            string strFvsVariant = Convert.ToString(oDataMgr.m_DataReader["FVS_VARIANT"]);
                            string strFiaSpCd = Convert.ToString(oDataMgr.m_DataReader["FIA_SPCD"]);
                            string strRegion = SI_EMPTY;
                            if (oDataMgr.m_DataReader["REGION"] != System.DBNull.Value)
                            {
                                strRegion = Convert.ToString(oDataMgr.m_DataReader["REGION"]);
                            }
                            string strSlfSpcd = SI_EMPTY;
                            if (oDataMgr.m_DataReader["SLF_SPCD"] != System.DBNull.Value)
                            {
                                strSlfSpcd = Convert.ToString(oDataMgr.m_DataReader["SLF_SPCD"]);
                            }
                            string strEquation = SI_EMPTY;
                            if (oDataMgr.m_DataReader["EQUATION"] != System.DBNull.Value)
                            {
                                strEquation = Convert.ToString(oDataMgr.m_DataReader["EQUATION"]);
                            }
                            string strBaseAge = SI_EMPTY;
                            if (oDataMgr.m_DataReader["BASE_AGE"] != System.DBNull.Value)
                            {
                                strBaseAge = Convert.ToString(oDataMgr.m_DataReader["BASE_AGE"]);
                            }

                            string strValue = strEquation + SI_DELIM + strSlfSpcd + SI_DELIM + strRegion + SI_DELIM + strBaseAge;
                            if (!_dictSiteIdxEq.ContainsKey(strFvsVariant + SI_DELIM + strFiaSpCd))
                            {
                                _dictSiteIdxEq.Add(strFvsVariant + SI_DELIM + strFiaSpCd, strValue);
                            }
                            
                        }
                    }
                }
            }
            oDataMgr = null;
            return _dictSiteIdxEq;
        }

        
        public void StartFIA2FVS(ODBCMgr odbcmgr, dao_data_access oDao,
            ado_data_access oAdo, string strTempMDB, bool bOverwrite,
            string strDebugFile, string strVariant, List<string> lstStates, string strSourceStandTableAlias, string strSourceTreeTableAlias)
        {
            m_strTempMDBFileNewProcess = strTempMDB;

            // Copy the target database from BioSum application directory
            string applicationDb = frmMain.g_oEnv.strAppDir + "\\db\\" + Tables.FIA2FVS.DefaultFvsInputFile;
            string strInDirAndFile = m_strDataDirectory + "\\" + Tables.FIA2FVS.DefaultFvsInputFile;
            if (!File.Exists(strInDirAndFile))
            {
                // create FVSIn.db if missing
                File.Copy(applicationDb, strInDirAndFile, true);
            }

            // Create temporary SQLite database for the variant being worked on
            env oEnv = new env();
            string strVariantWorkDb = frmMain.g_oUtils.getRandomFile(oEnv.strTempDir, "db");
            File.Copy(applicationDb, strVariantWorkDb, true);

            int intProgressBarCounter = 0;

            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            string strWorkDbConn = oDataMgr.GetConnectionString(strVariantWorkDb);
            string strTargetDbConn = oDataMgr.GetConnectionString(strInDirAndFile);
            this.m_strDebugFile = strDebugFile;
            DebugLogMessage("*****START*****" + System.DateTime.Now.ToString() + "\r\n");
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strTargetDbConn))
            {
                // Connect to target database (FVSIn.db)
                con.Open();
                bool bCreateTables = false;
                if (!oDataMgr.TableExist(con, Tables.FIA2FVS.DefaultFvsInputStandTableName) ||
                    !oDataMgr.TableExist(con, Tables.FIA2FVS.DefaultFvsInputTreeTableName))
                {
                    bCreateTables = true;
                }
                if (bCreateTables == true)
                {
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                        "Creating FVS_INIT_COND tables For Variant " + strVariant);
                    CreateFVSInputTables(con);
                }
                else if (bCreateTables == false && lstStates.Count > 0)
                {
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                        "Deleting existing records from FVS_STANDINIT_COND and FVS_TREEINIT_COND tables For Variant " + strVariant);

                    // Delete existing state records if there is conflict with source
                    string csv = String.Join(",", lstStates.Select(x => x.ToString()).ToArray());

                    // Delete existing records from tree table
                    string strSql = "DELETE FROM " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                            " WHERE stand_cn IN(select stand_cn from " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                            " where state in (" + csv + ") AND VARIANT = '" + strVariant + "')";
                    DebugLogSQL(strSql);
                    oDataMgr.SqlNonQuery(con, strSql);

                    // Delete existing records from stand table
                    strSql = "DELETE FROM " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                            " WHERE STATE IN (" + csv + ")" +
                            " AND VARIANT = '" + strVariant + "'";
                    DebugLogSQL(strSql);
                    oDataMgr.SqlNonQuery(con, strSql);
                }
                else if (bOverwrite == true)
                {
                    // Delete existing records from tree table that correspond to selected variant
                    string strSql = "DELETE FROM " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                        " WHERE TRIM(VARIANT) = '" + strVariant + "'";
                    DebugLogSQL(strSql);
                    oDataMgr.SqlNonQuery(con, strSql);

                    // Delete existing records from stand table for selected variant
                    strSql = "DELETE FROM " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                        " WHERE TRIM(VARIANT) = '" + strVariant + "'";
                    DebugLogSQL(strSql);
                    oDataMgr.SqlNonQuery(con, strSql);
                }


                // Ensure FIADB indexes are in plac
                AddFIADBIndexes(con);
            }

            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strWorkDbConn))
            {
                con.Open();

                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                        "Creating work FVS_INIT_COND tables For Variant " + strVariant);
                CreateFVSInputTables(con);

                // Ensure FIADB indexes are in place
                AddFIADBIndexes(con);
            }

            frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
            // Check to see if the target DSN exists and if so, delete so we can add
            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.Fia2FvsOutputDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.Fia2FvsOutputDsnName);
            }
            odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.Fia2FvsOutputDsnName, strVariantWorkDb);
            if (!string.IsNullOrEmpty(odbcmgr.m_strError))
            {
                DebugLogSQL("ODBCMgr error: " + odbcmgr.m_strError);
                return;
            }
            else
            {
                DebugLogSQL("Created DSN for " + ODBCMgr.DSN_KEYS.Fia2FvsOutputDsnName);
            }
            // Link stand table to temp database
            oDao.CreateSQLiteTableLink(m_strTempMDBFileNewProcess, Tables.FIA2FVS.DefaultFvsInputStandTableName, Tables.FIA2FVS.DefaultFvsInputStandTableName,
                ODBCMgr.DSN_KEYS.Fia2FvsOutputDsnName, strVariantWorkDb);
            // Set the index, required to by ODBC to update
            if (!oDao.IndexExists(m_strTempMDBFileNewProcess, Tables.FIA2FVS.DefaultFvsInputStandTableName, "STAND_CN"))
            {
                oDao.CreatePrimaryKeyIndex(m_strTempMDBFileNewProcess, Tables.FIA2FVS.DefaultFvsInputStandTableName, "STAND_CN");
            }
            // Link tree table to temp database
            oDao.CreateSQLiteTableLink(m_strTempMDBFileNewProcess, Tables.FIA2FVS.DefaultFvsInputTreeTableName, Tables.FIA2FVS.DefaultFvsInputTreeTableName,
                ODBCMgr.DSN_KEYS.Fia2FvsOutputDsnName, strVariantWorkDb);
            // Set the index, required to by ODBC to update
            if (!oDao.IndexExists(m_strTempMDBFileNewProcess, Tables.FIA2FVS.DefaultFvsInputTreeTableName, "TREE_CN"))
            {
                oDao.CreatePrimaryKeyIndex(m_strTempMDBFileNewProcess, Tables.FIA2FVS.DefaultFvsInputTreeTableName, "TREE_CN");
            }

            // Link DWM tables
            string strMasterAuxDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\master_aux.db";
            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.DWMDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.DWMDsnName);
            }
            odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.DWMDsnName, strMasterAuxDb);
            string strBiosumRefDb = frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                        frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceSqliteFile;
            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.BiosumRefDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.BiosumRefDsnName);
            }
            odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.BiosumRefDsnName, strBiosumRefDb);

            if (!string.IsNullOrEmpty(odbcmgr.m_strError))
            {
                DebugLogSQL("ODBCMgr error: " + odbcmgr.m_strError);
                return;
            }
            else
            {
                DebugLogSQL("Created DSN for " + ODBCMgr.DSN_KEYS.DWMDsnName);
            }

            // Attach tables needed for Site Index and Fuel columns
            LinkSiteIndexAndFuelColumnsTables();

            frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                "Populating FVS_STANDINIT_COND and FVS_TREEINIT_COND tables For Variant " + strVariant);
            string strAccdbConnection = oAdo.getMDBConnString(m_strTempMDBFileNewProcess, "", "");
            using (OleDbConnection oAccessConn = new OleDbConnection(strAccdbConnection))
            {
                oAccessConn.Open();

                DebugLogSQL(Queries.FVS.FVSInput.StandInit.PopulateStandInitOld(strSourceStandTableAlias, this.m_strCondTable,
                    strVariant));
                oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.StandInit.PopulateStandInitOld(strSourceStandTableAlias, this.m_strCondTable,
                    strVariant));

                DebugLogSQL(Queries.FVS.FVSInput.TreeInit.PopulateTreeInitOld(strSourceTreeTableAlias, strSourceStandTableAlias,
                    this.m_strCondTable, this.m_strTreeTable, strVariant));
                oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.TreeInit.PopulateTreeInitOld(strSourceTreeTableAlias, strSourceStandTableAlias,
                    this.m_strCondTable, this.m_strTreeTable, strVariant));

                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                    "Customizing FVS_STANDINIT_COND table For Variant " + strVariant);
                // Populate the STAND_ID and SAM_WT from the BioSum Cond table
                DebugLogSQL(Queries.FVS.FVSInput.StandInit.UpdateFromCondOld(this.m_strCondTable, strVariant));
                oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.StandInit.UpdateFromCondOld(this.m_strCondTable, strVariant));

                //Overwrite FOREST_TYPE with FOREST_TYPE_FIA, PV_CODE with PV_FIA_HABTYPCD1
                DebugLogSQL(Queries.FVS.FVSInput.StandInit.UpdateForestTypeAndPvCode());
                oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.StandInit.UpdateForestTypeAndPvCode());

                //Null FUEL_MODEL if not checked
                if (new int[]
                {
                    (int) m_enumDWMOption.SKIP_FUEL_MODEL_AND_DWM_DATA,
                    (int) m_enumDWMOption.USE_DWM_DATA_ONLY
                }.Contains(m_intDWMOption))
                {
                    DebugLogSQL(Queries.FVS.FVSInput.StandInit.SetFuelModelToNull());
                    oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.StandInit.SetFuelModelToNull());
                }

                //Create lookup table for DWM Fuelbed Type Code to convert to FVS format
                if (m_intError == 0 && new int[]
                {
                    (int) m_enumDWMOption.USE_FUEL_MODEL_OR_DWM_DATA,
                    (int) m_enumDWMOption.USE_FUEL_MODEL_ONLY
                }.Contains(m_intDWMOption))
                {

                    CreateDwmFuelbedTypeCodeFvsConversionTable(oAccessConn);
                    m_ado.m_strSQL =
                        Queries.FVS.FVSInput.StandInit.UpdateFuelModel(Tables.FIA2FVS.DefaultFvsInputStandTableName, m_strCondTable);
                    m_ado.SqlNonQuery(oAccessConn, m_ado.m_strSQL);
                    m_ado.SqlNonQuery(oAccessConn, "DROP TABLE Ref_DWM_Fuelbed_Type_Codes;");
                }

                // SITE_INDEX and SITE_SPECIES 
                if (m_intError == 0)
                {
                    GenerateSiteIndexAndSiteSpeciesSQL(strVariant);
                }
                
                // set fuel columns to null
                
                using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(strWorkDbConn))
                {
                    workConn.Open();
                    DebugLogSQL(Queries.FVS.FVSInput.StandInit.SetFuelColumnsToNull());
                    oDataMgr.SqlNonQuery(workConn, Queries.FVS.FVSInput.StandInit.SetFuelColumnsToNull());
                }
                // Copy DWM fields from FVSIn.accdb
                if (m_intError == 0 && new int[]
                {
                    (int) m_enumDWMOption.USE_FUEL_MODEL_OR_DWM_DATA,
                    (int) m_enumDWMOption.USE_DWM_DATA_ONLY
                }.Contains(m_intDWMOption))
                {
                    PopulateFuelColumns(strAccdbConnection, strVariantWorkDb, Tables.FIA2FVS.DefaultFvsInputStandTableName);
                }

                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                    "Customizing FVS_TREEINIT_COND table For Variant " + strVariant);
                // Create temp table in MS Access
                string strTempTreeTable = "temp_tree";
                if (oAdo.TableExist(oAccessConn, strTempTreeTable))
                {
                    oAdo.SqlNonQuery(oAccessConn, "DROP TABLE " + strTempTreeTable);
                }
                string strSQL = "SELECT * INTO " + strTempTreeTable + " FROM " + Tables.FIA2FVS.DefaultFvsInputTreeTableName;
                DebugLogSQL(strSQL);
                oAdo.SqlNonQuery(oAccessConn, strSQL);

                //Add indexes to improve performance
                string[] arrTempTreeIndexes = new string[] { "TEMP_TREE_TREE_CN_IDX", "TEMP_TREE_TREE_ID_IDX" };
                string[] arrTempTreeFields = new string[] { "TREE_CN", "TREE_ID" };
                for (int i = 0; i < arrTempTreeIndexes.Length; i++)
                {
                    if (oAdo.ColumnExist(oAccessConn, strTempTreeTable, arrTempTreeFields[i]))
                    {
                        oAdo.AddIndex(oAccessConn, strTempTreeTable, arrTempTreeIndexes[i], arrTempTreeFields[i]);
                    }
                }

                // These are the fields we are changing that need to be updated at the end
                IList<string> lstFields = new List<string>();

                // Delete seedlings if not desired
                if (!bIncludeSeedlings)
                {
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
                    DebugLogSQL(Queries.FVS.FVSInput.TreeInit.DeleteSeedlings(strTempTreeTable));
                    oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.TreeInit.DeleteSeedlings(strTempTreeTable));
                    DebugLogSQL(Queries.FVS.FVSInput.TreeInit.DeleteSeedlings(Tables.FIA2FVS.DefaultFvsInputTreeTableName));
                    oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.TreeInit.DeleteSeedlings(Tables.FIA2FVS.DefaultFvsInputTreeTableName));
                }

                // Populate the STAND_ID from the BioSum Cond table
                DebugLogSQL(Queries.FVS.FVSInput.TreeInit.UpdateFromCond(this.m_strCondTable, strVariant, strTempTreeTable));
                oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.TreeInit.UpdateFromCond(this.m_strCondTable, strVariant, strTempTreeTable));
                lstFields.Add("STAND_ID");

                // Populate the TREE_ID and BioSum Tree table
                DebugLogSQL(Queries.FVS.FVSInput.TreeInit.UpdateFromTree(this.m_strTreeTable, strTempTreeTable));
                oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.TreeInit.UpdateFromTree(this.m_strTreeTable, strTempTreeTable));
                lstFields.Add("TREE_ID");

                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
                // Update tree_count
                DebugLogSQL(Queries.FVS.FVSInput.TreeInit.UpdateTreeCount(this.m_strTreeTable, strTempTreeTable));
                oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.TreeInit.UpdateTreeCount(this.m_strTreeTable, strTempTreeTable));
                lstFields.Add("TREE_COUNT");

                // Manage calibraton columns; Set to null if not selected
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
                // 30-OCT-2023: Start using DG and HTG from FIA2FVS. Set to null if calibration data not selected
                if (!bUsePrevHtCalibrationData)
                {
                    DebugLogSQL(Queries.FVS.FVSInput.TreeInit.SetCalibrationColumnsToNull(strTempTreeTable, "HTG"));
                    oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.TreeInit.SetCalibrationColumnsToNull(strTempTreeTable, "HTG"));
                    lstFields.Add("HTG");
                }
                if (!bUsePrevDiaCalibrationData)
                {
                    DebugLogSQL(Queries.FVS.FVSInput.TreeInit.SetCalibrationColumnsToNull(strTempTreeTable, "DG"));
                    oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.TreeInit.SetCalibrationColumnsToNull(strTempTreeTable, "DG"));
                    lstFields.Add("DG");
                }

                // Update damage codes for cull and mistletoe
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
                UpdateDamageCodesForCullAndMistletoe(m_strTreeTable, strTempMDB, strTempTreeTable);
                lstFields.Add("Damage1");
                lstFields.Add("Severity1");
                lstFields.Add("Damage2");
                lstFields.Add("Severity2");
                lstFields.Add("Damage3");
                lstFields.Add("Severity3");
                lstFields.Add("TREEVALUE");

                // Copy modified fields to final FVSIn.db
                DebugLogSQL(Queries.FVS.FVSInput.TreeInit.UpdateFieldsFromTempTable(strTempTreeTable, lstFields));
                oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.TreeInit.UpdateFieldsFromTempTable(strTempTreeTable, lstFields));

                // Delete link to target stand table from temp database
                string strSql = "DROP TABLE " + Tables.FIA2FVS.DefaultFvsInputStandTableName;
                DebugLogSQL(strSql);
                oAdo.SqlNonQuery(oAccessConn, strSql);

                // Delete link to target tree table from temp database
                strSql = "DROP TABLE " + Tables.FIA2FVS.DefaultFvsInputTreeTableName;
                DebugLogSQL(strSql);
                oAdo.SqlNonQuery(oAccessConn, strSql);
            }

            // Make changes to the FIA2FVS records as needed
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strWorkDbConn))
            {
                con.Open();
                oDataMgr.m_strSQL = "UPDATE " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                                " SET GROUPS = '" + strGroup + "'";
                DebugLogSQL(oDataMgr.m_strSQL);
                oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);

                // Overwrite BASAL_AREA_FACTOR, INV_PLOT_SIZE, BRK_DBH, NUM_PLOTS in stand table
                // to prevent inflated TPAs for variants outside of CA, OR, WA
                oDataMgr.m_strSQL = Queries.FVS.FVSInput.StandInit.OverwriteFieldsForTPA();
                DebugLogSQL(oDataMgr.m_strSQL);
                oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);

                // Overwrite PLOT_ID in tree table
                // to prevent inflated TPAs for variants outside of CA, OR, WA
                oDataMgr.m_strSQL = Queries.FVS.FVSInput.TreeInit.OverwritePlotID();
                DebugLogSQL(oDataMgr.m_strSQL);
                oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);

                oDataMgr.m_strSQL = "ATTACH DATABASE '" + strInDirAndFile + "' AS target";
                oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);

                // Populate variant field in tree table
                oDataMgr.m_strSQL = "UPDATE " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                    " SET VARIANT = '" + strVariant + "' WHERE VARIANT IS NULL";
                oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);

                oDataMgr.m_strSQL = "INSERT INTO target." + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                    " SELECT * FROM " + Tables.FIA2FVS.DefaultFvsInputStandTableName;
                oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);

                oDataMgr.m_strSQL = "INSERT INTO target." + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                    " SELECT * FROM " + Tables.FIA2FVS.DefaultFvsInputTreeTableName;
                oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);
            }


            UpdateFvsInSqliteConfigurationTable();

            // Remove DSNs if they exist
            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.Fia2FvsOutputDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.Fia2FvsOutputDsnName);
            }
            if (!string.IsNullOrEmpty(odbcmgr.m_strError))
            {
                DebugLogSQL("ODBCMgr error: " + odbcmgr.m_strError);
            }
            else
            {
                DebugLogSQL("Removed DSN for " + ODBCMgr.DSN_KEYS.Fia2FvsOutputDsnName);
            }
            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.DWMDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.DWMDsnName);
            }
            if (!string.IsNullOrEmpty(odbcmgr.m_strError))
            {
                DebugLogSQL("ODBCMgr error: " + odbcmgr.m_strError);
            }
            else
            {
                DebugLogSQL("Removed DSN for " + ODBCMgr.DSN_KEYS.DWMDsnName);
            }
            DebugLogMessage("*****END*****" + System.DateTime.Now.ToString() + "\r\n", 2);
        }

        public void StartFIA2FVSNew(SQLite.ADO.DataMgr oDataMgr, bool bOverwrite, string strMasterDb,
            string strDebugFile, string strVariant, List<string> lstStates)
        {
            // Copy the target database from BioSum application directory
            string applicationDb = frmMain.g_oEnv.strAppDir + "\\db\\" + Tables.FIA2FVS.DefaultFvsInputFile;
            string strInDirAndFile = m_strDataDirectory + "\\" + Tables.FIA2FVS.DefaultFvsInputFile;
            if (!File.Exists(strInDirAndFile))
            {
                // create FVSIn.db if missing
                File.Copy(applicationDb, strInDirAndFile, true);
            }

            int intProgressBarCounter = 0;

            // Create temporary SQLite database for the variant being worked on
            env oEnv = new env();
            string strTempDb = frmMain.g_oUtils.getRandomFile(oEnv.strTempDir, "db");
            File.Copy(applicationDb, strTempDb, true);

            string strTempDbConn = oDataMgr.GetConnectionString(strTempDb);
            string strTargetDbConn = oDataMgr.GetConnectionString(strInDirAndFile);
            m_strDebugFile = strDebugFile;
            DebugLogMessage("*****START*****" + System.DateTime.Now.ToString() + "\r\n");

            using (System.Data.SQLite.SQLiteConnection targetConn = new System.Data.SQLite.SQLiteConnection(strTargetDbConn))
            {
                // Connect to target database (FVSIn.db)
                targetConn.Open();
                bool bCreateTables = false;
                if (!oDataMgr.TableExist(targetConn, Tables.FIA2FVS.DefaultFvsInputStandTableName) ||
                    !oDataMgr.TableExist(targetConn, Tables.FIA2FVS.DefaultFvsInputTreeTableName))
                {
                    bCreateTables = true;
                }
                if (bCreateTables == true)
                {
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                        "Creating FVS_INIT_COND tables For Variant " + strVariant);
                    CreateFVSInputTables(targetConn);
                }
                else if (bCreateTables == false && lstStates.Count > 0)
                {
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                        "Deleting existing records from FVS_STANDINIT_COND and FVS_TREEINIT_COND tables For Variant " + strVariant);

                    // Delete existing state records if there is conflict with source
                    string csv = String.Join(",", lstStates.Select(x => x.ToString()).ToArray());

                    // Delete existing records from tree table
                    oDataMgr.m_strSQL = "DELETE FROM " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                            " WHERE stand_cn IN(SELECT stand_cn from " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                            " WHERE state IN (" + csv + ") AND VARIANT = '" + strVariant + "')";
                    DebugLogSQL(oDataMgr.m_strSQL);
                    oDataMgr.SqlNonQuery(targetConn, oDataMgr.m_strSQL);

                    // Delete existing records from stand table
                    oDataMgr.m_strSQL = "DELETE FROM " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                            " WHERE STATE IN (" + csv + ")" +
                            " AND VARIANT = '" + strVariant + "'";
                    DebugLogSQL(oDataMgr.m_strSQL);
                    oDataMgr.SqlNonQuery(targetConn, oDataMgr.m_strSQL);
                }
                else if (bOverwrite == true)
                {
                    // Delete existing records from tree table that correspond to selected variant
                    oDataMgr.m_strSQL = "DELETE FROM " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                        " WHERE TRIM(VARIANT) = '" + strVariant + "'";
                    DebugLogSQL(oDataMgr.m_strSQL);
                    oDataMgr.SqlNonQuery(targetConn, oDataMgr.m_strSQL);

                    // Delete existing records from stand table for selected variant
                    oDataMgr.m_strSQL = "DELETE FROM " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                        " WHERE TRIM(VARIANT) = '" + strVariant + "'";
                    DebugLogSQL(oDataMgr.m_strSQL);
                    oDataMgr.SqlNonQuery(targetConn, oDataMgr.m_strSQL);
                }

                // Ensure FIADB indexes are in place
                AddFIADBIndexes(targetConn);
            }

            using (System.Data.SQLite.SQLiteConnection tempConn = new System.Data.SQLite.SQLiteConnection(strTempDbConn))
            {
                tempConn.Open();

                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                        "Creating work FVS_INIT_COND tables For Variant " + strVariant);
                CreateFVSInputTables(tempConn);

                // Ensure FIADB indexes are in place
                AddFIADBIndexes(tempConn);

                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                "Populating FVS_STANDINIT_COND and FVS_TREEINIT_COND tables For Variant " + strVariant);

                // Attach FIADB and master databases
                if (!oDataMgr.DatabaseAttached(tempConn, strSourceFiaDb))
                {
                    oDataMgr.m_strSQL = "ATTACH DATABASE '" + strSourceFiaDb + "' AS source";
                    oDataMgr.SqlNonQuery(tempConn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.DatabaseAttached(tempConn, strMasterDb))
                {
                    oDataMgr.m_strSQL = "ATTACH DATABASE '" + strMasterDb + "' AS master";
                    oDataMgr.SqlNonQuery(tempConn, oDataMgr.m_strSQL);
                }

                // Get columns in from FIADB version of tables
                string strStandInitFields = oDataMgr.getFieldNames(tempConn, "SELECT * FROM source." + Tables.FIA2FVS.DefaultFvsInputStandTableName);
                string strTreeInitFields = oDataMgr.getFieldNames(tempConn, "SELECT * FROM source." + Tables.FIA2FVS.DefaultFvsInputTreeTableName);

                DebugLogSQL(Queries.FVS.FVSInput.StandInit.PopulateStandInit(Tables.FIA2FVS.DefaultFvsInputStandTableName, strStandInitFields,
                    m_strCondTable, strVariant));
                oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.StandInit.PopulateStandInit("source." + Tables.FIA2FVS.DefaultFvsInputStandTableName, strStandInitFields,
                    "master." + m_strCondTable, strVariant));

                DebugLogSQL(Queries.FVS.FVSInput.TreeInit.PopulateTreeInit(Tables.FIA2FVS.DefaultFvsInputTreeTableName, strTreeInitFields,
                    Tables.FIA2FVS.DefaultFvsInputStandTableName, m_strCondTable, m_strTreeTable, strVariant));
                oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.TreeInit.PopulateTreeInit("source." + Tables.FIA2FVS.DefaultFvsInputTreeTableName, strTreeInitFields,
                    "source." + Tables.FIA2FVS.DefaultFvsInputStandTableName, "master." + m_strCondTable, "master." + m_strTreeTable, strVariant));

                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                    "Customizing FVS_STANDINIT_COND table For Variant " + strVariant);
                // Populate the STAND_ID and SAM_WT from the BioSum Cond table
                DebugLogSQL(Queries.FVS.FVSInput.StandInit.UpdateFromCond(m_strCondTable, strVariant));
                oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.StandInit.UpdateFromCond("master." + m_strCondTable, strVariant));

                //Overwrite FOREST_TYPE with FOREST_TYPE_FIA, PV_CODE with PV_FIA_HABTYPCD1
                DebugLogSQL(Queries.FVS.FVSInput.StandInit.UpdateForestTypeAndPvCode());
                oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.StandInit.UpdateForestTypeAndPvCode());

                //Null FUEL_MODEL if not checked
                if (new int[]
                {
                    (int) m_enumDWMOption.SKIP_FUEL_MODEL_AND_DWM_DATA,
                    (int) m_enumDWMOption.USE_DWM_DATA_ONLY
                }.Contains(m_intDWMOption))
                {
                    DebugLogSQL(Queries.FVS.FVSInput.StandInit.SetFuelModelToNull());
                    oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.StandInit.SetFuelModelToNull());
                }

                //Create lookup table for DWM Fuelbed Type Code to convert to FVS format
                if (m_intError == 0 && new int[]
                {
                    (int) m_enumDWMOption.USE_FUEL_MODEL_OR_DWM_DATA,
                    (int) m_enumDWMOption.USE_FUEL_MODEL_ONLY
                }.Contains(m_intDWMOption))
                {
                    CreateDwmFuelbedTypeCodeFvsConversionTableNew(oDataMgr, tempConn);
                    oDataMgr.m_strSQL =
                        Queries.FVS.FVSInput.StandInit.UpdateFuelModelNew(Tables.FIA2FVS.DefaultFvsInputStandTableName, "master." + m_strCondTable);
                    oDataMgr.SqlNonQuery(tempConn, oDataMgr.m_strSQL);
                    oDataMgr.SqlNonQuery(tempConn, "DROP TABLE Ref_DWM_Fuelbed_Type_Codes");
                }

                // SITE_INDEX and SITE_SPECIES 
                if (m_intError == 0)
                {
                    GenerateSiteIndexAndSiteSpeciesSQLNew(oDataMgr, tempConn, strVariant);
                }

                // set fuel columns to null
                DebugLogSQL(Queries.FVS.FVSInput.StandInit.SetFuelColumnsToNull());
                oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.StandInit.SetFuelColumnsToNull());

                // Copy DWM fields
                if (m_intError == 0 && new int[]
                {
                    (int) m_enumDWMOption.USE_FUEL_MODEL_OR_DWM_DATA,
                    (int) m_enumDWMOption.USE_DWM_DATA_ONLY
                }.Contains(m_intDWMOption))
                {
                    PopulateFuelColumnsNew(oDataMgr, tempConn, Tables.FIA2FVS.DefaultFvsInputStandTableName);
                }

                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.lblMsg, "Text",
                    "Customizing FVS_TREEINIT_COND table For Variant " + strVariant);
                // Create temp table in temp database
                string strTempTreeTable = "temp_tree";
                if (oDataMgr.TableExist(tempConn, strTempTreeTable))
                {
                    oDataMgr.SqlNonQuery(tempConn, "DROP TABLE " + strTempTreeTable);
                }
                oDataMgr.m_strSQL = "CREATE TABLE " + strTempTreeTable + " AS SELECT * FROM " + Tables.FIA2FVS.DefaultFvsInputTreeTableName;
                DebugLogSQL(oDataMgr.m_strSQL);
                oDataMgr.SqlNonQuery(tempConn, oDataMgr.m_strSQL);

                //Add indexes to improve performance
                string[] arrTempTreeIndexes = new string[] { "TEMP_TREE_TREE_CN_IDX", "TEMP_TREE_TREE_ID_IDX" };
                string[] arrTempTreeFields = new string[] { "TREE_CN", "TREE_ID" };
                for (int i = 0; i < arrTempTreeIndexes.Length; i++)
                {
                    if (oDataMgr.ColumnExist(tempConn, strTempTreeTable, arrTempTreeFields[i]))
                    {
                        oDataMgr.AddIndex(tempConn, strTempTreeTable, arrTempTreeIndexes[i], arrTempTreeFields[i]);
                    }
                }

                // These are the fields we are changing that need to be updated at the end
                IList<string> lstFields = new List<string>();

                // Delete seedlings if not desired
                if (!bIncludeSeedlings)
                {
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
                    DebugLogSQL(Queries.FVS.FVSInput.TreeInit.DeleteSeedlingsNew(strTempTreeTable));
                    oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.TreeInit.DeleteSeedlingsNew(strTempTreeTable));
                    DebugLogSQL(Queries.FVS.FVSInput.TreeInit.DeleteSeedlingsNew(Tables.FIA2FVS.DefaultFvsInputTreeTableName));
                    oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.TreeInit.DeleteSeedlingsNew(Tables.FIA2FVS.DefaultFvsInputTreeTableName));
                }

                // Populate the STAND_ID from the BioSum Cond table
                DebugLogSQL(Queries.FVS.FVSInput.TreeInit.UpdateFromCondNew( m_strCondTable, strVariant, strTempTreeTable));
                oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.TreeInit.UpdateFromCondNew("master." + m_strCondTable, strVariant, strTempTreeTable));
                lstFields.Add("STAND_ID");

                // Populate the TREE_ID and BioSum Tree table
                DebugLogSQL(Queries.FVS.FVSInput.TreeInit.UpdateFromTreeNew(m_strTreeTable, strTempTreeTable));
                oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.TreeInit.UpdateFromTreeNew("master." + m_strTreeTable, strTempTreeTable));
                lstFields.Add("TREE_ID");

                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
                // Update tree_count
                DebugLogSQL(Queries.FVS.FVSInput.TreeInit.UpdateTreeCountNew(m_strTreeTable, strTempTreeTable));
                oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.TreeInit.UpdateTreeCountNew("master." + m_strTreeTable, strTempTreeTable));
                lstFields.Add("TREE_COUNT");

                // Manage calibraton columns; Set to null if not selected
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
                // 30-OCT-2023: Start using DG and HTG from FIA2FVS. Set to null if calibration data not selected
                if (!bUsePrevHtCalibrationData)
                {
                    DebugLogSQL(Queries.FVS.FVSInput.TreeInit.SetCalibrationColumnsToNull(strTempTreeTable, "HTG"));
                    oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.TreeInit.SetCalibrationColumnsToNull(strTempTreeTable, "HTG"));
                    lstFields.Add("HTG");
                }
                if (!bUsePrevDiaCalibrationData)
                {
                    DebugLogSQL(Queries.FVS.FVSInput.TreeInit.SetCalibrationColumnsToNull(strTempTreeTable, "DG"));
                    oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.TreeInit.SetCalibrationColumnsToNull(strTempTreeTable, "DG"));
                    lstFields.Add("DG");
                }

                // Update damage codes for cull and mistletoe
                frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar1, "Value", intProgressBarCounter++);
                UpdateDamageCodesForCullAndMistletoeNew(oDataMgr, tempConn, "master." + m_strTreeTable, strTempTreeTable);
                lstFields.Add("Damage1");
                lstFields.Add("Severity1");
                lstFields.Add("Damage2");
                lstFields.Add("Severity2");
                lstFields.Add("Damage3");
                lstFields.Add("Severity3");
                lstFields.Add("TREEVALUE");

                // Copy modified fields to fvs in tables
                DebugLogSQL(Queries.FVS.FVSInput.TreeInit.UpdateFieldsFromTempTableNew(strTempTreeTable, lstFields));
                oDataMgr.SqlNonQuery(tempConn, Queries.FVS.FVSInput.TreeInit.UpdateFieldsFromTempTableNew(strTempTreeTable, lstFields));

                oDataMgr.m_strSQL = "UPDATE " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                                " SET GROUPS = '" + strGroup + "'";
                DebugLogSQL(oDataMgr.m_strSQL);
                oDataMgr.SqlNonQuery(tempConn, oDataMgr.m_strSQL);

                // Overwrite BASAL_AREA_FACTOR, INV_PLOT_SIZE, BRK_DBH, NUM_PLOTS in stand table
                // to prevent inflated TPAs for variants outside of CA, OR, WA
                oDataMgr.m_strSQL = Queries.FVS.FVSInput.StandInit.OverwriteFieldsForTPA();
                DebugLogSQL(oDataMgr.m_strSQL);
                oDataMgr.SqlNonQuery(tempConn, oDataMgr.m_strSQL);

                // Overwrite PLOT_ID in tree table
                // to prevent inflated TPAs for variants outside of CA, OR, WA
                oDataMgr.m_strSQL = Queries.FVS.FVSInput.TreeInit.OverwritePlotID();
                DebugLogSQL(oDataMgr.m_strSQL);
                oDataMgr.SqlNonQuery(tempConn, oDataMgr.m_strSQL);

                // Populate variant field in tree table
                oDataMgr.m_strSQL = "UPDATE " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                    " SET VARIANT = '" + strVariant + "' WHERE VARIANT IS NULL";
                oDataMgr.SqlNonQuery(tempConn, oDataMgr.m_strSQL);

                // Copy data from temp tables to final tables
                if (!oDataMgr.DatabaseAttached(tempConn, strInDirAndFile))
                {
                    oDataMgr.m_strSQL = "ATTACH DATABASE '" + strInDirAndFile + "' AS target";
                    oDataMgr.SqlNonQuery(tempConn, oDataMgr.m_strSQL);
                }
                
                oDataMgr.m_strSQL = "INSERT INTO target." + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                    " SELECT * FROM " + Tables.FIA2FVS.DefaultFvsInputStandTableName;
                oDataMgr.SqlNonQuery(tempConn, oDataMgr.m_strSQL);

                oDataMgr.m_strSQL = "INSERT INTO target." + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                    " SELECT * FROM " + Tables.FIA2FVS.DefaultFvsInputTreeTableName;
                oDataMgr.SqlNonQuery(tempConn, oDataMgr.m_strSQL);
            }

            UpdateFvsInSqliteConfigurationTable();


        }
        private void CopyFuelColumns(string strTempMDBFile, string strVariant)
        {
            ado_data_access oAdo = new ado_data_access();
            this.m_strInDir = m_strDataDirectory + "\\" + strVariant.Trim();
            this.strFVSInDBFile = "FVSIn.accdb";
            if (File.Exists(this.m_strInDir + "\\" + this.strFVSInDBFile))
            {
                string strConn = oAdo.getMDBConnString(strTempMDBFile, "", "");
                using (OleDbConnection oAccessConn = new OleDbConnection(strConn))
                {
                    oAccessConn.ConnectionString = strConn;
                    oAccessConn.Open();
                    oAdo.OpenConnection(strConn);
                    DebugLogSQL(Queries.FVS.FVSInput.StandInit.CopyDwmColumns(strVariant));
                    oAdo.SqlNonQuery(oAccessConn, Queries.FVS.FVSInput.StandInit.CopyDwmColumns(strVariant));
                }
            }
            else
            {
                DebugLogSQL("Error: Unable to locate BioSum FVSIn.accdb database. DWM data cannot be copied!!");
            }
            //m_dao.CreateTableLink(this.m_strTempMDBFile, "FVS_TreeInit", this.m_strInDir + "\\" + this.strFVSInMDBFile,
            //    "FVS_TreeInit", true);

        }

        private void LinkSiteIndexAndFuelColumnsTables()
        {
            dao_data_access oDao = new dao_data_access();
            ado_data_access oAdo = new ado_data_access();

            using (OleDbConnection con = new OleDbConnection(oAdo.getMDBConnString(m_strTempMDBFileNewProcess, "", "")))
            {
                con.Open();

                // FIA_TREE_SPECIES_REF
                if (!oAdo.TableExist(con, m_strFiaTreeSpeciesRefTable))
                {
                    oDao.CreateTableLink(m_strTempMDBFileNewProcess, m_strFiaTreeSpeciesRefTable,
                        frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                        frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceDbFile, m_strFiaTreeSpeciesRefTable);
                }

                // FVS_TREE_SPECIES
                if (!oAdo.TableExist(con, Tables.Reference.DefaultFVSTreeSpeciesTableName))
                {
                    oDao.CreateTableLink(m_strTempMDBFileNewProcess, Tables.Reference.DefaultFVSTreeSpeciesTableName,
                        frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                        frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceDbFile, Tables.Reference.DefaultFVSTreeSpeciesTableName);
                }

                // REF_FOREST_TYPE and REF_FOREST_TYPE_GROUP
                string strBioSumRefDb = frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                        frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceSqliteFile;
                if (!oAdo.TableExist(con, m_strRefForestTypeTable))
                {
                    oDao.CreateSQLiteTableLink(m_strTempMDBFileNewProcess, m_strRefForestTypeTable,
                        m_strRefForestTypeTable, ODBCMgr.DSN_KEYS.BiosumRefDsnName, strBioSumRefDb);
                }
                if (!oAdo.TableExist(con, m_strRefForestTypeGroupTable))
                {
                    oDao.CreateSQLiteTableLink(m_strTempMDBFileNewProcess, m_strRefForestTypeGroupTable,
                        m_strRefForestTypeGroupTable, ODBCMgr.DSN_KEYS.BiosumRefDsnName, strBioSumRefDb);
                }

                // DWM tables
                string strMasterAuxDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\master_aux.db";
                if (!oAdo.TableExist(con, m_strDwmCoarseWoodyDebrisTable))
                {
                    oDao.CreateSQLiteTableLink(m_strTempMDBFileNewProcess, m_strDwmCoarseWoodyDebrisTable, m_strDwmCoarseWoodyDebrisTable,
                        ODBCMgr.DSN_KEYS.DWMDsnName, strMasterAuxDb);
                }
                if (!oAdo.TableExist(con, m_strDwmDuffLitterTable))
                {
                    oDao.CreateSQLiteTableLink(m_strTempMDBFileNewProcess, m_strDwmDuffLitterTable, m_strDwmDuffLitterTable,
                        ODBCMgr.DSN_KEYS.DWMDsnName, strMasterAuxDb);
                }
                if (!oAdo.TableExist(con, m_strDwmFineWoodyDebrisTable))
                {
                    oDao.CreateSQLiteTableLink(m_strTempMDBFileNewProcess, m_strDwmFineWoodyDebrisTable, m_strDwmFineWoodyDebrisTable,
                        ODBCMgr.DSN_KEYS.DWMDsnName, strMasterAuxDb);
                }
                if (!oAdo.TableExist(con, m_strDwmTransectSegmentTable))
                {
                    oDao.CreateSQLiteTableLink(m_strTempMDBFileNewProcess, m_strDwmTransectSegmentTable, m_strDwmTransectSegmentTable,
                        ODBCMgr.DSN_KEYS.DWMDsnName, strMasterAuxDb);
                }


            }
        }

        public void CreateFVSInputTables(System.Data.SQLite.SQLiteConnection p_Conn)
        {
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();

            string strSql = "ATTACH DATABASE '" + m_strSourceFiaDb + "' AS source";
            DebugLogSQL(strSql);
            oDataMgr.SqlNonQuery(p_Conn, strSql);

            // Query schema and create empty stand table
            strSql = "SELECT sql FROM source.sqlite_master WHERE type = 'table' " +
                "AND name = '" + Tables.FIA2FVS.DefaultFvsInputStandTableName + "'";
            DebugLogSQL(strSql);
            strSql = oDataMgr.getSingleStringValueFromSQLQuery(p_Conn, strSql, "sqlite_master");
            if (!String.IsNullOrEmpty(strSql))
            {
                DebugLogSQL(strSql);
                oDataMgr.SqlNonQuery(p_Conn, strSql);
            }

            // Add BioSum DWM fields to empty stand table
            string[] arrDwmFields = new string[] {"SmallMediumTotalLength", "LargeTotalLength", "CWDTotalLength",
                                                        "DuffPitCount", "LitterPitCount"};
            string[] arrDwmTypes = new string[] { "DOUBLE", "DOUBLE", "DOUBLE", "LONG", "LONG" };
            for (int i = 0; i < arrDwmFields.Length; i++)
            {
                oDataMgr.AddColumn(p_Conn, Tables.FIA2FVS.DefaultFvsInputStandTableName, arrDwmFields[i], arrDwmTypes[i], null);
                DebugLogSQL("Added column " + arrDwmFields[i] + " to table " + Tables.FIA2FVS.DefaultFvsInputStandTableName);
            }

            // Create empty tree table
            strSql = "SELECT sql FROM source.sqlite_master WHERE type = 'table' " +
                "AND name = '" + Tables.FIA2FVS.DefaultFvsInputTreeTableName + "'";
            DebugLogSQL(strSql);
            strSql = oDataMgr.getSingleStringValueFromSQLQuery(p_Conn, strSql, "sqlite_master");
            if (!String.IsNullOrEmpty(strSql))
            {
                DebugLogSQL(strSql);
                oDataMgr.SqlNonQuery(p_Conn, strSql);
            }

            // Add BiosumGeneratedSiteIndex_YN column to stand table
            oDataMgr.AddColumn(p_Conn, Tables.FIA2FVS.DefaultFvsInputStandTableName, "BioSumGeneratedSiteIndex_YN",
                "CHAR", "1");

            // Add VARIANT field to empty tree table
            oDataMgr.AddColumn(p_Conn, Tables.FIA2FVS.DefaultFvsInputTreeTableName, "VARIANT", "CHAR", "2");

            // Disconnect from source database
            strSql = "DETACH DATABASE 'source'";
            DebugLogSQL(strSql);
            oDataMgr.SqlNonQuery(p_Conn, strSql);
        }

        public void AddFIADBIndexes(System.Data.SQLite.SQLiteConnection p_Conn)
        {
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();

            string[] arrStandIndexNames = new string[] { "FVS_STANDINIT_COND_STAND_CN_IDX", "FVS_STANDINIT_COND_STAND_ID_IDX" };
            string[] arrStandIndexColumns = new string[] { "STAND_CN", "STAND_ID" };
            for (int i = 0; i < arrStandIndexColumns.Length; i++)
            {
                if (!oDataMgr.IndexExist(p_Conn, arrStandIndexNames[i]) &&
                     oDataMgr.FieldExist(p_Conn, $@"SELECT * FROM {Tables.FIA2FVS.DefaultFvsInputStandTableName}", arrStandIndexColumns[i]))
                {
                    oDataMgr.AddIndex(p_Conn, Tables.FIA2FVS.DefaultFvsInputStandTableName, arrStandIndexNames[i], arrStandIndexColumns[i]);
                }
            }
            string[] arrTreeIndexNames = new string[] { "FVS_TREEINIT_COND_STAND_CN_IDX", "FVS_TREEINIT_COND_STAND_ID_IDX", "FVS_TREEINIT_COND_STANDPLOT_CN_IDX", "FVS_TREEINIT_COND_STANDPLOT_ID_IDX", "FVS_TREEINIT_COND_TREE_CN_IDX", "FVS_TREEINIT_TREE_ID_IDX" };
            string[] arrTreeIndexColumns = new string[] { "STAND_CN", "STAND_ID", "STANDPLOT_CN", "STANDPLOT_ID", "TREE_CN", "TREE_ID" };
            for (int i = 0; i < arrTreeIndexColumns.Length; i++)
            {
                if (!oDataMgr.IndexExist(p_Conn, arrTreeIndexNames[i]) &&
                     oDataMgr.FieldExist(p_Conn, $@"SELECT * FROM {Tables.FIA2FVS.DefaultFvsInputTreeTableName}", arrTreeIndexColumns[i]))
                {
                    oDataMgr.AddIndex(p_Conn, Tables.FIA2FVS.DefaultFvsInputTreeTableName, arrTreeIndexNames[i], arrTreeIndexColumns[i]);
                }
            }
        }

        private void UpdateDamageCodesForCullAndMistletoe(string strTreeTable, string strTempMDBFile, string strTempTreeTable)
        {
            ado_data_access oAdo = new ado_data_access();
            string[] arrQueries = Queries.FVS.FVSInput.TreeInit.UpdateDamageCodesForCull(strTreeTable, strTempTreeTable);
            string strConn = oAdo.getMDBConnString(strTempMDBFile, "", "");
            using (OleDbConnection oAccessConn = new OleDbConnection(strConn))
            {
                oAccessConn.Open();
                for (int i = 0; i < arrQueries.Length; i++)
                {
                    DebugLogSQL(arrQueries[i]);
                    oAdo.SqlNonQuery(oAccessConn, arrQueries[i]);

                }

                // Update treevalue codes; They are based on damage codes
                arrQueries = Queries.FVS.FVSInput.TreeInit.UpdateTreeValue(strTempTreeTable);
                for (int i = 0; i < arrQueries.Length; i++)
                {
                    DebugLogSQL(arrQueries[i]);
                    oAdo.SqlNonQuery(oAccessConn, arrQueries[i]);
                }
            }
        }
        private void UpdateDamageCodesForCullAndMistletoeNew(SQLite.ADO.DataMgr oDataMgr, System.Data.SQLite.SQLiteConnection conn,
            string strTreeTable, string strTempTreeTable)
        {
            string[] arrQueries = Queries.FVS.FVSInput.TreeInit.UpdateDamageCodesForCullNew(strTreeTable, strTempTreeTable);

            for (int i = 0; i < arrQueries.Length; i++)
            {
                DebugLogSQL(arrQueries[i]);
                oDataMgr.SqlNonQuery(conn, arrQueries[i]);
            }

            // Update treevalue codes; They are based on damage codes
            arrQueries = Queries.FVS.FVSInput.TreeInit.UpdateTreeValue(strTempTreeTable);
            for (int i = 0; i < arrQueries.Length; i++)
            {
                DebugLogSQL(arrQueries[i]);
                oDataMgr.SqlNonQuery(conn, arrQueries[i]);
            }
        }

    }		

	
}

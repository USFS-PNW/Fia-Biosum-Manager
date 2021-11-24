using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_scenario_merch_chip_value.
	/// </summary>
	public class uc_processor_scenario_merch_chip_value : System.Windows.Forms.UserControl
	{
		public int m_intError=0;
		public string m_strError="";
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Label lblSpcGrp;
		private System.Windows.Forms.Label lblDbhClass;
		private System.Windows.Forms.Label lblMerchValue;
        private System.Windows.Forms.Panel pnlMerchValues;
        private FIA_Biosum_Manager.uc_processor_scenario_spc_dbh_group_value_collection uc_processor_scenario_spc_dbh_group_value_collection1 = new uc_processor_scenario_spc_dbh_group_value_collection();
		private System.Windows.Forms.Label lblChipValue;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtChipValue;
		private RxTools m_oRxTools = new RxTools();
		private ado_data_access m_oAdo = null;
		private string _strScenarioId="";
		private frmProcessorScenario _frmProcessorScenario=null;
		public FIA_Biosum_Manager.ResizeFormUsingVisibleScrollBars m_oResizeForm = new ResizeFormUsingVisibleScrollBars();
		private FIA_Biosum_Manager.ValidateNumericValues m_oValidate = new ValidateNumericValues();
		private string m_strChipValueSave="";
        private Label lblChips;
        private uc_processor_scenario_spc_dbh_group_value uc_processor_scenario_spc_dbh_group_value1;
		
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_processor_scenario_merch_chip_value()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			m_oResizeForm.ScrollBarParentControl=panel1;
			m_oResizeForm.ResizeWidth=true;
			m_oResizeForm.ResizeHeight=true;
			m_oResizeForm.MaximumHeight = 650;
			this.m_oValidate.MinValue=0;
			this.m_oValidate.Money=true;
			this.m_oValidate.RoundDecimalLength=2;
			this.m_oValidate.NullsAllowed=false;
			this.m_oValidate.TestForMax=false;
			this.m_oValidate.TestForMin=true;

			// TODO: Add any initialization after the InitializeComponent call

		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}
		public void loadvalues()
		{
			int x;
            //
            //SCENARIO ID
            //
            ScenarioId = this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();

            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                if (m_oAdo != null && m_oAdo.m_OleDbConnection != null)
                    if (m_oAdo.m_OleDbConnection.State == System.Data.ConnectionState.Open) m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);

                //
                //SCENARIO MDB
                //
                string strScenarioMDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\db\\scenario_processor_rule_definitions.mdb";

                //
                //CREATE LINK IN TEMP MDB TO SCENARIO TREE SPECIES DIAM DOLLAR VALUES TABLE
                //
                dao_data_access oDao = new dao_data_access();
                //link to tree species groups table
                oDao.CreateTableLink(this.ReferenceProcessorScenarioForm.LoadedQueries.m_strTempDbFile,
                                     "scenario_tree_species_diam_dollar_values",
                                     strScenarioMDB, "scenario_tree_species_diam_dollar_values", true);
                oDao.m_DaoWorkspace.Close();
                oDao = null;
                //
                //OPEN CONNECTION TO TEMP DB FILE
                //
                m_oAdo = new ado_data_access();
                m_oAdo.OpenConnection(m_oAdo.getMDBConnString(this.ReferenceProcessorScenarioForm.LoadedQueries.m_strTempDbFile, "", ""));
                if (m_oAdo.m_intError == 0)
                {
                    ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadSpeciesAndDiameterGroupDollarValues
                        (m_oAdo, m_oAdo.m_OleDbConnection, ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);
                }
            }
            else
            {
                //
                //SCENARIO MDB
                //
                string strScenarioMDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadSpeciesAndDiameterGroupDollarValuesSqlite(strScenarioMDB,
                    ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);
            }


                //REMOVE OLD CONTROLS FROM FORM IF THEY EXIST
                string strName = "uc_processor_scenario_spc_dbh_group_value2";
                if (this.pnlMerchValues.Controls[strName] != null)
                {
                    for (x = 2; x <= this.uc_processor_scenario_spc_dbh_group_value_collection1.Count; x++)
                    {
                        strName = "uc_processor_scenario_spc_dbh_group_value" + x;
                        uc_processor_scenario_spc_dbh_group_value oItem = (uc_processor_scenario_spc_dbh_group_value) this.pnlMerchValues.Controls[strName];
                        if (oItem != null)
                        {
                            this.pnlMerchValues.Controls.Remove(oItem);
                        }
                    }
                }

                //REMOVE OLD ITEMS FROM COLLECTION IF THEY EXIST
                if (this.uc_processor_scenario_spc_dbh_group_value_collection1.Count > 0)
                {
                    this.uc_processor_scenario_spc_dbh_group_value_collection1.Clear();
                }
                
                for (x = 0; x <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Count - 1; x++)
                {
                    if (this.uc_processor_scenario_spc_dbh_group_value_collection1.Count == 0)
                    {
                        if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim().Length > 0)
                        {
                          
                            //
                            //SPECIES GROUP
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.SpeciesGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim();
                            //
                            //DBH GROUP
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.DbhGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).DbhGroup.Trim();
                            //
                            //MERCH CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchDollarPerCubicFootValue);
                            this.uc_processor_scenario_spc_dbh_group_value1.CubicFootDollarValue = this.m_oValidate.ReturnValue;
                            //
                            //USE AS ENERGY WOOD
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.EnergyWood =
                                  ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).UseAsEnergyWood;

                            this.uc_processor_scenario_spc_dbh_group_value1.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;

                            this.uc_processor_scenario_spc_dbh_group_value_collection1.Add(this.uc_processor_scenario_spc_dbh_group_value1);
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).ChipsDollarPerCubicFootValue);
                            this.txtChipValue.Text = this.m_oValidate.ReturnValue;
                            this.m_strChipValueSave = this.txtChipValue.Text;
                        }
                    }
                    else
                    {
                        if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim().Length > 0)
                        {
                            uc_processor_scenario_spc_dbh_group_value oItem = new uc_processor_scenario_spc_dbh_group_value();
                            //
                            //SPECIES GROUP
                            //
                            oItem.SpeciesGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim();
                            //
                            //DBH GROUP
                            //
                            oItem.DbhGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).DbhGroup.Trim();
                            //
                            //MERCH CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchDollarPerCubicFootValue);
                            oItem.CubicFootDollarValue = this.m_oValidate.ReturnValue;
                            //
                            //USE AS ENERGY WOOD
                            oItem.EnergyWood =
                                  ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).UseAsEnergyWood;
                            oItem.Name = "uc_processor_scenario_spc_dbh_group_value" + Convert.ToString(uc_processor_scenario_spc_dbh_group_value_collection1.Count + 1).Trim();
                            this.pnlMerchValues.Controls.Add(oItem);
                            oItem.Top = uc_processor_scenario_spc_dbh_group_value_collection1.Item(uc_processor_scenario_spc_dbh_group_value_collection1.Count - 1).Top +
                                uc_processor_scenario_spc_dbh_group_value_collection1.Item(uc_processor_scenario_spc_dbh_group_value_collection1.Count - 1).Height;
                            oItem.Left = uc_processor_scenario_spc_dbh_group_value_collection1.Item(uc_processor_scenario_spc_dbh_group_value_collection1.Count - 1).Left;
                            oItem.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
                            oItem.Visible = true;
                            this.uc_processor_scenario_spc_dbh_group_value_collection1.Add(oItem);
                        }
                    }
                }
		}

        public void loadvalues_FromProperties()
        {
            int x;

            //REMOVE OLD CONTROLS FROM FORM IF THEY EXIST
            string strName = "uc_processor_scenario_spc_dbh_group_value2";
            if (this.pnlMerchValues.Controls[strName] != null)
            {
                for (x = 2; x <= this.uc_processor_scenario_spc_dbh_group_value_collection1.Count; x++)
                {
                    strName = "uc_processor_scenario_spc_dbh_group_value" + x;
                    uc_processor_scenario_spc_dbh_group_value oItem = (uc_processor_scenario_spc_dbh_group_value)this.pnlMerchValues.Controls[strName];
                    if (oItem != null)
                    {
                        this.pnlMerchValues.Controls.Remove(oItem);
                    }
                }
            }

            //REMOVE OLD ITEMS FROM COLLECTION IF THEY EXIST
            if (this.uc_processor_scenario_spc_dbh_group_value_collection1.Count > 0)
            {
                this.uc_processor_scenario_spc_dbh_group_value_collection1.Clear();
            }

            if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection != null)
            {
                for (x = 0; x <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Count - 1; x++)
                {
                    if (this.uc_processor_scenario_spc_dbh_group_value_collection1.Count == 0)
                    {
                        if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim().Length > 0)
                        {

                            //
                            //SPECIES GROUP
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.SpeciesGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim();
                            //
                            //DBH GROUP
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.DbhGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).DbhGroup.Trim();
                            //
                            //MERCH CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchDollarPerCubicFootValue);
                            this.uc_processor_scenario_spc_dbh_group_value1.CubicFootDollarValue = this.m_oValidate.ReturnValue;
                            //
                            //USE AS ENERGY WOOD
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.EnergyWood =
                                  ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).UseAsEnergyWood;

                            this.uc_processor_scenario_spc_dbh_group_value1.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;

                            this.uc_processor_scenario_spc_dbh_group_value_collection1.Add(this.uc_processor_scenario_spc_dbh_group_value1);
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).ChipsDollarPerCubicFootValue);
                            this.txtChipValue.Text = this.m_oValidate.ReturnValue;
                            this.m_strChipValueSave = this.txtChipValue.Text;
                        }
                    }
                    else
                    {
                        if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim().Length > 0)
                        {
                            uc_processor_scenario_spc_dbh_group_value oItem = new uc_processor_scenario_spc_dbh_group_value();
                            //
                            //SPECIES GROUP
                            //
                            oItem.SpeciesGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim();
                            //
                            //DBH GROUP
                            //
                            oItem.DbhGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).DbhGroup.Trim();
                            //
                            //MERCH CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchDollarPerCubicFootValue);
                            oItem.CubicFootDollarValue = this.m_oValidate.ReturnValue;
                            //
                            //USE AS ENERGY WOOD
                            //
                            oItem.EnergyWood =
                                  ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).UseAsEnergyWood;
                            oItem.Name = "uc_processor_scenario_spc_dbh_group_value" + Convert.ToString(uc_processor_scenario_spc_dbh_group_value_collection1.Count + 1).Trim();
                            this.pnlMerchValues.Controls.Add(oItem);
                            oItem.Top = uc_processor_scenario_spc_dbh_group_value_collection1.Item(uc_processor_scenario_spc_dbh_group_value_collection1.Count - 1).Top +
                                uc_processor_scenario_spc_dbh_group_value_collection1.Item(uc_processor_scenario_spc_dbh_group_value_collection1.Count - 1).Height;
                            oItem.Left = uc_processor_scenario_spc_dbh_group_value_collection1.Item(uc_processor_scenario_spc_dbh_group_value_collection1.Count - 1).Left;
                            oItem.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
                            oItem.Visible = true;
                            this.uc_processor_scenario_spc_dbh_group_value_collection1.Add(oItem);
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).ChipsDollarPerCubicFootValue);
                            this.txtChipValue.Text = this.m_oValidate.ReturnValue;
                            this.m_strChipValueSave = this.txtChipValue.Text;
                        }
                    }
                }
            }
        }

		public void savevalues()
		{
			int x;
			string strSpcGrp="";
			string strDbhGrp="";
			string strMerchValue="";
			string strChipValue="";
            string strWoodBin = "";
            string strSql = "";

            SQLite.ADO.DataMgr oDataMgr = null;
            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                //
                //DELETE THE CURRENT SCENARIO RECORDS
                //
                m_oAdo.m_strSQL = "DELETE FROM scenario_tree_species_diam_dollar_values " +
                                "WHERE TRIM(scenario_id)='" + this.ScenarioId.Trim() + "'";
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                //
                //DELETE THE WORK TABLE
                //
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "spcgrp_dbhgrp"))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE spcgrp_dbhgrp");
                //
                //CREATE AND POPULATE WORK TABLE
                //
                m_oAdo.m_strSQL = "CREATE TABLE spcgrp_dbhgrp (" +
                        "species_group INTEGER," +
                        "species_label CHAR(50)," +
                        "diam_group INTEGER," +
                        "diam_class CHAR(15))";
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

                foreach (ProcessorScenarioItem.SpcGroupItem objSpcGroup in ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupItem_Collection)
                {
                    foreach (ProcessorScenarioItem.TreeDiamGroupsItem objDiamGroup in ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection)
                    {
                        // INITIALIZE RECORDS IN WORK TABLE
                        m_oAdo.m_strSQL = "INSERT INTO spcgrp_dbhgrp (species_group,species_label, diam_group, diam_class) " +
                            "VALUES (" + objSpcGroup.SpeciesGroup + ",'" + objSpcGroup.SpeciesGroupLabel + "'," +
                            objDiamGroup.DiamGroup + ",'" + objDiamGroup.DiamClass + "')";
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                        //
                        //INSERT SCENARIO RECORDS
                        //
                        m_oAdo.m_strSQL = "INSERT INTO scenario_tree_species_diam_dollar_values (scenario_id,species_group,diam_group) " +
                                        "VALUES ('" + ScenarioId.Trim() + "'," + objSpcGroup.SpeciesGroup + "," +
                                        objDiamGroup.DiamGroup + ")";
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                    }
                }
            }
            else
            {
                oDataMgr = new SQLite.ADO.DataMgr();
                string strScenarioDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                oDataMgr.OpenConnection(oDataMgr.GetConnectionString(strScenarioDB));
                if (oDataMgr.m_intError != 0)
                {
                    m_intError = oDataMgr.m_intError;
                    m_strError = oDataMgr.m_strError;
                    oDataMgr = null;
                    return;
                }
                m_intError = 0;
                m_strError = "";
                //
                //DELETE THE CURRENT SCENARIO RECORDS
                //
                oDataMgr.m_strSQL = "DELETE FROM scenario_tree_species_diam_dollar_values " +
                                "WHERE TRIM(scenario_id)='" + this.ScenarioId.Trim() + "'";
                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                //
                //DELETE THE WORK TABLE
                //
                if (oDataMgr.TableExist(oDataMgr.m_Connection, "spcgrp_dbhgrp"))
                    oDataMgr.SqlNonQuery(oDataMgr.m_Connection, "DROP TABLE spcgrp_dbhgrp");
                //
                //CREATE AND POPULATE WORK TABLE
                //
                oDataMgr.m_strSQL = "CREATE TABLE spcgrp_dbhgrp (" +
                        "species_group INTEGER," +
                        "species_label TEXT," +
                        "diam_group INTEGER," +
                        "diam_class TEXT)";
                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);

                foreach (ProcessorScenarioItem.SpcGroupItem objSpcGroup in ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupItem_Collection)
                {
                    foreach (ProcessorScenarioItem.TreeDiamGroupsItem objDiamGroup in ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection)
                    {
                        // INITIALIZE RECORDS IN WORK TABLE
                        oDataMgr.m_strSQL = "INSERT INTO spcgrp_dbhgrp (species_group,species_label, diam_group, diam_class) " +
                            "VALUES (" + objSpcGroup.SpeciesGroup + ",'" + objSpcGroup.SpeciesGroupLabel + "'," +
                            objDiamGroup.DiamGroup + ",'" + objDiamGroup.DiamClass + "')";
                        oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                        //
                        //INSERT SCENARIO RECORDS
                        //
                        oDataMgr.m_strSQL = "INSERT INTO scenario_tree_species_diam_dollar_values (scenario_id,species_group,diam_group) " +
                                        "VALUES ('" + ScenarioId.Trim() + "'," + objSpcGroup.SpeciesGroup + "," +
                                        objDiamGroup.DiamGroup + ")";
                        oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                    }
                }
            }


			//
			//UPDATE SCENARIO RECORDS WITH MERCH AND CHIP VALUES
			//
			for (x=0;x<=this.uc_processor_scenario_spc_dbh_group_value_collection1.Count-1;x++)
			{
				strSpcGrp = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).SpeciesGroup.Trim();
				strDbhGrp = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).DbhGroup.Trim();
                strWoodBin = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).GetWoodBin();
				strMerchValue = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).CubicFootDollarValue.Trim();
				strMerchValue = strMerchValue.Replace("$","");
				strChipValue = this.txtChipValue.Text.Trim();
				strChipValue = strChipValue.Replace("$","");

                if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
                {
                    strSql = "UPDATE scenario_tree_species_diam_dollar_values a " +
                                  "INNER JOIN spcgrp_dbhgrp b " +
                                  "ON  a.species_group=b.species_group AND " +
                                      "a.diam_group=b.diam_group " +
                                  "SET a.merch_value=" + strMerchValue + "," +
                                      "a.chip_value=" + strChipValue + ", " +
                                      "a.wood_bin='" + strWoodBin.Trim() + "' " +
                                  "WHERE TRIM(a.scenario_id)='" + ScenarioId.Trim() + "' AND " +
                                        "TRIM(b.species_label)='" + strSpcGrp + "' AND " +
                                        "TRIM(b.diam_class)='" + strDbhGrp + "'";
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, strSql);
                }
                else
                {
                    strSql = "UPDATE scenario_tree_species_diam_dollar_values " +
                             "SET merch_value = " + strMerchValue + "," +
                             "chip_value = " + strChipValue + "," +
                             "wood_bin = '" + strWoodBin.Trim() + "' WHERE EXISTS (" +
                             "SELECT * FROM spcgrp_dbhgrp " +
                             "WHERE scenario_tree_species_diam_dollar_values.species_group = spcgrp_dbhgrp.species_group " +
                             "AND scenario_tree_species_diam_dollar_values.diam_group = spcgrp_dbhgrp.diam_group " +
                             "AND TRIM(scenario_tree_species_diam_dollar_values.scenario_id) = '" + ScenarioId.Trim() + "' " +
                             "AND TRIM(spcgrp_dbhgrp.species_label) = '" + strSpcGrp + "' " +
                             "AND TRIM(spcgrp_dbhgrp.diam_class) = '" + strDbhGrp + "')";

                    oDataMgr.SqlNonQuery(oDataMgr.m_Connection, strSql);
                }
				
				uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).SaveValues();							
			}
			this.m_strChipValueSave= this.txtChipValue.Text;
            if (ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                //
                //DELETE THE WORK TABLE AND CLOSE CONNECTION
                //
                if (oDataMgr.TableExist(oDataMgr.m_Connection, "spcgrp_dbhgrp"))
                    oDataMgr.SqlNonQuery(oDataMgr.m_Connection, "DROP TABLE spcgrp_dbhgrp");
                    m_intError = oDataMgr.m_intError;

                    oDataMgr.CloseConnection(oDataMgr.m_Connection);
                    oDataMgr = null;
            }
			
		}
		public frmProcessorScenario ReferenceProcessorScenarioForm
		{
			get {return this._frmProcessorScenario;}
			set {this._frmProcessorScenario=value;}
		}
		public string ScenarioId
		{
			get {return _strScenarioId;}
			set {_strScenarioId=value;}
		}
        public uc_processor_scenario_spc_dbh_group_value_collection ReferenceUserControlMarketValueSpeciesDbhGroupCollection
        {
            get { return uc_processor_scenario_spc_dbh_group_value_collection1; }
        }
        public string MarketValueChips
        {
            get { return txtChipValue.Text.Trim(); }
        }


		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblChips = new System.Windows.Forms.Label();
            this.txtChipValue = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lblChipValue = new System.Windows.Forms.Label();
            this.pnlMerchValues = new System.Windows.Forms.Panel();
            this.uc_processor_scenario_spc_dbh_group_value1 = new FIA_Biosum_Manager.uc_processor_scenario_spc_dbh_group_value();
            this.lblMerchValue = new System.Windows.Forms.Label();
            this.lblDbhClass = new System.Windows.Forms.Label();
            this.lblSpcGrp = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.pnlMerchValues.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(696, 504);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.lblChips);
            this.panel1.Controls.Add(this.txtChipValue);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.lblChipValue);
            this.panel1.Controls.Add(this.pnlMerchValues);
            this.panel1.Controls.Add(this.lblMerchValue);
            this.panel1.Controls.Add(this.lblDbhClass);
            this.panel1.Controls.Add(this.lblSpcGrp);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 48);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(690, 453);
            this.panel1.TabIndex = 30;
            // 
            // lblChips
            // 
            this.lblChips.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblChips.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblChips.Location = new System.Drawing.Point(310, 54);
            this.lblChips.Name = "lblChips";
            this.lblChips.Size = new System.Drawing.Size(66, 65);
            this.lblChips.TabIndex = 9;
            this.lblChips.Text = "Allocate to Energy Wood";
            this.lblChips.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // txtChipValue
            // 
            this.txtChipValue.Location = new System.Drawing.Point(546, 103);
            this.txtChipValue.Name = "txtChipValue";
            this.txtChipValue.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtChipValue.Size = new System.Drawing.Size(72, 20);
            this.txtChipValue.TabIndex = 8;
            this.txtChipValue.Text = "$0.00";
            this.txtChipValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtChipValue.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtChipValue_KeyPress);
            this.txtChipValue.Leave += new System.EventHandler(this.txtChipValue_Leave);
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label2.Location = new System.Drawing.Point(521, 49);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(148, 51);
            this.label2.TabIndex = 7;
            this.label2.Text = " Enter value for energy wood (chips) in $/green ton";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label1.Location = new System.Drawing.Point(15, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(424, 32);
            this.label1.TabIndex = 6;
            this.label1.Text = " Enter Average value in $/cubic foot for each species group and diameter class fo" +
    "r merchantable wood";
            // 
            // lblChipValue
            // 
            this.lblChipValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblChipValue.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblChipValue.Location = new System.Drawing.Point(484, 16);
            this.lblChipValue.Name = "lblChipValue";
            this.lblChipValue.Size = new System.Drawing.Size(203, 24);
            this.lblChipValue.TabIndex = 5;
            this.lblChipValue.Text = "Chip (Energy) Wood Values";
            // 
            // pnlMerchValues
            // 
            this.pnlMerchValues.AutoScroll = true;
            this.pnlMerchValues.BackColor = System.Drawing.SystemColors.Control;
            this.pnlMerchValues.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnlMerchValues.Controls.Add(this.uc_processor_scenario_spc_dbh_group_value1);
            this.pnlMerchValues.Location = new System.Drawing.Point(8, 122);
            this.pnlMerchValues.Name = "pnlMerchValues";
            this.pnlMerchValues.Size = new System.Drawing.Size(500, 317);
            this.pnlMerchValues.TabIndex = 3;
            // 
            // uc_processor_scenario_spc_dbh_group_value1
            // 
            this.uc_processor_scenario_spc_dbh_group_value1.CubicFootDollarValue = "$0.00";
            this.uc_processor_scenario_spc_dbh_group_value1.DbhGroup = "";
            this.uc_processor_scenario_spc_dbh_group_value1.EnergyWood = false;
            this.uc_processor_scenario_spc_dbh_group_value1.Location = new System.Drawing.Point(-2, 3);
            this.uc_processor_scenario_spc_dbh_group_value1.Name = "uc_processor_scenario_spc_dbh_group_value1";
            this.uc_processor_scenario_spc_dbh_group_value1.ReferenceProcessorScenarioForm = null;
            this.uc_processor_scenario_spc_dbh_group_value1.Size = new System.Drawing.Size(478, 34);
            this.uc_processor_scenario_spc_dbh_group_value1.SpeciesGroup = "";
            this.uc_processor_scenario_spc_dbh_group_value1.TabIndex = 0;
            // 
            // lblMerchValue
            // 
            this.lblMerchValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMerchValue.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblMerchValue.Location = new System.Drawing.Point(382, 79);
            this.lblMerchValue.Name = "lblMerchValue";
            this.lblMerchValue.Size = new System.Drawing.Size(104, 40);
            this.lblMerchValue.TabIndex = 2;
            this.lblMerchValue.Text = "Merchantable Value in $/ft3";
            this.lblMerchValue.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // lblDbhClass
            // 
            this.lblDbhClass.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDbhClass.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblDbhClass.Location = new System.Drawing.Point(216, 85);
            this.lblDbhClass.Name = "lblDbhClass";
            this.lblDbhClass.Size = new System.Drawing.Size(88, 34);
            this.lblDbhClass.TabIndex = 1;
            this.lblDbhClass.Text = "Tree DBH Class";
            this.lblDbhClass.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // lblSpcGrp
            // 
            this.lblSpcGrp.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSpcGrp.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblSpcGrp.Location = new System.Drawing.Point(15, 95);
            this.lblSpcGrp.Name = "lblSpcGrp";
            this.lblSpcGrp.Size = new System.Drawing.Size(104, 24);
            this.lblSpcGrp.TabIndex = 0;
            this.lblSpcGrp.Text = "Species Group";
            this.lblSpcGrp.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(690, 32);
            this.lblTitle.TabIndex = 29;
            this.lblTitle.Text = "Values Assumed for Delivered Wood at the Mill or Processing Site gate";
            // 
            // uc_processor_scenario_merch_chip_value
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_scenario_merch_chip_value";
            this.Size = new System.Drawing.Size(696, 504);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.pnlMerchValues.ResumeLayout(false);
            this.ResumeLayout(false);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;

        }
		#endregion

		private void label2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtChipValue_Leave(object sender, System.EventArgs e)
		{
			m_oValidate.Money=true;
			m_oValidate.RoundDecimalLength=2;
			m_oValidate.NullsAllowed=false;
			m_oValidate.TestForMax=false;
			m_oValidate.TestForMin=true;
			m_oValidate.MinValue=0;
			m_oValidate.ValidateDecimal(this.txtChipValue.Text.Trim());
            if (m_oValidate.m_intError == 0)
            {
                
                this.txtChipValue.Text = m_oValidate.ReturnValue;
                
            }
            else
                this.txtChipValue.Focus();

		}

        private void txtChipValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }
	}
}

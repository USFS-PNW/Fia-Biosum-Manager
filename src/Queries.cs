using System;
using System.Windows.Forms;
using System.Collections.Generic;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for Tables.
	/// </summary>
	public class Queries
	{
        public int m_intError=0;
		public string m_strError="";
		public Project m_oProject = new Project();
		public CoreScenarioResults m_oCoreScenarioResults = new CoreScenarioResults();
		public CoreScenarioRuleDefinitions m_oCoreScenarioRuleDef = new CoreScenarioRuleDefinitions();
        public Queries.CoreScenarioRun m_oCoreAnalysisScenarioRun = new CoreScenarioRun();
		public FIAPlot m_oFIAPlot = new FIAPlot();
		public FVS m_oFvs = new FVS();
		public TravelTime m_oTravelTime = new TravelTime();
		public Processor m_oProcessor = new Processor();
        public ProcessorScenarioRun m_oProcessorScenarioRun = new ProcessorScenarioRun();
		public Audit m_oAudit = new Audit();
		public Reference m_oReference = new Reference();
		public FIA_Biosum_Manager.Datasource m_oDataSource;
        public VolumeAndBiomass m_oVolumeAndBiomass = new VolumeAndBiomass();
		public string m_strTempDbFile;
		private bool _bScenario=false;
		private string _strScenarioType="";

		
		public Queries()
		{
			m_oFvs.ReferenceQueries=this;
			m_oFIAPlot.ReferenceQueries=this;
			m_oReference.ReferenceQueries=this;
            m_oProcessor.ReferenceQueries = this;
            m_oTravelTime.ReferenceQueries = this;
			//
			// TODO: Add constructor logic here
			//
		}
		public bool Scenario
		{
			get {return _bScenario;}
			set {_bScenario=value;}
		}
		public string ScenarioType
		{
			get {return _strScenarioType;}
			set {_strScenarioType=value;}
		}
		/// <summary>
		/// use the DataSource class to get DB files and table names.
		/// </summary>
		/// <param name="p_bLimited">TRUE=do not open and load the table record numbers, table column names, and table column data types</param>
		public void LoadDatasources(bool  p_bLimited)
		{
			if (p_bLimited)
			{
				LoadLimitedDatasources();
				
			}
			else
			{
				
			}
			if (this.m_oFvs.LoadDatasource) this.m_oFvs.LoadDatasources();
			if (this.m_oFIAPlot.LoadDatasource) this.m_oFIAPlot.LoadDatasources();
			if (this.m_oReference.LoadDatasource) this.m_oReference.LoadDatasources();
            if (this.m_oProcessor.LoadDatasource) this.m_oProcessor.LoadDatasources();
            if (this.m_oTravelTime.LoadDatasource) this.m_oTravelTime.LoadDatasources();
			m_strTempDbFile = this.m_oDataSource.CreateMDBAndTableDataSourceLinks();
		}
		public void LoadDatasources(bool p_bLimited, bool p_bUsingSqlite, string p_strScenarioType, string p_strScenarioId)
		{
			Scenario=true;
			ScenarioType=p_strScenarioType;
			if (p_bLimited)
			{
				LoadLimitedDatasources(p_strScenarioType, p_bUsingSqlite, p_strScenarioId);
			}
            if (this.m_oDataSource.m_intError < 0)
            {
                // An error has occurred in LoadLimitedDatasources likely due to dao 'too many client tasks'
                // The error originates in populate_datasource_array()
                MessageBox.Show("An error occurred while loading data sources! Close the current window " +
                                "and try again. If the problem persists, close and restart FIA Biosum Manager.",
                                "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
			}
			if (this.m_oFvs.LoadDatasource) this.m_oFvs.LoadDatasources();
			if (this.m_oFIAPlot.LoadDatasource) this.m_oFIAPlot.LoadDatasources();
			if (this.m_oReference.LoadDatasource) this.m_oReference.LoadDatasources();
            if (this.m_oProcessor.LoadDatasource) this.m_oProcessor.LoadDatasources();
			m_strTempDbFile = this.m_oDataSource.CreateMDBAndTableDataSourceLinks();
		}

		protected void LoadLimitedDatasources()
		{
			string strProjDir=frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
			
			m_oDataSource = new Datasource();
			m_oDataSource.LoadTableColumnNamesAndDataTypes=false;
			m_oDataSource.LoadTableRecordCount=false;
			m_oDataSource.m_strDataSourceMDBFile = strProjDir.Trim() + "\\db\\project.mdb";
			m_oDataSource.m_strDataSourceTableName = "datasource";
			m_oDataSource.m_strScenarioId="";
			m_oDataSource.populate_datasource_array();
			
			
		}
		protected void LoadLimitedDatasources(string p_strScenarioType, bool p_bUsingSqlite, string p_strScenarioId)
		{
			string strProjDir=frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
			
			m_oDataSource = new Datasource();
			m_oDataSource.LoadTableColumnNamesAndDataTypes=false;
			m_oDataSource.m_strScenarioId=p_strScenarioId.Trim();
			m_oDataSource.LoadTableRecordCount=false;
            m_oDataSource.m_strDataSourceTableName = "scenario_datasource";
            if (!p_bUsingSqlite)
            {
                m_oDataSource.m_strDataSourceMDBFile = strProjDir.Trim() + "\\" + p_strScenarioType + "\\db\\scenario_" + p_strScenarioType + "_rule_definitions.mdb";
                m_oDataSource.populate_datasource_array();
            }
            else
            {
                m_oDataSource.m_strDataSourceMDBFile = strProjDir.Trim() + "\\" + p_strScenarioType + "\\db\\scenario_" + p_strScenarioType + "_rule_definitions.db";
                m_oDataSource.populate_datasource_array_sqlite();
            }
        }
		static public string GetInsertSQL(string p_strFields, string p_strValues,string p_strTable)
		{
			return "INSERT INTO " + p_strTable + " (" + p_strFields + ") VALUES (" + p_strValues + ")";
		}
		public static string GenericSelectSQL(string p_strTable,string p_strFields,string p_strWhereExpression)
		{
			return "SELECT " + p_strFields + " FROM " + p_strTable + " WHERE " + p_strWhereExpression;
		}
		public class Project
		{
			
		}
		public class CoreScenarioResults
		{
			private string strSQL = "";
			public CoreScenarioResults()
			{
			}
		}
		public class CoreScenarioRuleDefinitions
		{
			private string strSQL = "";
			public CoreScenarioRuleDefinitions()
			{
			}
		}
        public class CoreScenarioRun
        {
            private string strSQL = "";
            private string _strScenarioId = "";
            public string ScenarioId
            {
                get { return _strScenarioId; }
                set { _strScenarioId = value; }
            }
            public CoreScenarioRun()
            {
            }
        }
		public class FVS
		{
			public string m_strRxTable;
			public string m_strFvsCmdTable;
			public string m_strFvsCatTable;
			public string m_strFvsSubCatTable;
			public string m_strRxFvsCmdTable;
			public string m_strRxHarvestCostColumnsTable;
			public string m_strRxPackageTable;
			public string m_strRxPackageMembersTable;
			public string m_strRxPackageFvsCmdTable;
			public string m_strRxPackageFvsCmdOrderTable;
			public string m_strTreeSpcTable;
			public string m_strFvsTreeTable;
			public string m_strFvsTreeSpcRefTable;
            public string m_strFVSWesternTreeSpeciesTable;
            public string m_strFVSEasternTreeSpeciesTable;
            public string m_strFVSPrePostSeqNumTable;
            public string m_strFVSPrePostSeqNumRxPackageAssgnTable;
			private Queries _oQueries=null;	
			private bool _bLoadDataSources=true;
			string m_strSql="";

			public FVS()
			{
			}
			public Queries ReferenceQueries
			{
				get {return _oQueries;}
				set {_oQueries=value;}
			}
			public bool LoadDatasource
			{
				get {return _bLoadDataSources;}
				set {_bLoadDataSources=value;}
			}


			public void LoadDatasources()
			{
				m_strRxTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PRESCRIPTIONS");
				m_strRxFvsCmdTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PRESCRIPTIONS ASSIGNED FVS COMMANDS");
				m_strRxHarvestCostColumnsTable=ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PRESCRIPTIONS HARVEST COST COLUMNS");
				m_strFvsCmdTable  = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("FVS COMMANDS");
                m_strFVSWesternTreeSpeciesTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("FVS WESTERN TREE SPECIES TRANSLATOR");
                m_strFVSEasternTreeSpeciesTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("FVS EASTERN TREE SPECIES TRANSLATOR");
				m_strFvsCatTable =  ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PRESCRIPTION CATEGORIES");
				m_strFvsSubCatTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PRESCRIPTION SUBCATEGORIES");
				m_strRxPackageTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PACKAGES");
				m_strRxPackageMembersTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PACKAGE MEMBERS");
				m_strRxPackageFvsCmdTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PACKAGE ASSIGNED FVS COMMANDS");
				m_strRxPackageFvsCmdOrderTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PACKAGE FVS COMMANDS ORDER");
				m_strTreeSpcTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREE SPECIES");
				m_strFvsTreeTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("FVS TREE LIST FOR PROCESSOR");
				m_strFvsTreeSpcRefTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName(Datasource.TableTypes.FvsTreeSpecies.ToUpper());
                m_strFVSPrePostSeqNumTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("FVS PRE-POST SEQNUM DEFINITIONS");
                m_strFVSPrePostSeqNumRxPackageAssgnTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("FVS PRE-POST SEQNUM TREATMENT PACKAGE ASSIGNMENTS");
				
			
				if (this.m_strRxTable.Trim().Length == 0) 
				{
					
					MessageBox.Show("!!Could Not Locate Rx Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strFvsCatTable.Trim().Length == 0 && !ReferenceQueries.Scenario)
				{
					MessageBox.Show("!!Could Not Locate FVS Category Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strFvsSubCatTable.Trim().Length == 0 && !ReferenceQueries.Scenario)
				{
					MessageBox.Show("!!Could Not Locate FVS Subcategory Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strFvsCmdTable.Trim().Length == 0 && !ReferenceQueries.Scenario)
				{
					MessageBox.Show("!!Could Not Locate FVS Command Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strRxFvsCmdTable.Trim().Length == 0 && !ReferenceQueries.Scenario)
				{
					MessageBox.Show("!!Could Not Locate Rx FVS Command Assignments Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strRxHarvestCostColumnsTable.Trim().Length == 0 && ReferenceQueries._strScenarioType!="optimizer")
				{
					MessageBox.Show("!!Could Not Locate Rx Harvest Cost Columns Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}

                if (this.m_strTreeSpcTable.Trim().Length == 0 && ReferenceQueries._strScenarioType != "optimizer")
				{
					MessageBox.Show("!!Could Not Locate Tree Species Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}

				if (this.m_strFvsTreeSpcRefTable.Trim().Length == 0 && !ReferenceQueries.Scenario)
				{
					MessageBox.Show("!!Could Not Locate FVS Tree Species Reference Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
			}
			
			public string GetRxItemCategoryDescriptionSQL(string p_strRxId)
			{
				return "SELECT desc " + 
					   "FROM " + m_strFvsCatTable + " " + 
                       "WHERE MAX >= " + p_strRxId + " AND " + 
					   "MIN <= " + p_strRxId;
			}
			public string GetRxItemSubCategoryDescriptionSQL(string p_strRxId)
			{
				return "SELECT desc " + 
					   "FROM " + m_strFvsSubCatTable + " " + 
					   "WHERE MAX >= " + p_strRxId + " AND " + 
					   "MIN <= " + p_strRxId;
			}
			public string GetCategoryIdFromDescriptionSQL(string p_strDesc)
			{
				return "SELECT DISTINCT CSTR(catid) " + 
					   "FROM " + m_strFvsCatTable + " c " + 
					   "WHERE TRIM(UCASE(c.DESC)) = '" + p_strDesc.Trim().ToUpper() + "'";
			}
			public string GetSubCategoryIdFromDescriptionSQL(string p_strDesc)
			{
				return "SELECT DISTINCT CSTR(subcatid) " + 
					"FROM " + m_strFvsSubCatTable + " c " + 
					"WHERE TRIM(UCASE(c.DESC)) = '" + p_strDesc.Trim().ToUpper() + "'";
			}
			public string GetCategoryDescriptionFromCategoryIdSQL(string p_strCatId)
			{
				return "SELECT DISTINCT desc " + 
					"FROM " + m_strFvsCatTable + " c " + 
					"WHERE c.catid = " + p_strCatId;
			}
			public string GetSubCategoryDescriptionFromCategoryIdAndSubCategoryIdSQL(string p_strCatId,string p_strSubCatId)
			{
				return "SELECT DISTINCT desc " + 
					"FROM " + this.m_strFvsSubCatTable + " c " + 
					"WHERE c.catid = " + p_strCatId + " AND " + 
						  "c.subcatid = " + p_strSubCatId;
			}
			
			public string GetFVSCommandPropertiesSQL(string p_strCmd)
			{
				return "SELECT * " + 
					   "FROM " + m_strFvsCmdTable + " " + 
					   "WHERE TRIM(UCASE(fvscmd))='" + p_strCmd.Trim().ToUpper() + "'";
			}
			public string GetInsertRxSQL(string p_strFields, string p_strValues)
			{
				return "INSERT INTO " + this.m_strRxTable + " (" + p_strFields + ") VALUES (" + p_strValues + ")";
			}
			public string GetInsertRxFvsCmdSQL(string p_strFields, string p_strValues)
			{
				return "INSERT INTO " + this.m_strRxTable + " (" + p_strFields + ") VALUES (" + p_strValues + ")";
			}

			/// <summary>
			/// get the subcategories that are child members of the parent category.
			/// </summary>
			/// <param name="p_strCat"></param>
			/// <returns></returns>
			public string GetFVSSubCategoriesSQL(string p_strCat)
			{
				return "SELECT s.desc, s.min, s.max " + 
					   "FROM " + this.m_strFvsSubCatTable + " s," +
					             this.m_strFvsCatTable + " c " + 
					   "WHERE TRIM(UCASE(c.desc))='" + p_strCat.Trim().ToUpper() + "' AND " + 
					         "s.catid=c.catid";
			}
			/// <summary>
			/// return the query that will assign a variant to each rx package
			/// </summary>
			/// <param name="p_strPlotTable"></param>
			/// <param name="p_strRxPackageTable"></param>
			/// <returns></returns>
			static public string GetFVSVariantRxPackageSQL(string p_strPlotTable, string p_strRxPackageTable)
			{
				return "SELECT DISTINCT  a.fvs_variant,  b.rxpackage, b.rxcycle_length, b.simyear1_rx," + 
					                                                                   "b.simyear2_rx," + 
					                                                                   "b.simyear3_rx," + 
					                                                                   "b.simyear4_rx  " + 
					   "FROM " + p_strPlotTable + " a, " + 
					      "(SELECT rxpackage,simyear1_rx,simyear2_rx,simyear3_rx,simyear4_rx,rxcycle_length " + 
					       "FROM " + p_strRxPackageTable +  ") b " + 
					   "WHERE a.fvs_variant IS NOT NULL AND " + 
					          "LEN(TRIM(a.fvs_variant)) > 0 AND " + 
					          "b.rxpackage IS NOT NULL AND LEN(TRIM(b.rxpackage)) > 0;";
			}
			/// <summary>
			/// return the query that will assign a variant to each rx package that is in the list
			/// </summary>
			/// <param name="p_strPlotTable"></param>
			/// <param name="p_strRxPackageTable"></param>
			/// <param name="p_strRxPackageList"></param>
			/// <param name="p_strFilter">include a filter using alias (a) for plot table and alias (b) for rxpackage table</param>
			/// <returns></returns>
			static public string GetFVSVariantRxPackageSQL(string p_strPlotTable, string p_strRxPackageTable,string p_strFilter)
			{
				return "SELECT DISTINCT  a.fvs_variant,  b.rxpackage, b.rxcycle_length, b.simyear1_rx," + 
																					   "b.simyear2_rx," + 
																					   "b.simyear3_rx," + 
																					   "b.simyear4_rx  " + 
					"FROM " + p_strPlotTable + " a, " + 
						"(SELECT rxpackage,simyear1_rx,simyear2_rx,simyear3_rx,simyear4_rx,rxcycle_length " + 
						 "FROM " + p_strRxPackageTable +  ") b " + 
					"WHERE a.fvs_variant IS NOT NULL AND " + 
						 "LEN(TRIM(a.fvs_variant)) > 0 AND " + 
						"b.rxpackage IS NOT NULL AND LEN(TRIM(b.rxpackage)) > 0 AND " + 
					    p_strFilter;
			}
            /// <summary>
            ///  Assign a sequence number to each record of the FVS output table and group by standid,year
            /// </summary>
            /// <param name="p_strIntoTable"></param>
            /// <param name="p_strFVSOutputTable"></param>
            /// <param name="p_bAllColumns"></param>
            /// <returns></returns>
            static public string FVSOutputTable_PrePostGenericSQL(string p_strIntoTable, string p_strFVSOutputTable, bool p_bAllColumns)
            {
                string strSQL = "";
                if (p_bAllColumns)
                {
                    strSQL = "SELECT d.SeqNum,a.* " +
                              "FROM " + p_strFVSOutputTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year " +
                                  "FROM " + p_strFVSOutputTable + " b," +
                                        "(SELECT standid,year " +
                                         "FROM " + p_strFVSOutputTable + ") c " +
                                 "WHERE b.standid=c.standid " +
                                 "GROUP BY b.standid,b.year) d " +
                              "WHERE a.standid=d.standid AND a.year=d.year";
                }
                else
                {
                    if (p_strIntoTable.Trim().Length > 0)
                    {
                        strSQL = "SELECT  SUM(IIF(a.year >= b.year,1,0)) AS SeqNum," +
                                          "a.standid, a.year " +
                                 "INTO " + p_strIntoTable + " " + 
                                 "FROM " + p_strFVSOutputTable + " a," +
                                      "(SELECT standid,year " +
                                       "FROM " + p_strFVSOutputTable + ") b " +
                                 "WHERE a.standid=b.standid " +
                                 "GROUP BY a.standid,a.year";
                    }
                    else
                    {
                        strSQL = "SELECT  SUM(IIF(a.year >= b.year,1,0)) AS SeqNum," +
                                          "a.standid, a.year " +
                                 "FROM " + p_strFVSOutputTable + " a," +
                                      "(SELECT standid,year " +
                                       "FROM " + p_strFVSOutputTable + ") b " +
                                 "WHERE a.standid=b.standid " +
                                 "GROUP BY a.standid,a.year";
                    }
                }
                return strSQL;
               
            }
            static public string[] FVSOutputTable_PrePostPotFireBaseYearSQL(string p_strPotFireBaseYearTable, string p_strPotFireTable,string p_strWorkTableName)
            {
                string[] strSQL = new string[9];

                //create a baseyear table that contains only standid's that are in the standard table
                strSQL[0] = "SELECT DISTINCT a.* " +
                            "INTO tempBASEYEAR " +
                            "FROM " + p_strPotFireBaseYearTable + " a " +
                            "INNER JOIN " + p_strPotFireTable + " b ON a.standid=b.standid";

                //get the potfire base year records into baseyear temp table
                strSQL[1] = "SELECT 'Y' AS BASEYEAR_YN,a.* " +
                          "INTO BASEYEAR " +
                          "FROM tempBASEYEAR a," +
                            "(SELECT STANDID, MIN([YEAR]) AS BASEYEAR, 'Y' AS BASEYEAR_YN " +
                             "FROM tempBASEYEAR " +
                            "GROUP BY STANDID) b " +
                         "WHERE a.standid = b.standid AND a.year=b.baseyear";

                //get fvs_potfire  records into nonbaseyear temp table and increment the year by 1
                strSQL[2] = "SELECT 'N' AS BASEYEAR_YN,(a.[YEAR] + 1) AS [NEWYEAR], a.* " +
                            "INTO NONBASEYEAR " +
                            "FROM " + p_strPotFireTable + " a ";

                //update the year column to the newyear from the previous step
                strSQL[3] = "UPDATE NONBASEYEAR SET [YEAR]=NEWYEAR";

                //drop the newyear column
                strSQL[4] = "ALTER TABLE NONBASEYEAR DROP COLUMN NEWYEAR";

                strSQL[5] = "SELECT * INTO " + p_strWorkTableName + " FROM BASEYEAR";

                strSQL[6] = "INSERT INTO " + p_strWorkTableName + " SELECT * FROM NONBASEYEAR";

                strSQL[7] = "DROP TABLE NONBASEYEAR";

                strSQL[8] = "DROP TABLE BASEYEAR";

                strSQL[8] = "DROP TABLE tempBASEYEAR";

                return strSQL;
            }
            static public string[] FVSOutputTable_PrePostPotFireBaseYearIDColumnSQL(string p_strPotFireBaseYearTable, string p_strPotFireTable, string p_strWorkTableName)
            {
                string[] strSQL = new string[12];

                //create a baseyear table that contains only standid's that are in the standard table
                strSQL[0] = "SELECT DISTINCT a.* " +
                            "INTO tempBASEYEAR " +
                            "FROM " + p_strPotFireBaseYearTable + " a " +
                            "INNER JOIN " + p_strPotFireTable + " b ON a.standid=b.standid";

                //get the potfire base year records into baseyear temp table
                strSQL[1] = "SELECT 'Y' AS BASEYEAR_YN,a.* " +
                          "INTO BASEYEAR " +
                          "FROM tempBASEYEAR a," +
                            "(SELECT STANDID, MIN([YEAR]) AS BASEYEAR, 'Y' AS BASEYEAR_YN " +
                             "FROM tempBASEYEAR " +
                            "GROUP BY STANDID) b " +
                         "WHERE a.standid = b.standid AND a.year=b.baseyear";

                //get fvs_potfire  records into nonbaseyear temp table and increment the year by 1
                strSQL[2] = "SELECT 'N' AS BASEYEAR_YN,(a.[YEAR] + 1) AS [NEWYEAR], a.* " +
                            "INTO NONBASEYEAR " +
                            "FROM " + p_strPotFireTable + " a ";

                //update the year column to the newyear from the previous step
                strSQL[3] = "UPDATE NONBASEYEAR SET [YEAR]=NEWYEAR";

                //drop the newyear column
                strSQL[4] = "ALTER TABLE NONBASEYEAR DROP COLUMN NEWYEAR";



                //get the maximum id value from the nonbaseyear table, create a rownumber for each row in the baseyear table,
                //and add the rownumber to the maxid to get a unique id
                strSQL[5] = "SELECT a.*,b.rownumber + c.maxid AS tempid INTO " + p_strWorkTableName + " FROM BASEYEAR a,";
                strSQL[5] = strSQL[5] +
                            "(" + Queries.Utilities.AssignRowNumberToEachRow("", "BASEYEAR", "STANDID", "ROWNUMBER") +
                            ") b," +
                            "(SELECT MAX(ID) AS maxid FROM NONBASEYEAR) c " +
                            "WHERE a.standid=b.standid";

                //update in the baseyear temp table with the new id
                strSQL[6] = "UPDATE " + p_strWorkTableName + " SET ID=TEMPID";


                strSQL[7] = "ALTER TABLE " + p_strWorkTableName + " DROP COLUMN TEMPID";

                strSQL[8] = "INSERT INTO " + p_strWorkTableName + " SELECT * FROM NONBASEYEAR";

              
                strSQL[9] = "DROP TABLE NONBASEYEAR";

                strSQL[10] = "DROP TABLE BASEYEAR";

                strSQL[11] = "DROP TABLE tempBASEYEAR";

                return strSQL;
            }
            /// <summary>
            /// Assign a sequence number to each record of the FVS output table (FVS_STRCLASS) and group by standid,year and removal code.
            /// </summary>
            /// <param name="p_strIntoTable"></param>
            /// <param name="p_strFVSOutputTable">FVS_STRCLASS table name</param>
            /// <param name="p_bAllColumns"></param>
            /// <param name="p_strRemovalCode">Values: 0 (before removal) or 1 (after removal)</param>
            /// <returns></returns>
            static public string FVSOutputTable_PrePostStrClassSQL(string p_strIntoTable,string p_strFVSOutputTable,bool p_bAllColumns,string p_strRemovalCode)
            {
                string strSQL = "";
                if (p_bAllColumns)
                {
                    strSQL = "SELECT d.seqnum, z.* FROM " + p_strFVSOutputTable + " z," +
                                    "(SELECT  SUM(IIF(a.year >= b.year AND " + 
                                                     "a.removal_code=" + p_strRemovalCode + ",1,0)) AS SeqNum," + 
                                              "a.standid, a.year,b.removal_code " + 
                                     "FROM fvs_strclass a," +
                                        "(SELECT standid,year,removal_code " + 
                                         "FROM " + p_strFVSOutputTable + " " + 
                                         "WHERE removal_code=" + p_strRemovalCode + ") b " +
                                    "WHERE a.standid=b.standid AND " + 
                                          "a.removal_code=b.removal_code AND " + 
                                          "b.removal_code=" + p_strRemovalCode + " " +
                                    "GROUP BY a.standid,a.year,b.removal_code) d " +
                             "WHERE z.standid=d.standid AND z.year=d.year AND z.removal_code=d.removal_code";
                }
                else
                {
                    if (p_strIntoTable.Trim().Length > 0)
                    {
                        strSQL = "SELECT  SUM(IIF(a.year >= b.year,1,0)) AS SeqNum," +
                                          "a.standid, a.year " +
                                 "INTO " + p_strIntoTable + " " +
                                 "FROM " + p_strFVSOutputTable + " a," +
                                      "(SELECT standid,year " +
                                       "FROM " + p_strFVSOutputTable + ") b " +
                                 "WHERE a.standid=b.standid " +
                                 "GROUP BY a.standid,a.year";
                    }
                    else
                    {
                        strSQL = "SELECT  SUM(IIF(a.year >= b.year,1,0)) AS SeqNum," +
                                          "a.standid, a.year " +
                                 "FROM " + p_strFVSOutputTable + " a," +
                                      "(SELECT standid,year " +
                                       "FROM " + p_strFVSOutputTable + ") b " +
                                 "WHERE a.standid=b.standid " +
                                 "GROUP BY a.standid,a.year";
                    }
                }
                return strSQL;

            }
            /// <summary>
            /// SQL for creating the sequence number configuration table from the FVS Output table.
            /// </summary>
            /// <param name="p_strIntoTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <param name="p_strFVSOutputTable"></param>
            /// <param name="p_bAllColumns"></param>
            /// <returns></returns>
            static public string FVSOutputTable_AuditPrePostGenericSQL(string p_strIntoTable, string p_strFVSOutputTable, bool p_bAllColumns)
            {
                string strSQL = "";
               
                    if (p_strIntoTable.Trim().Length > 0)
                    {
                        strSQL = "SELECT d.SeqNum,a.standid,a.year," + 
                                       "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," + 
                                       "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," + 
                                       "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                       "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " + 
                                 "INTO " + p_strIntoTable + " " + 
                                 "FROM " + p_strFVSOutputTable + " a," +
                                     "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                              "b.standid, b.year " +
                                      "FROM " + p_strFVSOutputTable + " b," +
                                            "(SELECT standid,year " +
                                             "FROM " + p_strFVSOutputTable + ") c " +
                                     "WHERE b.standid=c.standid " +
                                     "GROUP BY b.standid,b.year) d " +
                                  "WHERE a.standid=d.standid AND a.year=d.year";
                    }
                    else
                    {
                        strSQL = "SELECT d.SeqNum,a.standid,a.year," +
                                       "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                       "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                       "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                       "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " + 
                                 "FROM " + p_strFVSOutputTable + " a," +
                                    "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                             "b.standid, b.year " +
                                     "FROM " + p_strFVSOutputTable + " b," +
                                           "(SELECT standid,year " +
                                            "FROM " + p_strFVSOutputTable + ") c " +
                                    "WHERE b.standid=c.standid " +
                                    "GROUP BY b.standid,b.year) d " +
                                 "WHERE a.standid=d.standid AND a.year=d.year " + 
                                 "ORDER BY a.standid,d.SeqNum";
                    }
               
                    return strSQL;

            }
            /// <summary>
            /// SQL for creating the sequence number configuration table from the FVS Output table.
            /// </summary>
            /// <param name="p_strIntoTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <param name="p_strFVSOutputTable"></param>
            /// <param name="p_bAllColumns"></param>
            /// <returns></returns>
            static public string SqliteFVSOutputTable_AuditPrePostGenericSQL(string p_strIntoTable, string p_strFVSOutputTable, bool p_bAllColumns,
                string strRxPackage)
            {
                string strSQL = "";

                if (p_strIntoTable.Trim().Length > 0)
                {
                    strSQL = "SELECT d.SeqNum,a.standid,a.year," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN, " +
                                   "substr(e.RunTitle, 12, 3) AS RXPACKAGE " +
                             "INTO " + p_strIntoTable + " " +
                             "FROM FVS." + p_strFVSOutputTable + " a, FVS.FVS_CASES e," +
                                 "(SELECT  SUM(CASE WHEN b.year >= c.year THEN 1 ELSE 0 END) AS SeqNum," +
                                          "b.standid, b.year " +
                                  "FROM FVS." + p_strFVSOutputTable + " b," +
                                        "(SELECT standid,year " +
                                         "FROM FVS." + p_strFVSOutputTable + ") c " +
                                 "WHERE b.standid=c.standid " +
                                 "GROUP BY b.standid,b.year) d " +
                                 "WHERE a.CaseID = e.CaseID AND substr(e.RunTitle, 12, 3) = '" + strRxPackage + 
                                 "' AND a.standid=d.standid AND a.year=d.year ";
                }
                else
                {
                    strSQL = "SELECT d.SeqNum,a.standid,a.year," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN, " +
                                   "substr(e.RunTitle, 12, 3) AS RXPACKAGE " +
                             "FROM FVS." + p_strFVSOutputTable + " a, FVS.FVS_CASES e," +
                              //"(SELECT  SUM(CASE WHEN b.year >= c.year THEN 1 ELSE 0 END) AS SeqNum," +
                              //         "b.standid, b.year " +
                              // "FROM FVS." + p_strFVSOutputTable + " b," +
                              //       "(SELECT standid,year " +
                              //        "FROM FVS." + p_strFVSOutputTable + ") c " +
                              //"WHERE b.standid=c.standid " +
                              //"GROUP BY b.standid,b.year) d " +
                              "(SELECT SUM(CASE WHEN b.year >= c.year THEN 1 ELSE 0 END) AS SeqNum," +
                                "b.standid, b.year FROM FVS." + p_strFVSOutputTable + " b, FVS.FVS_CASES f," +
                                "(SELECT g.standid, g.year FROM FVS." + p_strFVSOutputTable + " g, FVS.FVS_CASES h " +
                                "WHERE g.CaseID = h.CaseID and substr(h.RunTitle, 12, 3) = '" + strRxPackage + "' ) c " +
                                "WHERE b.standid = c.standid and b.CaseID = f.CaseID " +
                                "and substr(f.RunTitle, 12, 3) = '" + strRxPackage + "' " +
                                "GROUP BY b.standid,b.year) d " +
                             "WHERE a.CaseID = e.CaseID AND substr(e.RunTitle, 12, 3) = '" + strRxPackage + 
                             "' AND a.standid=d.standid AND a.year=d.year " +
                             "ORDER BY a.standid,d.SeqNum";
                }

                return strSQL;

            }
            /// <summary>
            /// Update SQL for assigning the sequence number to the PRE or POST cycle.
            /// </summary>
            /// <param name="p_oItem"></param>
            /// <param name="p_strUpdateTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <returns></returns>
            static public string FVSOutputTable_AuditUpdatePrePostGenericSQL(FVSPrePostSeqNumItem p_oItem,string p_strUpdateTable)
            {
                string strSQL="";
                 if (p_oItem.RxCycle1PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle1PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE1_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle1PreSeqNum + ",'Y','N'),";

                }
                 if (p_oItem.RxCycle2PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle2PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE2_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle2PreSeqNum + ",'Y','N'),";
                }

                 if (p_oItem.RxCycle3PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle3PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE3_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle3PreSeqNum + ",'Y','N'),";
                }
                 if (p_oItem.RxCycle4PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle4PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE4_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle4PreSeqNum + ",'Y','N'),";
                }
                 if (p_oItem.RxCycle1PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle1PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE1_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle1PostSeqNum + ",'Y','N'),";

                }
                 if (p_oItem.RxCycle2PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle2PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE2_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle2PostSeqNum + ",'Y','N'),";
                }

                 if (p_oItem.RxCycle3PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle3PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE3_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle3PostSeqNum + ",'Y','N'),";
                }
                 if (p_oItem.RxCycle4PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle4PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE4_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle4PostSeqNum + ",'Y','N'),";
                }
                if (strSQL.Trim().Length > 0)
                {
                    strSQL = strSQL.Substring(0, strSQL.Length - 1);
                    strSQL = "UPDATE " + p_strUpdateTable + " " +
                                      "SET " + strSQL;
                }
                return strSQL;
            }
            /// <summary>
            /// Update SQL for assigning the sequence number to the PRE or POST cycle.
            /// </summary>
            /// <param name="p_oItem"></param>
            /// <param name="p_strUpdateTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <returns></returns>
            static public string SqliteFVSOutputTable_AuditUpdatePrePostGenericSQL(FVSPrePostSeqNumItem p_oItem, string p_strUpdateTable)
            {
                string strSQL = "";
                if (p_oItem.RxCycle1PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle1PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE1_PRE_YN=CASE WHEN SeqNum=" + p_oItem.RxCycle1PreSeqNum + " THEN 'Y' ELSE 'N' END,";
                }
                if (p_oItem.RxCycle2PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle2PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE2_PRE_YN=CASE WHEN SeqNum=" + p_oItem.RxCycle2PreSeqNum + " THEN 'Y' ELSE 'N' END,";
                }

                if (p_oItem.RxCycle3PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle3PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE3_PRE_YN=CASE WHEN SeqNum=" + p_oItem.RxCycle3PreSeqNum + " THEN 'Y' ELSE 'N' END,";
                }
                if (p_oItem.RxCycle4PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle4PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE4_PRE_YN=CASE WHEN SeqNum=" + p_oItem.RxCycle4PreSeqNum + " THEN 'Y' ELSE 'N' END,";
                }
                if (p_oItem.RxCycle1PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle1PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE1_POST_YN=CASE WHEN SeqNum=" + p_oItem.RxCycle1PostSeqNum + " THEN 'Y' ELSE 'N' END,";

                }
                if (p_oItem.RxCycle2PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle2PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE2_POST_YN=CASE WHEN SeqNum=" + p_oItem.RxCycle2PostSeqNum + " THEN 'Y' ELSE 'N' END,";
                }

                if (p_oItem.RxCycle3PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle3PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE3_POST_YN=CASE WHEN SeqNum=" + p_oItem.RxCycle3PostSeqNum + " THEN 'Y' ELSE 'N' END,";
                }
                if (p_oItem.RxCycle4PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle4PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + "CYCLE4_POST_YN=CASE WHEN SeqNum=" + p_oItem.RxCycle4PostSeqNum + " THEN 'Y' ELSE 'N' END,";
                }
                if (strSQL.Trim().Length > 0)
                {
                    strSQL = strSQL.Substring(0, strSQL.Length - 1);
                    strSQL = "UPDATE " + p_strUpdateTable + " " +
                                      "SET " + strSQL;
                }
                return strSQL;
            }
            /// <summary>
            /// Update SQL for assigning the sequence number to the PRE or POST cycle.
            /// </summary>
            /// <param name="p_oItem"></param>
            /// <param name="p_strUpdateTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <returns></returns>
            static public string FVSOutputTable_AuditUpdatePrePostStrClassSQL(FVSPrePostSeqNumItem p_oItem, string p_strUpdateTable)
            {
                string strSQL = "";
                if (p_oItem.RxCycle1PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle1PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle1PreStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE1_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle1PreSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE1_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle1PreSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }

                }
                if (p_oItem.RxCycle2PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle2PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle2PreStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE2_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle2PreSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE2_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle2PreSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }
                }

                if (p_oItem.RxCycle3PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle3PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle3PreStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE3_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle3PreSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE3_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle3PreSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }
                }
                if (p_oItem.RxCycle4PreSeqNum.Trim().Length > 0 && p_oItem.RxCycle4PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle4PreStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE4_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle4PreSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE4_PRE_YN=IIF(SeqNum=" + p_oItem.RxCycle4PreSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }
                }
                if (p_oItem.RxCycle1PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle1PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle1PostStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE1_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle1PostSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE1_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle1PostSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }

                }
                if (p_oItem.RxCycle2PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle2PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle2PostStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE2_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle2PostSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE2_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle2PostSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }
                }

                if (p_oItem.RxCycle3PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle3PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle3PostStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE3_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle3PostSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE3_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle3PostSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }
                }
                if (p_oItem.RxCycle4PostSeqNum.Trim().Length > 0 && p_oItem.RxCycle4PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    if (p_oItem.RxCycle4PostStrClassBeforeTreeRemovalYN == "N")
                    {
                        strSQL = strSQL + "CYCLE4_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle4PostSeqNum + " AND removal_code=1 ,'Y','N'),";
                    }
                    else
                    {
                        strSQL = strSQL + "CYCLE4_POST_YN=IIF(SeqNum=" + p_oItem.RxCycle4PostSeqNum + " AND removal_code=0 ,'Y','N'),";
                    }
                }
                if (strSQL.Trim().Length > 0)
                {
                    strSQL = strSQL.Substring(0, strSQL.Length - 1);
                    strSQL = "UPDATE " + p_strUpdateTable + " " +
                                      "SET " + strSQL;
                }
                return strSQL;
            }
            /// <summary>
            /// SQL for creating the sequence number configuration table from the FVS Output table FVS_STRCLASS.
            /// </summary>
            /// <param name="p_strIntoTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <param name="p_strFVSOutputTable"></param>
            /// <param name="p_bAllColumns"></param>
            /// <returns></returns>
            static public string FVSOutputTable_AuditPrePostStrClassSQL(string p_strIntoTable, string p_strFVSOutputTable, bool p_bAllColumns)
            {
                string strSQL = "";

                if (p_strIntoTable.Trim().Length > 0)
                {
                    strSQL = "SELECT d.SeqNum,a.standid,a.year,a.removal_code," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " +
                             "INTO " + p_strIntoTable + " " +
                             "FROM " + p_strFVSOutputTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year,b.removal_code " +
                                  "FROM " + p_strFVSOutputTable + " b," +
                                        "(SELECT standid,year,removal_code " +
                                         "FROM " + p_strFVSOutputTable + ") c " +
                                 "WHERE b.standid=c.standid AND b.removal_code=c.removal_code " +
                                 "GROUP BY b.standid,b.year,b.removal_code) d " +
                              "WHERE a.standid=d.standid AND a.year=d.year AND a.removal_code=d.removal_code";
                }
                else
                {
                    strSQL = "SELECT d.SeqNum,a.standid,a.year,a.removal_code," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " +
                             "FROM " + p_strFVSOutputTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year,b.removal_code " +
                                  "FROM " + p_strFVSOutputTable + " b," +
                                        "(SELECT standid,year,removal_code " +
                                         "FROM " + p_strFVSOutputTable + ") c " +
                                 "WHERE b.standid=c.standid AND b.removal_code=c.removal_code " +
                                 "GROUP BY b.standid,b.year,b.removal_code) d " +
                             "WHERE a.standid=d.standid AND a.year=d.year AND a.removal_code=d.removal_code";
                }

                return strSQL;

            }
            /// <summary>
            /// SQL for creating the sequence number configuration table from the FVS Output table FVS_SUMMARY.
            /// </summary>
            /// <param name="p_strIntoTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <param name="p_strFVSOutputTable"></param>
            /// <param name="p_bAllColumns"></param>
            /// <returns></returns>
            static public string[] FVSOutputTable_AuditPrePostFvsStrClassUsingFVSSummarySQL(string p_strIntoTable, string p_strFVSSummaryTable, bool p_bAllColumns)
            {
                string[] strSQL = new string[2];

                if (p_strIntoTable.Trim().Length > 0)
                {
                 strSQL[0] = "SELECT d.SeqNum,a.standid,a.year,0 AS removal_code," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " +
                             "INTO " + p_strIntoTable + " " +
                             "FROM " + p_strFVSSummaryTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year " +
                                  "FROM " + p_strFVSSummaryTable + " b," +
                                        "(SELECT standid,year " +
                                         "FROM " + p_strFVSSummaryTable + ") c " +
                                 "WHERE b.standid=c.standid " +
                                 "GROUP BY b.standid,b.year) d " +
                              "WHERE a.standid=d.standid AND a.year=d.year";

                 strSQL[1] = "INSERT INTO " + p_strIntoTable + " " +
                             "SELECT d.SeqNum,a.standid,a.year,1 AS removal_code," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " +
                             "FROM " + p_strFVSSummaryTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year " +
                                  "FROM " + p_strFVSSummaryTable + " b," +
                                        "(SELECT standid,year " +
                                         "FROM " + p_strFVSSummaryTable + ") c " +
                                 "WHERE b.standid=c.standid " +
                                 "GROUP BY b.standid,b.year) d " +
                             "WHERE a.standid=d.standid AND a.year=d.year";
                }
                else
                {
                    strSQL[0] = "SELECT d.SeqNum,a.standid,a.year,0 AS removal_code," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " +
                             "FROM " + p_strFVSSummaryTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year " +
                                  "FROM " + p_strFVSSummaryTable + " b," +
                                        "(SELECT standid,year " +
                                         "FROM " + p_strFVSSummaryTable + ") c " +
                                 "WHERE b.standid=c.standid " +
                                 "GROUP BY b.standid,b.year) d " +
                             "WHERE a.standid=d.standid AND a.year=d.year";

                    strSQL[1] = "SELECT d.SeqNum,a.standid,a.year,1 AS removal_code," +
                                   "'N' AS CYCLE1_PRE_YN,'N' AS CYCLE1_POST_YN," +
                                   "'N' AS CYCLE2_PRE_YN,'N' AS CYCLE2_POST_YN," +
                                   "'N' AS CYCLE3_PRE_YN,'N' AS CYCLE3_POST_YN," +
                                   "'N' AS CYCLE4_PRE_YN,'N' AS CYCLE4_POST_YN " +
                             "FROM " + p_strFVSSummaryTable + " a," +
                                 "(SELECT  SUM(IIF(b.year >= c.year,1,0)) AS SeqNum," +
                                          "b.standid, b.year " +
                                  "FROM " + p_strFVSSummaryTable + " b," +
                                        "(SELECT standid,year " +
                                         "FROM " + p_strFVSSummaryTable + ") c " +
                                 "WHERE b.standid=c.standid " +
                                 "GROUP BY b.standid,b.year) d " +
                             "WHERE a.standid=d.standid AND a.year=d.year";
                }

                return strSQL;

            }

            /// <summary>
            /// Audit to identify assigned sequence numbers that cannot be found in the Sequence Number Matrix table (FVSTableName_PREPOST_SEQNUM_MATRIX)
            /// </summary>
            /// <param name="p_oFVSPrePostSeqNumItem"></param>
            /// <param name="p_strIntoTable"></param>
            /// <param name="p_strSourceTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <returns></returns>
            static public string FVSOutputTable_AuditSelectIntoPrePostSeqNumCount(FVSPrePostSeqNumItem p_oFVSPrePostSeqNumItem,string p_strIntoTable,string p_strSourceTable)
            {
                string strSQL = "";
                int x;
                int z = 0;
               
                string strAlpha = "cdefghij";
                int intAlias = 0;
                string strSelectColumns = "a.standid,b.totalrows,";
                
                //cycle 1 seqnum
                if (p_oFVSPrePostSeqNumItem.RxCycle1PreSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle1PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle1PreSeqNum + ",1,0)) AS pre_cycle1rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".pre_cycle1rows,";
                    intAlias++;
                }
                if (p_oFVSPrePostSeqNumItem.RxCycle1PostSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle1PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle1PostSeqNum + ",1,0)) AS post_cycle1rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".post_cycle1rows,";
                    intAlias++;
                }
                //cycle 2 seqnum
                if (p_oFVSPrePostSeqNumItem.RxCycle2PreSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle2PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle2PreSeqNum + ",1,0)) AS pre_cycle2rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".pre_cycle2rows,";
                    intAlias++;
                }
                if (p_oFVSPrePostSeqNumItem.RxCycle2PostSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle2PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle2PostSeqNum + ",1,0)) AS post_cycle2rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".post_cycle2rows,";
                    intAlias++;
                }
                //cycle 3 seqnum
                if (p_oFVSPrePostSeqNumItem.RxCycle3PreSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle3PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle3PreSeqNum + ",1,0)) AS pre_cycle3rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".pre_cycle3rows,";
                    intAlias++;
                }
                if (p_oFVSPrePostSeqNumItem.RxCycle3PostSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle3PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle3PostSeqNum + ",1,0)) AS post_cycle3rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".post_cycle3rows,";
                    intAlias++;
                }
                //cycle 4 seqnum
                if (p_oFVSPrePostSeqNumItem.RxCycle4PreSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle4PreSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle4PreSeqNum + ",1,0)) AS pre_cycle4rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".pre_cycle4rows,";
                    intAlias++;
                }
                if (p_oFVSPrePostSeqNumItem.RxCycle4PostSeqNum.Trim().Length > 0 && p_oFVSPrePostSeqNumItem.RxCycle4PostSeqNum.Trim().ToUpper() != "NOT USED")
                {
                    strSQL = strSQL + " (SELECT " +
                                         "SUM(IIF(SeqNum=" + p_oFVSPrePostSeqNumItem.RxCycle4PostSeqNum + ",1,0)) AS post_cycle4rows," +
                                         "STANDID " +
                                       "FROM " + p_strSourceTable + " " +
                                       "GROUP BY STANDID) " + strAlpha.Substring(intAlias, 1) + ",";
                    strSelectColumns = strSelectColumns + strAlpha.Substring(intAlias, 1) + ".post_cycle4rows,";
                    intAlias++;
                }
                strSQL = strSQL.Substring(0, strSQL.Length - 1);
                strSelectColumns = strSelectColumns.Substring(0, strSelectColumns.Length - 1);

                strSQL = "SELECT DISTINCT " + strSelectColumns + " " +
                                 "INTO " + p_strIntoTable + " " +
                                 "FROM " + p_strSourceTable + " a," +
                                    "(SELECT COUNT(*) AS totalrows," +
                                            "STANDID " +
                                     "FROM " + p_strSourceTable + " " +
                                 "GROUP BY standid) b," +
                                 strSQL + " " +
                                 "WHERE a.standid=b.standid AND ";

                for (x = 0; x <= intAlias - 1; x++)
                {
                    strSQL = strSQL + "a.standid=" + strAlpha.Substring(x, 1) + ".standid AND ";
                }

                strSQL = strSQL.Substring(0, strSQL.Length - 5);
                return strSQL;
            }
            /// <summary>
            /// Load data that will show which fvs_summary SeqNum are associated with harvested tree records
            /// </summary>
            /// <param name="p_strIntoTable1">Table to display to user</param>
            /// <param name="p_strIntoTable2">Temporary table</param>
            /// <param name="p_strFVSOutputSummaryTable">FVS output summary table</param>
            /// <param name="p_strFVSOutputTreeTable">FVS output tree cut table</param>
            /// <returns></returns>
            static public string[] FVSOutputTable_TreeHarvestSeqNumSQL(string p_strIntoTable1,string p_strIntoTable2, string p_strFVSOutputSummaryTable, string p_strFVSOutputTreeTable)
            {
                string[] strSQL = new string[4];

                strSQL[0] = "SELECT  SUM(IIF(a.year >= b.year,1,0)) AS FVS_Summary_SeqNum," +
                                          "a.standid, a.year,0 AS treecount  " +
                                 "INTO " + p_strIntoTable1 + " " + 
                                 "FROM " + p_strFVSOutputSummaryTable + " a," +
                                      "(SELECT standid,year " +
                                       "FROM " + p_strFVSOutputSummaryTable + ") b " +
                                 "WHERE a.standid=b.standid " +
                                 "GROUP BY a.standid,a.year";

                strSQL[1] = "SELECT SUM(1) AS TREECOUNT, STANDID, YEAR INTO " + p_strIntoTable2 + " " +
                            "FROM " + p_strFVSOutputTreeTable + " " + 
                            "GROUP BY STANDID,YEAR";

                strSQL[2] = "UPDATE " + p_strIntoTable1 + " a " +
                            "INNER JOIN " + p_strIntoTable2 + " b " +
                            "ON a.standid=b.standid AND a.year=b.year " +
                            "SET a.treecount=b.treecount";

                strSQL[3] = "DROP TABLE " + p_strIntoTable2;

                return strSQL;

            }
            /// <summary>
            /// Load data that will show which fvs_summary SeqNum are associated with target table records
            /// </summary>
            /// <param name="p_strIntoTable1">Table to display to user</param>
            /// <param name="p_strIntoTable2">Temporary table</param>
            /// <param name="p_strFVSOutputSummaryTable">FVS output summary table</param>
            /// <param name="p_strFVSOutputTreeTable">FVS output table</param>
            /// <returns></returns>
            static public string[] FVSOutputTable_AssignSummarySeqNumSQL(string p_strIntoTable1,string p_strIntoTable2, string p_strFVSOutputSummaryTable, string p_strFVSOutputTable)
            {
                string[] strSQL = new string[4];

                strSQL[0] = "SELECT  SUM(IIF(a.year >= b.year,1,0)) AS FVS_Summary_SeqNum," +
                                          "a.standid, a.year,0 AS rowcount  " +
                                 "INTO " + p_strIntoTable1 + " " + 
                                 "FROM " + p_strFVSOutputSummaryTable + " a," +
                                      "(SELECT standid,year " +
                                       "FROM " + p_strFVSOutputSummaryTable + ") b " +
                                 "WHERE a.standid=b.standid " +
                                 "GROUP BY a.standid,a.year";

                strSQL[1] = "SELECT SUM(1) AS ROWCOUNT, STANDID, YEAR INTO " + p_strIntoTable2 + " " +
                            "FROM " + p_strFVSOutputTable + " " + 
                            "GROUP BY STANDID,YEAR";

                strSQL[2] = "UPDATE " + p_strIntoTable1 + " a " +
                            "INNER JOIN " + p_strIntoTable2 + " b " +
                            "ON a.standid=b.standid AND a.year=b.year " +
                            "SET a.rowcount=b.rowcount";

                strSQL[3] = "DROP TABLE " + p_strIntoTable2;

                return strSQL;

            }

            static public string[] FVSOutputTable_AuditFVSSummaryTableRowCountsSQL(string p_strIntoTable, string p_strSummaryAuditTable, string p_strFVSOutputTable, string p_strRowCountColumn)
            {
                string[] strSQL = new string[4];

                strSQL[0] =  "SELECT DISTINCT a.standid,a.year,b.rowcount " + 
                             "INTO " + p_strIntoTable + " " + 
                             "FROM " + p_strFVSOutputTable + " a," +
                               "(SELECT COUNT(*) as rowcount,standid,year " + 
                                "FROM " + p_strFVSOutputTable + " " + 
                                "GROUP BY standid,year) b " + 
                             "WHERE a.standid=b.standid AND a.year=b.year";

               strSQL[1] = "UPDATE " + p_strSummaryAuditTable + " a " +  
                           "INNER JOIN " + p_strIntoTable + " b " + 
                           "ON a.standid=b.standid AND " + 
	                          "a.year=b.year " + 
                           "SET a." + p_strRowCountColumn + "=b.rowcount";

              strSQL[2] = "UPDATE " + p_strSummaryAuditTable + " " + 
                          "SET " + p_strRowCountColumn + "=0 " + 
                          "WHERE " + p_strRowCountColumn + " IS NULL";

              strSQL[3] = "DROP TABLE " + p_strIntoTable;

              return strSQL;
            }
			static public string FVSOutputTable_AuditSelectIntoCyleYearSQL(string p_strIntoTable, string p_strFVSOutputTable, int p_intCycleLength)
			{
					return "SELECT DISTINCT a.standid, b.year1, 'Y' AS year1_exists_yn," + 
											"b.year1 + " + p_intCycleLength.ToString() +  " AS year2," + 
											"'N' AS year2_exists_yn," + 
											"b.year1 + (" + p_intCycleLength.ToString() +  " * 2) AS year3," + 
											"'N' AS year3_exists_yn," + 
											"b.year1 + (" + p_intCycleLength.ToString() +  " * 3) AS year4," + 
											"'N' AS year4_exists_yn," + 
											"b.year1 + (" + p_intCycleLength.ToString() +  " * 4) AS year5," + 
											"'N' AS year5_exists_yn" + " " + 
						   "INTO " + p_strIntoTable + " " + 
						   "FROM " + p_strFVSOutputTable + " a," + 
								"(SELECT MIN([YEAR]) AS year1, standid " + 
						         "FROM " + p_strFVSOutputTable + " " + 
						         "GROUP BY standid) b" + " " + 
						   "WHERE a.standid = b.standid AND a.[year]=b.year1";
			}
			static public string FVSOutputTable_AuditUpdateCyleYearSQL(string p_strUpdateTableName, string p_strFVSOutputTable,string p_strCycle)
			{
				return "UPDATE " + p_strUpdateTableName + " a " + 
					   "INNER JOIN " + p_strFVSOutputTable + " b " + 
					   "ON a.standid=b.standid AND a.year" + p_strCycle.Trim() + "= b.year " + 
					   "SET year" + p_strCycle.Trim() + "_exists_yn=IIF(b.year=a.year" + p_strCycle.Trim() + ",'Y','N')";
			}
			static public string FVSOutputTable_AuditInsertUpdateCycleYearSQL(string p_strInsertTable, string p_strSourceTable,string p_strCycle)
			{
				return "INSERT INTO " + p_strInsertTable + " " + 
					      "(standid,`year`, cycle) " + 
					      "SELECT standid,year" + p_strCycle + " AS `year`," + 
								p_strCycle + " AS cycle " + 
					      "FROM " + p_strSourceTable;
			}
			static public string FVSOutputTable_AuditSelectIntoRxPostCyleYearSQL(string p_strIntoTable, 
															   string p_strFVSOutputTable,
				                                               string p_strFVSOutputCycleYearTable, 
				                                               string p_strCycle)
			{

				switch (p_strCycle)
				{
				    case "1":
						return  "SELECT a.standid, MIN(a.year) AS post_year" + p_strCycle + " " + 
							"INTO "  + p_strIntoTable + " " + 
							"FROM  " + p_strFVSOutputTable + " a, " + 
							p_strFVSOutputCycleYearTable + " b " + 
							"WHERE a.standid=b.standid AND " + 
							"a.year > b.year1 GROUP BY a.standid";
					case "2":
						return  "SELECT a.standid, MIN(a.year) AS post_year" + p_strCycle + " " + 
							"INTO "  + p_strIntoTable + " " + 
							"FROM  " + p_strFVSOutputTable + " a, " + 
							p_strFVSOutputCycleYearTable + " b " + 
							"WHERE a.standid=b.standid AND " + 
							"a.year <> b.year1 AND a.year > b.year2 GROUP BY a.standid";
					case "3":
						return  "SELECT a.standid, MIN(a.year) AS post_year" + p_strCycle + " " + 
							"INTO "  + p_strIntoTable + " " + 
							"FROM  " + p_strFVSOutputTable + " a, " + 
							p_strFVSOutputCycleYearTable + " b " + 
							"WHERE a.standid=b.standid AND " + 
						    "a.year <> b.year1 AND a.year <> b.year2 AND a.year > b.year3 GROUP BY a.standid";
					case "4":
						return  "SELECT a.standid, MIN(a.year) AS post_year" + p_strCycle + " " + 
							"INTO "  + p_strIntoTable + " " + 
							"FROM  " + p_strFVSOutputTable + " a, " + 
							p_strFVSOutputCycleYearTable + " b " + 
							"WHERE a.standid=b.standid AND " + 
							"a.year <> b.year1 AND a.year <> b.year2 AND a.year <> b.year3 AND a.year > b.year4 GROUP BY a.standid";




				}
				return "";
			}
            /// <summary>
			/// Returns the SQL used to insert the RX pretreatment year for each of the 4 cycles 
			/// </summary>
			/// <param name="p_strInsertTable"></param>
			/// <param name="p_strFVSOutputCycleYearTable"></param>
			/// <returns></returns>
			static public string FVSOutputTable_AuditInsertRxPreCycleYearTableSQL(string p_strInsertTable, string p_strFVSOutputCycleYearTable)
			{
				return "INSERT INTO " + p_strInsertTable + " (standid,pre_year1,pre_year2,pre_year3,pre_year4) " + 
				          "SELECT a.standid," + 
				          "IIF(a.year1_exists_yn='Y',a.year1,-1) AS pre_year1," + 
				          "IIF(a.year2_exists_yn='Y',a.year2,-1) AS pre_year2," + 
				          "IIF(a.year3_exists_yn='Y',a.year3,-1) AS pre_year3," + 
				          "IIF(a.year4_exists_yn='Y',a.year4,-1) AS pre_year4 " + 
				          "FROM " + p_strFVSOutputCycleYearTable + " a";
			}

			static public string FVSOutputTable_AuditUpdateRxPostCyleYearSQL(string p_strSourceTable, 
				string p_strJoinTable,
				string p_strCycle)
			{

				switch (p_strCycle)
				{
					case "1":
						return "UPDATE " + p_strSourceTable + " a " +
							"INNER JOIN " + p_strJoinTable + " b " + 
							"ON a.standid=b.standid SET a.post_year1=b.post_year1";
					case "2":
						return  "UPDATE " + p_strSourceTable + " a " +
							"INNER JOIN " + p_strJoinTable + " b " + 
							"ON a.standid=b.standid SET a.post_year2=IIF(a.pre_year2=-1,-1,b.post_year2)";
					case "3":
						return   "UPDATE " + p_strSourceTable + " a " +
							"INNER JOIN " + p_strJoinTable + " b " + 
							"ON a.standid=b.standid SET a.post_year3=IIF(a.pre_year3=-1,-1,b.post_year3)";
					case "4":
						return  "UPDATE " + p_strSourceTable + " a " +
							"INNER JOIN " + p_strJoinTable + " b " + 
							"ON a.standid=b.standid SET a.post_year4=IIF(a.pre_year4=-1,-1,b.post_year4)";




				}
				return "";
			}
            /// <summary>
            /// Audit to find missing stand,year combinations.
            /// Define the SQL Queries that will identify standid,year rows in the tree list tables that
            /// are not found in the FVS_SUMMARY table.
            /// </summary>
            /// <param name="p_strIntoTempTreeListTableName"></param>
            /// <param name="p_strIntoTempSummaryTableName"></param>
            /// <param name="p_strIntoTempMissingRowsTableName"></param>
            /// <param name="p_strTreeListTableName"></param>
            /// <param name="p_strSummaryTableName"></param>
            /// <returns></returns>
			static public string[] FVSOutputTable_AuditSelectTreeListCyleYearExistInFVSSummaryTableSQL(
                string p_strIntoTempTreeListTableName,
                string p_strIntoTempSummaryTableName,
                string p_strIntoTempMissingRowsTableName,
                string p_strTreeListTableName,
                string p_strSummaryTableName)
			{
                string[] strSQL = new string[6];
                strSQL[0] = "SELECT COUNT(*) AS TREECOUNT, STANDID,YEAR " +
                            "INTO " + p_strIntoTempTreeListTableName + " " +
                            "FROM " + p_strTreeListTableName + " " +
                            "WHERE standid IS NOT NULL and year > 0 " +
                            "GROUP BY CASEID,STANDID,YEAR";

                strSQL[1] = "SELECT DISTINCT STANDID,YEAR " +
                            "INTO " + p_strIntoTempSummaryTableName + " " +
                            "FROM " + p_strSummaryTableName + " " +
                            "WHERE standid IS NOT NULL and year > 0 ";

                strSQL[2] = "SELECT a.STANDID,a.YEAR,a.TREECOUNT, " + 
                              "SUM(IIF(a.standid=b.standid AND a.year=b.year,1,0)) AS ROWCOUNT " + 
                            "INTO " + p_strIntoTempMissingRowsTableName + " " + 
                            "FROM " + p_strIntoTempTreeListTableName + " a," + 
                             "(SELECT DISTINCT STANDID,YEAR " + 
                              "FROM " + p_strIntoTempSummaryTableName + " " + 
                              "WHERE STANDID IS NOT NULL) b " + 
                            "WHERE a.standid=b.standid " + 
                            "GROUP BY a.standid,a.year,a.treecount";

                strSQL[3] = "DELETE FROM " + p_strIntoTempMissingRowsTableName + "  WHERE ROWCOUNT > 0";

                strSQL[4] = "DROP TABLE " + p_strIntoTempSummaryTableName;

                strSQL[5] = "DROP TABLE " + p_strIntoTempTreeListTableName;


                return strSQL;

               
			}
            /// <summary>
            /// View the PRE-POST records from the FVS Output table (FVS_STRCLASS) that will be retrieved 
            /// based on sequential number assignment filters and removal code filters
            /// </summary>
            /// <param name="p_strFVSOutTable">FVS_STRCLASS table name</param>
            /// <param name="p_strSeqNumMatrixTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <param name="p_oItem"></param>
            /// <returns></returns>
            static public string FVSOutputTable_StrClassPrePostSeqNumByCycle(string p_strFVSOutTable,string p_strSeqNumMatrixTable,FVSPrePostSeqNumItem p_oItem)
            {

                string strSQL="";
                string strPreRemovalCode = "";
                string strPostRemovalCode = "";
                for (int x=1;x<=4;x++)
                {
                    if (x == 1)
                    {
                        strPreRemovalCode = p_oItem.RxCycle1PreStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                        strPostRemovalCode = p_oItem.RxCycle1PostStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                    }
                    else if (x == 2)
                    {
                        strPreRemovalCode = p_oItem.RxCycle2PreStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                        strPostRemovalCode = p_oItem.RxCycle2PostStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                    }
                    else if (x == 3)
                    {
                        strPreRemovalCode = p_oItem.RxCycle3PreStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                        strPostRemovalCode = p_oItem.RxCycle3PostStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                    }
                    else if (x == 4)
                    {
                        strPreRemovalCode = p_oItem.RxCycle4PreStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                        strPostRemovalCode = p_oItem.RxCycle4PostStrClassBeforeTreeRemovalYN == "Y" ? "0" : "1";
                    }
                    strSQL = strSQL + "SELECT b.seqnum, b.cycle,b.type, a.* " +
                                      "FROM " + p_strFVSOutTable.Trim() + " AS a," + 
                                         "(SELECT seqnum,standid,year,'" + x.ToString().Trim() + "' AS cycle,'PRE' AS [type] " + 
                                          "FROM " + p_strSeqNumMatrixTable + " WHERE CYCLE" + x.ToString().Trim() + "_PRE_YN='Y')  AS b " + 
                                     "WHERE a.standid=b.standid AND a.year=b.year AND a.Removal_Code=" + strPreRemovalCode + " " + 
                                     "UNION " + 
                                     "SELECT b.seqnum, b.cycle,b.type, a.* " +
                                      "FROM " + p_strFVSOutTable + " AS a," +
                                         "(SELECT seqnum,standid,year,'" + x.ToString().Trim() + "' AS cycle,'POST' AS [type] " + 
                                          "FROM " + p_strSeqNumMatrixTable + " WHERE CYCLE" + x.ToString().Trim() + "_POST_YN='Y')  AS b " +
                                     "WHERE a.standid=b.standid AND a.year=b.year AND a.Removal_Code=" + strPostRemovalCode + " " + 
                                     "UNION ";

                }
                strSQL = strSQL.Substring(0, strSQL.Length - 6);
                return strSQL;
               


            }
            /// <summary>
            /// View the PRE-POST records from the FVS Output table that will be retrieved 
            /// based on sequential number assignment filters
            /// </summary>
            /// <param name="p_strFVSOutTable">FVS output table</param>
            /// <param name="p_strSeqNumMatrixTable">FVSTableName_PREPOST_SEQNUM_MATRIX</param>
            /// <returns></returns>
            static public string FVSOutputTable_GenericPrePostSeqNumByCycle(string p_strFVSOutTable, string p_strSeqNumMatrixTable)
            {

                string strSQL = "";

                for (int x = 1; x <= 4; x++)
                {
                    strSQL = strSQL + "SELECT b.seqnum, b.cycle,b.type, a.* " +
                                      "FROM " + p_strFVSOutTable.Trim() + " AS a," +
                                         "(SELECT seqnum,standid,year,'" + x.ToString().Trim() + "' AS cycle,'PRE' AS [type] " +
                                          "FROM " + p_strSeqNumMatrixTable + " WHERE CYCLE" + x.ToString().Trim() + "_PRE_YN='Y')  AS b " +
                                     "WHERE a.standid=b.standid AND a.year=b.year " +
                                     "UNION " +
                                     "SELECT b.seqnum, b.cycle,b.type, a.* " +
                                      "FROM " + p_strFVSOutTable + " AS a," +
                                         "(SELECT seqnum,standid,year,'" + x.ToString().Trim() + "' AS cycle,'POST' AS [type] " +
                                          "FROM " + p_strSeqNumMatrixTable + " WHERE CYCLE" + x.ToString().Trim() + "_POST_YN='Y')  AS b " +
                                     "WHERE a.standid=b.standid AND a.year=b.year " +
                                     "UNION ";

                }
                strSQL = strSQL.Substring(0, strSQL.Length - 6);
                return strSQL;
            }
            
			/// <summary>
			/// make sure the FVSOut tables pre_year1,pre_year2,pre_year3,and pre_year4 cycle values are the same in the
			/// FVS Summary table
			/// </summary>
			/// <param name="p_strFVSOutCycleYearTable"></param>
			/// <param name="p_strSummaryCycleYearTable"></param>
			/// <returns></returns>
			static public string FVSOutputTable_AuditSelectFVSOutCycleYearInFVSSummaryCycleYearTableSQL(string p_strFVSOutPrePostCycleYearTable, string p_strSummaryPrePostCycleYearTable)
			{
				return "SELECT a.standid,a.pre_year1,a.pre_year2,a.pre_year3,a.pre_year4," + 
										"b.pre_year1 AS summary_pre_year1,b.pre_year2 AS summary_pre_year2," + 
										"b.pre_year3 AS summary_pre_year3,b.pre_year4 AS summary_pre_year4 " + 
					  "FROM " + p_strSummaryPrePostCycleYearTable + " a, " + p_strFVSOutPrePostCycleYearTable + " b " +  
					  "WHERE a.standid=b.standid AND (a.pre_year1 <> b.pre_year1 OR " + 
					        "a.pre_year2 <> b.pre_year2 OR " + 
					        "a.pre_year3 <> b.pre_year3 OR " +  
					        "a.pre_year4 <> b.pre_year4)";
			}

			/// <summary>
			/// Update the FVSOut pre/post year table from the values in the FVSOUT summary pre/post table.
			/// If we find the FVSOUT summary pre/post table year in the FVSOut table then set the value
			/// to the summary year, otherwise, set the value to -1
			/// </summary>
			/// <param name="p_strFVSOutPrePostTable"></param>
			/// <param name="p_strFVSOutTable"></param>
			/// <param name="p_strFVSOutPrePostSummaryTable"></param>
			/// <returns></returns>
			static public string FVSOutputTable_AuditSelectIntoPreYears(string p_strIntoTable,string p_strFVSOutTable, string p_strFVSOutPrePostSummaryTable,string p_strCycle)
			{
				if (p_strFVSOutTable.Trim().ToUpper().IndexOf("STRCLASS",0) < 0)
				{
					return "SELECT b.standid, a.pre_year" + p_strCycle.Trim() + " " +  
						"INTO " + p_strIntoTable + " " + 
						"FROM " + p_strFVSOutPrePostSummaryTable + " a," + 
						p_strFVSOutTable + " b " + 
						"WHERE b.standid=a.standid AND b.year=a.pre_year" + p_strCycle.Trim();
				}
				else
				{
					return "SELECT b.standid, a.pre_year" + p_strCycle.Trim() + " " +  
						"INTO " + p_strIntoTable + " " + 
						"FROM " + p_strFVSOutPrePostSummaryTable + " a," + 
						p_strFVSOutTable + " b " + 
						"WHERE b.standid=a.standid AND b.year=a.pre_year" + p_strCycle.Trim() + " AND " + 
						      "b.Removal_Code IS NOT NULL AND b.Removal_Code=0";
				}

			}
			static public string FVSOutputTable_AuditSelectIntoPostYears(string p_strIntoTable,string p_strFVSOutPrePostTable,string p_strFVSOutTable)
			{
				return "SELECT a.standid, IIF(a.pre_year1=-1,-1,y1.min_post_year1) AS post_year1," + 
										 "IIF(a.pre_year2=-1,-1,y2.min_post_year2) AS post_year2," + 
										 "IIF(a.pre_year3=-1,-1,y3.min_post_year3) AS post_year3," + 
										 "IIF(a.pre_year4=-1,-1,y4.min_post_year4) AS post_year4 " + 
					   "INTO " + p_strIntoTable + " " + 
					   "FROM " + p_strFVSOutPrePostTable + " a," + 
						"(SELECT MIN(a.year) AS min_post_year4,a.standid " + 
						"FROM " + p_strFVSOutTable + " a," +  p_strFVSOutPrePostTable + " b " +  
						"WHERE a.year > b.pre_year4  AND a.standid=b.standid " + 
						"GROUP BY a.standid) y4," +
						"(SELECT MIN(a.year) AS min_post_year3,a.standid " + 
						"FROM " + p_strFVSOutTable + " a," + p_strFVSOutPrePostTable + " b " + 
						"WHERE a.year > b.pre_year3  AND a.standid=b.standid " + 
						"GROUP BY a.standid) y3," +
						"(SELECT MIN(a.year) AS min_post_year2,a.standid " + 
						"FROM " + p_strFVSOutTable + " a," + p_strFVSOutPrePostTable + " b " +  
						"WHERE a.year > b.pre_year2  AND a.standid=b.standid " + 
						"GROUP BY a.standid) y2," + 
						"(SELECT MIN(a.year) AS min_post_year1,a.standid " + 
						"FROM " + p_strFVSOutTable + " a," + p_strFVSOutPrePostTable + " b " +
						"WHERE a.year > b.pre_year1 AND a.standid=b.standid " + 
						"GROUP BY a.standid) y1 " +
					"WHERE a.standid=y4.standid AND " + 
						  "a.standid=y3.standid AND " + 
						  "a.standid=y2.standid AND " + 
						  "a.standid=y1.standid";
			}
			static public string FVSOutputTable_AuditUpdatePostYears(string p_strFVSOutPrePostTable,string p_strFVSOutPostWorkTable)
			{
				return "UPDATE " + p_strFVSOutPrePostTable.Trim() +  " a " +
					   "INNER JOIN " + p_strFVSOutPostWorkTable.Trim() + " b " + 
					   "ON a.standid=b.standid " + 
					   "SET a.post_year1=b.post_year1,a.post_year2=b.post_year2, " + 
					       "a.post_year3=b.post_year3,a.post_year4=b.post_year4";
			}
            /// <summary>
            /// Every FIA tree in the FVS_CUTLIST table should be found in the tree table. This query
            /// formats the variant and treeid column into the FVS_TREE_ID. FVS created trees are 
            /// not included (Seedlings and Compacted).
            /// </summary>
            /// <param name="p_strFvsTreeIdAuditTable"></param>
            /// <param name="p_strCasesTable"></param>
            /// <param name="p_strCutListTable"></param>
            /// <param name="p_strFVSCutListPrePostSeqNumTable"></param>
            /// <param name="p_strRxPackage"></param>
            /// <param name="p_strRx"></param>
            /// <param name="p_strRxCycle"></param>
            /// <param name="p_strRxYear"></param>
            /// <returns></returns>
            static public string FVSOutputTable_AuditFVSTreeId(string p_strFvsTreeIdAuditTable,
                                                               string p_strCasesTable,
                                                               string p_strCutListTable,
                                                               string p_strFVSCutListPrePostSeqNumTable,
                                                               string p_strRxPackage,
                                                               string p_strRx,
                                                               string p_strRxCycle,
                                                               string p_strRxYear)
            {
                return "INSERT INTO " + p_strFvsTreeIdAuditTable + " " +
                                "(biosum_cond_id, rxpackage,rx,rxcycle,rxyear,fvs_variant, fvs_tree_id) " +
                                 "SELECT DISTINCT c.StandID AS biosum_cond_id,'" + p_strRxPackage.Trim() + "' AS rxpackage," +
                                "'" + p_strRx.Trim() + "' AS rx,'" + p_strRxCycle.Trim() + "' AS rxcycle," +
                                "CSTR(t.year) AS rxyear," +
                                "c.Variant AS fvs_variant, " +
                                "Trim(t.treeid) AS fvs_tree_id " +
                                "FROM " + p_strCasesTable + " c," + p_strCutListTable + " t," + p_strFVSCutListPrePostSeqNumTable + " p " +
                                "WHERE c.CaseID = t.CaseID AND t.standid=p.standid AND t.year=p.year AND  " + 
                                      "p.cycle" + p_strRxCycle.Trim() + "_PRE_YN='Y' AND " + 
                                      "MID(t.treeid, 1, 2) <> 'ES'  AND MID(t.treeid, 1, 2)<> 'CM'";
  
            }
            static public string[] FVSOutputTable_AuditPostSummaryFVS(string p_strRxTable,string p_strRxPackageTable,string p_strTreeTable,
                string p_strPlotTable,string p_strCondTable, string p_strPostAuditSummaryTable,string p_strFvsTreeTableName,
                string p_strRxPackage)
            {
                string[] sqlArray = new string[15];
                
                sqlArray[0] = "SELECT * INTO rxpackage_work_table FROM (" +
                            "SELECT	 rxpackage, simyear1_fvscycle AS rxcycle, simyear1_rx as RX FROM " + p_strRxPackageTable + " " +
                            "UNION " +
                            "SELECT	 rxpackage,simyear2_fvscycle AS rxcycle, simyear2_rx as RX FROM " + p_strRxPackageTable + " " +
                            "UNION " +
                            "SELECT	 rxpackage,simyear3_fvscycle AS rxcycle, simyear3_rx as RX FROM " + p_strRxPackageTable + " " +
                            "UNION " +
                            "SELECT	 rxpackage,simyear4_fvscycle AS rxcycle, simyear4_rx as RX FROM " + p_strRxPackageTable + ")";

                sqlArray[1] = "SELECT BIOSUM_COND_ID INTO cond_biosum_cond_id_work_table FROM " + p_strCondTable;
                sqlArray[2] = "ALTER TABLE cond_biosum_cond_id_work_table ALTER COLUMN biosum_cond_id CHAR(25) PRIMARY KEY";

                sqlArray[3] = "SELECT BIOSUM_PLOT_ID INTO plot_biosum_plot_id_work_table FROM " + p_strPlotTable;
                sqlArray[4] = "ALTER TABLE plot_biosum_plot_id_work_table ALTER COLUMN biosum_plot_id CHAR(24) PRIMARY KEY";

                sqlArray[5] = "SELECT biosum_cond_id, SPCD, FVS_TREE_ID, DIA INTO tree_fvs_tree_id_work_table FROM " + p_strTreeTable + " WHERE FVS_TREE_ID IS NOT NULL AND LEN(TRIM(FVS_TREE_ID)) > 0";
                sqlArray[6] = "ALTER TABLE tree_fvs_tree_id_work_table ADD PRIMARY KEY (biosum_cond_id,fvs_tree_id);";

                sqlArray[7] = "SELECT DISTINCT " +
                                "IIF(BIOSUM_COND_ID IS NULL OR LEN(TRIM(BIOSUM_COND_ID)) = 0,''," +
                                "IIF(LEN(TRIM(BIOSUM_COND_ID)) >= 24,MID(BIOSUM_COND_ID,1,24),BIOSUM_COND_ID)) AS BIOSUM_PLOT_ID," +
                                "BIOSUM_COND_ID " +
                              "INTO fvs_tree_unique_biosum_plot_id_work_table " +
                              "FROM " + p_strFvsTreeTableName;

                sqlArray[8] = "ALTER TABLE fvs_tree_unique_biosum_plot_id_work_table ALTER COLUMN biosum_plot_id CHAR(24)";

                sqlArray[9] = "ALTER TABLE fvs_tree_unique_biosum_plot_id_work_table ALTER COLUMN biosum_cond_id CHAR(25)";

                sqlArray[10] = "CREATE TABLE fvs_tree_biosum_plot_id_work_table (" +
                                "ID INTEGER, " +
                                "RXPACKAGE CHAR(3), " +
                                "biosum_plot_id CHAR(24), " +
                                "biosum_cond_id CHAR(25)) ";

                sqlArray[11] = "INSERT INTO fvs_tree_biosum_plot_id_work_table " +
                                "SELECT ID," +
                                "'" + p_strRxPackage + "' AS RXPACKAGE," +
                                "IIF(BIOSUM_COND_ID IS NULL OR LEN(TRIM(BIOSUM_COND_ID)) = 0,''," +
                                "IIF(LEN(TRIM(BIOSUM_COND_ID)) >= 24,MID(BIOSUM_COND_ID,1,24),BIOSUM_COND_ID)) AS BIOSUM_PLOT_ID," +
                                "BIOSUM_COND_ID FROM " + p_strFvsTreeTableName;
                
                //sqlArray[10] = "SELECT ID," +
                //                "'" +  p_strFVSTreeFileName + "' AS FVS_TREE_FILE," +
                //                "IIF(BIOSUM_COND_ID IS NULL OR LEN(TRIM(BIOSUM_COND_ID)) = 0,''," +
                //                "IIF(LEN(TRIM(BIOSUM_COND_ID)) >= 24,MID(BIOSUM_COND_ID,1,24),BIOSUM_COND_ID)) AS BIOSUM_PLOT_ID," +
                //                "BIOSUM_COND_ID " +
                //              "INTO fvs_tree_biosum_plot_id_work_table " +
                //              "FROM " + p_strFvsTreeTableName;

                //sqlArray[11] = "ALTER TABLE fvs_tree_biosum_plot_id_work_table ALTER COLUMN biosum_plot_id CHAR(24)";

                //sqlArray[12] = "ALTER TABLE fvs_tree_biosum_plot_id_work_table ALTER COLUMN biosum_cond_id CHAR(25)";

                //sqlArray[13] = "ALTER TABLE fvs_tree_biosum_plot_id_work_table ALTER COLUMN fvs_tree_file CHAR(26)";

                sqlArray[12] = $@"DELETE FROM {p_strPostAuditSummaryTable} WHERE RXPACKAGE = '{p_strRxPackage}'"; 

                sqlArray[13] =
                    "INSERT INTO  " + p_strPostAuditSummaryTable + " " +
                    "SELECT * FROM " +
                    "(SELECT DISTINCT " +
                        "'001' AS [INDEX]," +
                        "'" + p_strRxPackage + "' AS RXPACKAGE," +
                        "'BIOSUM_COND_ID' AS COLUMN_NAME," +
                        "biosum_cond_id_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                        "biosum_cond_id_not_found_in_cond_table_count.NOT_FOUND_IN_COND_TABLE_COUNT AS NF_IN_COND_TABLE_ERROR," +
                        "biosum_cond_id_not_found_in_plot_table_count.NOT_FOUND_IN_PLOT_TABLE_COUNT AS NF_IN_PLOT_TABLE_ERROR," +
                        "'NA'  AS VALUE_ERROR," +
                        "'NA' AS NF_IN_RX_TABLE_ERROR," +
                        "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                        "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                        "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                        "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                     "FROM fvs_tree_unique_biosum_plot_id_work_table fvs," +
                        "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM fvs_tree_unique_biosum_plot_id_work_table " +
                         "WHERE BIOSUM_COND_ID IS NULL OR LEN(TRIM(BIOSUM_COND_ID))=0) biosum_cond_id_no_value_count," +
                        "(SELECT CSTR(COUNT(*)) AS NOT_FOUND_IN_COND_TABLE_COUNT FROM fvs_tree_unique_biosum_plot_id_work_table a " +
                         "WHERE a.BIOSUM_COND_ID IS NOT NULL AND " +
                               "LEN(TRIM(a.BIOSUM_COND_ID)) >  0 AND " +
                               "NOT EXISTS (SELECT b.BIOSUM_COND_ID FROM cond_biosum_cond_id_work_table b " +
                                           "WHERE a.BIOSUM_COND_ID = b.BIOSUM_COND_ID)) biosum_cond_id_not_found_in_cond_table_count," +
                        "(SELECT CSTR(COUNT(*)) AS NOT_FOUND_IN_PLOT_TABLE_COUNT FROM fvs_tree_unique_biosum_plot_id_work_table a " +
                         "WHERE a.BIOSUM_COND_ID IS NOT NULL AND LEN(TRIM(a.BIOSUM_COND_ID)) >  0 AND " +
                               "NOT EXISTS (SELECT b.BIOSUM_PLOT_ID FROM plot_biosum_plot_id_work_table b " +
                                           "WHERE b.BIOSUM_PLOT_ID = a.BIOSUM_PLOT_ID)) biosum_cond_id_not_found_in_plot_table_count " +
                "UNION " +
                "SELECT DISTINCT " +
                   "'002' AS [INDEX]," +
                    "'" + p_strRxPackage + "' AS RXPACKAGE," +
                   "'RXCYCLE' AS COLUMN_NAME," +
                   "rxcycle_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                   "'NA' AS NF_IN_COND_TABLE_ERROR," +
                   "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                   "rxcycle_value_notvalid_count.value_notvalid_count AS VALUE_ERROR," +
                   "'NA' AS NF_IN_RX_TABLE_ERROR," +
                   "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                   "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                   "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                   "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                "FROM " + p_strFvsTreeTableName + " fvs," +
                   "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                    "WHERE RXCYCLE IS NULL OR LEN(TRIM(RXCYCLE))=0) rxcycle_no_value_count," +
                   "(SELECT CSTR(COUNT(*)) AS VALUE_NOTVALID_COUNT FROM " + p_strFvsTreeTableName + " a " +
                    "WHERE a.RXCYCLE IS NULL OR a.RXCYCLE NOT IN ('1','2','3','4')) rxcycle_value_notvalid_count " +
                     "UNION " +
                     "SELECT DISTINCT " +
                        "'003' AS [INDEX]," +
						"'" + p_strRxPackage + "' AS RXPACKAGE," +
                        "'RXPACKAGE' AS COLUMN_NAME," +
                        "rxpackage_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                        "'NA' AS NF_IN_COND_TABLE_ERROR," +
                        "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                        "'NA' AS  VALUE_ERROR," +
                        "'NA' AS NF_IN_RX_TABLE_ERROR," +
                        "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                        "notfound_in_rxpackage_table.NF_IN_RXPACKAGE_TABLE_ERROR," +
                        "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                        "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                     "FROM " + p_strFvsTreeTableName + " fvs," +
                        "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                         "WHERE RXPACKAGE IS NULL OR LEN(TRIM(RXPACKAGE))=0) rxpackage_no_value_count," +
                        "(SELECT CSTR(COUNT(*)) AS NF_IN_RXPACKAGE_TABLE_ERROR FROM " + p_strFvsTreeTableName + " fvs " +
                         "WHERE fvs.RXPACKAGE IS NOT NULL AND LEN(TRIM(fvs.RXPACKAGE)) > 0 AND " +
                               "NOT EXISTS (SELECT rxp.rxpackage FROM " + p_strRxPackageTable + " rxp " +
                                           "WHERE fvs.rxpackage = rxp.rxpackage)) notfound_in_rxpackage_table " +
                     "UNION " +
                     "SELECT DISTINCT " +
                        "'004' AS [INDEX]," +
                        "'" + p_strRxPackage + "' AS RXPACKAGE," +
                        "'RX' AS COLUMN_NAME," +
                        "rx_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                        "'NA' AS NF_IN_COND_TABLE_ERROR," +
                        "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                        "'NA'  AS VALUE_ERROR," +
                        "rx_not_found_in_rx_table_count.NOT_FOUND_IN_RX_TABLE_COUNT  AS NF_IN_RX_TABLE_ERROR," +
                        "not_found_rxpackage_rxcycle_rx_combo_count.NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                        "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                        "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                        "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                     "FROM " + p_strFvsTreeTableName + " fvs," +
                        "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                         "WHERE RX IS NULL OR LEN(TRIM(RX))=0) rx_no_value_count," +
                        "(SELECT CSTR(COUNT(*)) AS NOT_FOUND_IN_RX_TABLE_COUNT FROM " + p_strFvsTreeTableName + " a " +
                         "WHERE a.RX IS NOT NULL AND LEN(TRIM(a.RX)) >  0 AND " +
                               "NOT EXISTS (SELECT b.RX FROM " + p_strRxTable + " b " +
                                           "WHERE a.RX = b.RX)) rx_not_found_in_rx_table_count," +
                        "(SELECT CSTR(COUNT(*)) AS NF_RXPACKAGE_RXCYCLE_RX_ERROR FROM " + p_strFvsTreeTableName + " fvs " +
                         "WHERE fvs.RX IS NOT NULL AND LEN(TRIM(fvs.RX)) >  0 AND " +
                               "NOT EXISTS (SELECT rxp.RX FROM rxpackage_work_table rxp " +
                                           "WHERE trim(fvs.rxpackage) = trim(rxp.rxpackage) AND " +
                                                 "TRIM(fvs.rxcycle)=TRIM(rxp.rxcycle) AND " +
                                                 "TRIM(fvs.rx)=TRIM(rxp.rx))) not_found_rxpackage_rxcycle_rx_combo_count " +
                "UNION " +
                "SELECT DISTINCT " +
                   "'005' AS [INDEX]," +
                    "'" + p_strRxPackage + "' AS RXPACKAGE," +
                   "'RXYEAR' AS COLUMN_NAME," +
                   "rxyear_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                   "'NA' AS NF_IN_COND_TABLE_ERROR," +
                   "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                   "'NA' AS  VALUE_ERROR," +
                   "'NA' AS NF_IN_RX_TABLE_ERROR," +
                   "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                   "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                   "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                   "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                "FROM " + p_strFvsTreeTableName + " fvs," +
                   "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                    "WHERE RXYEAR IS NULL OR LEN(TRIM(RXYEAR))=0) rxyear_no_value_count " +
                     "UNION " +
                     "SELECT DISTINCT " +
                        "'006' AS [INDEX]," +
                         "'" + p_strRxPackage + "' AS RXPACKAGE," +
                        "'DBH' AS COLUMN_NAME," +
                        "dbh_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                        "'NA' AS NF_IN_COND_TABLE_ERROR," +
                        "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                        "dia_value_error.VALUE_ERROR_COUNT AS  VALUE_ERROR," +
                        "'NA' AS NF_IN_RX_TABLE_ERROR," +
                        "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                        "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                        "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                        "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                     "FROM " + p_strFvsTreeTableName + " fvs," +
                        "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                         "WHERE DBH IS NULL) dbh_no_value_count, " +
                        "(SELECT CSTR(COUNT(*)) AS VALUE_ERROR_COUNT FROM " + p_strFvsTreeTableName + " fvs " +
                         "INNER JOIN tree_fvs_tree_id_work_table fia ON fvs.fvs_tree_id = fia.fvs_tree_id and fvs.biosum_cond_id = fia.biosum_cond_id " +
                         "WHERE fvs.FvsCreatedTree_YN='N' AND  " +
                               "fvs.rxcycle='1' AND " +
                               "fvs.FVS_TREE_ID IS NOT NULL AND " +
                               "LEN(TRIM(fvs.FVS_TREE_ID)) >  0 AND " +
                               "fvs.DBH <> fia.DIA) dia_value_error " +
                     "UNION " +
                     "SELECT DISTINCT " +
                        "'007' AS [INDEX]," +
                        "'" + p_strRxPackage + "' AS RXPACKAGE," +
                        "'TPA' AS COLUMN_NAME," +
                        "tpa_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                        "'NA' AS NF_IN_COND_TABLE_ERROR," +
                        "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                        "'NA' AS  VALUE_ERROR," +
                        "'NA' AS NF_IN_RX_TABLE_ERROR," +
                        "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                        "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                        "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                        "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                     "FROM " + p_strFvsTreeTableName + " fvs," +
                        "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                         "WHERE TPA IS NULL) tpa_no_value_count " +
                     "UNION " +
                     "SELECT DISTINCT " +
                        "'008' AS [INDEX]," +
                         "'" + p_strRxPackage + "' AS RXPACKAGE," +
                        "'VOLCFNET' AS COLUMN_NAME," +
                        "volcfnet_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                        "'NA' AS NF_IN_COND_TABLE_ERROR," +
                        "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                        "'NA' AS  VALUE_ERROR," +
                        "'NA' AS NF_IN_RX_TABLE_ERROR," +
                        "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                        "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                        "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                        "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                     "FROM " + p_strFvsTreeTableName + " fvs," +
                        "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                         "WHERE VOLCFNET IS NULL) volcfnet_no_value_count " +
                "UNION " +
                "SELECT DISTINCT " +
                   "'009' AS [INDEX]," +
                    "'" + p_strRxPackage + "' AS RXPACKAGE," +
                   "'VOLTSGRS' AS COLUMN_NAME," +
                   "voltsgrs_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                   "'NA' AS NF_IN_COND_TABLE_ERROR," +
                   "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                   "'NA' AS  VALUE_ERROR," +
                   "'NA' AS NF_IN_RX_TABLE_ERROR," +
                   "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                   "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                   "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                   "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                "FROM " + p_strFvsTreeTableName + " fvs," +
                   "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                    "WHERE VOLTSGRS IS NULL) voltsgrs_no_value_count " +
                "UNION " +
                "SELECT DISTINCT " +
                   "'010' AS [INDEX]," +
				   "'" + p_strRxPackage + "' AS RXPACKAGE," +
                   "'VOLCFGRS' AS COLUMN_NAME," +
                   "volcfgrs_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                   "'NA' AS NF_IN_COND_TABLE_ERROR," +
                   "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                   "'NA' AS  VALUE_ERROR," +
                   "'NA' AS NF_IN_RX_TABLE_ERROR," +
                   "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                   "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                   "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                   "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                "FROM " + p_strFvsTreeTableName + " fvs," +
                    "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                     "WHERE DBH IS NOT NULL AND DBH >= 5 AND VOLCFGRS IS NULL) volcfgrs_no_value_count " +
                "UNION " +
                "SELECT DISTINCT " +
                   "'011' AS [INDEX]," +
                   "'" + p_strRxPackage + "' AS RXPACKAGE," +
                   "'DRYBIOT' AS COLUMN_NAME," +
                   "drybiot_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                   "'NA' AS NF_IN_COND_TABLE_ERROR," +
                   "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                   "'NA' AS  VALUE_ERROR," +
                   "'NA' AS NF_IN_RX_TABLE_ERROR," +
                   "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                   "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                   "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                   "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                "FROM " + p_strFvsTreeTableName + " fvs," +
                   "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                    "WHERE DRYBIOT IS NULL) drybiot_no_value_count " +
                "UNION " +
                "SELECT DISTINCT " +
                   "'012' AS [INDEX]," +
                   "'" + p_strRxPackage + "' AS RXPACKAGE," +
                   "'DRYBIOM' AS COLUMN_NAME," +
                   "drybiom_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                   "'NA' AS NF_IN_COND_TABLE_ERROR," +
                   "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                   "'NA' AS  VALUE_ERROR," +
                   "'NA' AS NF_IN_RX_TABLE_ERROR," +
                   "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                   "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                   "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                   "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                "FROM " + p_strFvsTreeTableName + " fvs," +
                   "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                    "WHERE DBH IS NOT NULL AND DBH >= 5 AND DRYBIOM IS NULL) drybiom_no_value_count " +
                "UNION " +
                "SELECT DISTINCT " +
                   "'013' AS [INDEX]," +
                   "'" + p_strRxPackage + "' AS RXPACKAGE," +
                   "'FVS_TREE_ID' AS COLUMN_NAME," +
                   "fvs_tree_id_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                   "'NA' AS NF_IN_COND_TABLE_ERROR," +
                   "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                   "'NA' AS  VALUE_ERROR," +
                   "'NA' AS NF_IN_RX_TABLE_ERROR," +
                   "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                   "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                   "fvs_tree_id_not_found_in_tree_table_count.NOT_FOUND_IN_TREE_TABLE_COUNT AS NF_IN_TREE_TABLE_ERROR," +
                   "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                "FROM " + p_strFvsTreeTableName + " fvs," +
                   "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                    "WHERE FVS_TREE_ID IS NULL OR LEN(TRIM(FVS_TREE_ID))=0) fvs_tree_id_no_value_count," +
                   "(SELECT CSTR(COUNT(*)) AS NOT_FOUND_IN_TREE_TABLE_COUNT FROM " + p_strFvsTreeTableName + " a " +
                    "WHERE a.FvsCreatedTree_YN='N' AND  " +
                          "a.FVS_TREE_ID IS NOT NULL AND " +
                          "LEN(TRIM(a.FVS_TREE_ID)) >  0 AND " +
                          "NOT EXISTS (SELECT b.FVS_TREE_ID FROM tree_fvs_tree_id_work_table b " +
                              "WHERE a.fvs_tree_id = b.fvs_tree_id and a.biosum_cond_id = b.biosum_cond_id)) fvs_tree_id_not_found_in_tree_table_count " +
                "UNION " +
                "SELECT DISTINCT " +
                   "'014' AS [INDEX]," +
                    "'" + p_strRxPackage + "' AS RXPACKAGE," +
                   "'FVSCREATEDTREE_YN' AS COLUMN_NAME," +
                   "fvscreatedtree_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                   "'NA' AS NF_IN_COND_TABLE_ERROR," +
                   "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                   "'NA' AS  VALUE_ERROR," +
                   "'NA' AS NF_IN_RX_TABLE_ERROR," +
                   "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                   "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                   "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                   "'NA' AS TREE_SPECIES_CHANGE_WARNING " +
                "FROM " + p_strFvsTreeTableName + " fvs," +
                   "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                    "WHERE FvsCreatedTree_YN IS NULL OR LEN(TRIM(FvsCreatedTree_YN))=0) fvscreatedtree_no_value_count " +
                "UNION " +
                "SELECT DISTINCT " +
                   "'015' AS [INDEX]," +
                   "'" + p_strRxPackage + "' AS RXPACKAGE," +
                   "'FVS_SPECIES' AS COLUMN_NAME," +
                   "fvs_species_no_value_count.NOVALUE_COUNT AS NOVALUE_ERROR," +
                   "'NA' AS NF_IN_COND_TABLE_ERROR," +
                   "'NA' AS NF_IN_PLOT_TABLE_ERROR," +
                   "'NA' AS  VALUE_ERROR," +
                   "'NA' AS NF_IN_RX_TABLE_ERROR," +
                   "'NA' AS NF_RXPACKAGE_RXCYCLE_RX_ERROR," +
                   "'NA' AS NF_IN_RXPACKAGE_TABLE_ERROR," +
                   "'NA' AS NF_IN_TREE_TABLE_ERROR," +
                  "fvs_species_change_count.TREE_SPECIES_CHANGE_WARNING " +
                "FROM " + p_strFvsTreeTableName + " fvs," +
                   "(SELECT CSTR(COUNT(*)) AS NOVALUE_COUNT FROM " + p_strFvsTreeTableName + " " +
                    "WHERE FVS_SPECIES IS NULL OR LEN(TRIM(FVS_SPECIES))=0) fvs_species_no_value_count," +
                   "(SELECT CSTR(COUNT(*)) AS TREE_SPECIES_CHANGE_WARNING " +
                    "FROM " + p_strFvsTreeTableName + " a " +
                    "INNER JOIN tree_fvs_tree_id_work_table b " +
                    "ON a.fvs_tree_id = b.fvs_tree_id and a.biosum_cond_id = b.biosum_cond_id " +
                    "WHERE a.FvsCreatedTree_YN='N' AND " +
                          "a.FVS_TREE_ID IS NOT NULL AND " +
                          "LEN(TRIM(a.FVS_TREE_ID)) >  0 AND " +
                          "VAL(a.FVS_SPECIES) <> b.SPCD) fvs_species_change_count)";
                sqlArray[14] = $@"UPDATE {p_strPostAuditSummaryTable} SET CREATED_DATE='{DateTime.Now.ToString().Trim()}' 
                                WHERE RXPACKAGE = '{p_strRxPackage}'";

                return sqlArray;
            }

            public static string[] FVSOutputTable_AuditPostSummaryDetailFVS_SPCDCHANGE_WARNING(string p_strInsertTable, string p_strPostAuditSummaryTable, string p_strFvsTreeTableName,string p_strTreeTable, string p_strFvsVariant, string p_strRxPackage)
            {
                string[] sqlArray = new string[2];

                sqlArray[0] = "DELETE FROM " + p_strInsertTable + " WHERE RXPACKAGE = '" + p_strRxPackage + "'";

                sqlArray[1] = "INSERT INTO  " + p_strInsertTable + " " +
                                   "SELECT * FROM " +
                                        "(SELECT 'FVS_SPECIES' AS COLUMN_NAME," +
                                                "'SPCD DOES NOT MATCH: FVS=' + TRIM(FVS.FVS_SPECIES) + ' FIA=' + TRIM(CSTR(FIA.SPCD)) AS WARNING_DESC," +
                                                "fvs.ID," +
                                                "fvs.BIOSUM_COND_ID," +
                                                "fvs.RXPACKAGE," +
                                                "fvs.RXCYCLE," +
                                                "fvs.FVS_TREE_ID AS FVS_TREE_FVS_TREE_ID," +
                                                "fia.FVS_TREE_ID AS FIA_TREE_FVS_TREE_ID," +
                                                "FVS.FVS_SPECIES AS FVS_TREE_SPCD," +
                                                "FIA.SPCD AS FIA_TREE_SPCD," +
                                                "FVS.DBH AS FVS_TREE_DIA," +
                                                "FIA.DIA AS FIA_TREE_DIA," +
                                                "FVS.ESTHT AS FVS_TREE_ESTHT," +
                                                "FIA.HT AS FIA_TREE_ESTHT," +
                                                "FVS.HT AS FVS_TREE_ACTUALHT," +
                                                "FIA.ACTUALHT AS FIA_TREE_ACTUALHT," +
                                                "FVS.PCTCR AS FVS_TREE_CR," +
                                                "FIA.CR AS FIA_TREE_CR," +
                                                "FVS.VOLCSGRS AS FVS_TREE_VOLCSGRS," +
                                                "FIA.VOLCSGRS AS FIA_TREE_VOLCSGRS," +
                                                "FVS.VOLCFGRS AS FVS_TREE_VOLCFGRS," +
                                                "FIA.VOLCFGRS AS FIA_TREE_VOLCFGRS," +
                                                "FVS.VOLCFNET AS FVS_TREE_VOLCFNET," +
                                                "FIA.VOLCFNET AS FIA_TREE_VOLCFNET," +
                                                "FVS.VOLTSGRS AS FVS_TREE_VOLTSGRS," +
                                                "FIA.VOLTSGRS AS FIA_TREE_VOLTSGRS," +
                                                "FVS.DRYBIOT AS FVS_TREE_DRYBIOT," +
                                                "FIA.DRYBIOT AS FIA_TREE_DRYBIOT," +
                                                "FVS.DRYBIOM AS FVS_TREE_DRYBIOM," +
                                                "FIA.DRYBIOM AS FIA_TREE_DRYBIOM," +
                                                "FIA.STATUSCD AS FIA_TREE_STATUSCD," +
                                                "FIA.TREECLCD AS FIA_TREE_TREECLCD," +
                                                "FIA.CULL AS FIA_TREE_CULL," +
                                                "FIA.ROUGHCULL AS FIA_TREE_ROUGHCULL," +
                                                "FVS.FVSCREATEDTREE_YN " +
                                         "FROM " + p_strFvsTreeTableName + " fvs " +
                                         "INNER JOIN " + p_strTreeTable + " fia " +
                                         "ON fvs.fvs_tree_id = fia.fvs_tree_id and fvs.biosum_cond_id=fia.biosum_cond_id " +
                                         "WHERE fvs.FvsCreatedTree_YN='N' AND " +
                                               "fvs.FVS_TREE_ID IS NOT NULL AND " +
                                               "LEN(TRIM(fvs.FVS_TREE_ID)) >  0 AND " +
                                               "fvs.fvs_variant = '" + p_strFvsVariant + "' AND " +
                                               "fvs.rxpackage = '" + p_strRxPackage + "' AND " +
                                                "VAL(fvs.FVS_SPECIES) <> fia.SPCD)";

                return sqlArray;
            }
            public static string[] FVSOutputTable_AuditPostSummaryDetailFVS_TREEMATCH_ERROR(string p_strInsertTable, string p_strPostAuditSummaryTable, string p_strFvsTreeTableName, 
                string p_strTreeTable, string p_strFvsVariant, string p_strRxPackage)
            {
                string[] sqlArray = new string[2];

                sqlArray[0] = "DELETE FROM " + p_strInsertTable + " WHERE RXPACKAGE = '" + p_strRxPackage + "'";

                sqlArray[1] = "INSERT INTO  " + p_strInsertTable + " " +
                                   "SELECT * FROM " +
                                        "(SELECT 'DBH' AS COLUMN_NAME," +
                                                "'MATCHING FVS AND FIA TREE HAS DIFFERENT DBH VALUES FOR RX CYCLE 1' AS ERROR_DESC," +
                                                "fvs.ID," +
                                                "fvs.BIOSUM_COND_ID," +
                                                "fvs.RXPACKAGE," +
                                                "fvs.RXCYCLE," +
                                                "fvs.FVS_TREE_ID AS FVS_TREE_FVS_TREE_ID," +
                                                "fia.FVS_TREE_ID AS FIA_TREE_FVS_TREE_ID," +
                                                "FVS.FVS_SPECIES AS FVS_TREE_SPCD," +
                                                "FIA.SPCD AS FIA_TREE_SPCD," +
                                                "FVS.DBH AS FVS_TREE_DIA," +
                                                "FIA.DIA AS FIA_TREE_DIA," +
                                                "FVS.ESTHT AS FVS_TREE_ESTHT," +
                                                "FIA.HT AS FIA_TREE_ESTHT," +
                                                "FVS.HT AS FVS_TREE_ACTUALHT," +
                                                "FIA.ACTUALHT AS FIA_TREE_ACTUALHT," +
                                                "FVS.PCTCR AS FVS_TREE_CR," +
                                                "FIA.CR AS FIA_TREE_CR," +
                                                "FVS.VOLCSGRS AS FVS_TREE_VOLCSGRS," +
                                                "FIA.VOLCSGRS AS FIA_TREE_VOLCSGRS," +
                                                "FVS.VOLCFGRS AS FVS_TREE_VOLCFGRS," +
                                                "FIA.VOLCFGRS AS FIA_TREE_VOLCFGRS," +
                                                "FVS.VOLCFNET AS FVS_TREE_VOLCFNET," +
                                                "FIA.VOLCFNET AS FIA_TREE_VOLCFNET," +
                                                "FVS.VOLTSGRS AS FVS_TREE_VOLTSGRS," +
                                                "FIA.VOLTSGRS AS FIA_TREE_VOLTSGRS," +
                                                "FVS.DRYBIOT AS FVS_TREE_DRYBIOT," +
                                                "FIA.DRYBIOT AS FIA_TREE_DRYBIOT," +
                                                "FVS.DRYBIOM AS FVS_TREE_DRYBIOM," +
                                                "FIA.DRYBIOM AS FIA_TREE_DRYBIOM," +
                                                "FIA.STATUSCD AS FIA_TREE_STATUSCD," +
                                                "FIA.TREECLCD AS FIA_TREE_TREECLCD," +
                                                "FIA.CULL AS FIA_TREE_CULL," +
                                                "FIA.ROUGHCULL AS FIA_TREE_ROUGHCULL," +
                                                "FVS.FVSCREATEDTREE_YN " +
                                         "FROM " + p_strFvsTreeTableName + " fvs " +
                                         "INNER JOIN " + p_strTreeTable + " fia " +
                                         "ON fvs.fvs_tree_id = fia.fvs_tree_id and fvs.biosum_cond_id = fia.biosum_cond_id " +
                                         "WHERE fvs.FvsCreatedTree_YN='N' AND " +
                                               "fvs.RXCYCLE IS NOT NULL AND " +
                                               "fvs.RXPACKAGE = '" + p_strRxPackage + "' AND " +
                                               "fvs.FVS_VARIANT = '" + p_strFvsVariant + "' AND " +
                                               "LEN(TRIM(fvs.RXCYCLE)) > 0 AND " + 
                                               "fvs.RXCYCLE = '1' AND " +
                                               "fvs.FVS_TREE_ID IS NOT NULL AND " +
                                               "LEN(TRIM(fvs.FVS_TREE_ID)) >  0 AND " +
                                               "fvs.DBH <> fia.DIA) ";

                return sqlArray;
            }
            /// <summary>
            /// SQL for post-processing audit of the BIOSUMCALC\FVS_TREE tables. Find required columns with no values.
            /// </summary>
            /// <param name="p_strInsertTable"></param>
            /// <param name="p_strPostAuditSummaryTable"></param>
            /// <param name="p_strFvsTreeTableName"></param>
            /// <param name="p_strFVSTreeFileName"></param>
            /// <returns></returns>
            public static string[] FVSOutputTable_AuditPostSummaryDetailFVS_NOVALUE_ERROR(string p_strInsertTable,string p_strPostAuditSummaryTable,string p_strFvsTreeTableName,string p_strRxPackage)
            {
                
                string[] sqlArray = new string[2];

                sqlArray[0] = "DELETE FROM " + p_strInsertTable + " WHERE RXPACKAGE = '" + p_strRxPackage + "'";

                sqlArray[1] =    "INSERT INTO  " + p_strInsertTable + " " +
                                    "SELECT * FROM " +
                                         "(SELECT a.COLUMN_NAME,'NULLS NOT ALLOWED' AS ERROR_DESC, FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                             "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE BIOSUM_COND_ID IS NULL OR LEN(TRIM(BIOSUM_COND_ID))=0) FVS_TREE " +
                                          "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                 "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                 "a.NOVALUE_ERROR <> 'NA' AND " +
                                                 "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                 "TRIM(a.COLUMN_NAME) = 'BIOSUM_COND_ID' AND a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION " +
                                          "SELECT a.COLUMN_NAME,'NULLS NOT ALLOWED' AS ERROR_DESC, FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE RXCYCLE IS NULL OR LEN(TRIM(RXCYCLE))=0) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                   "a.NOVALUE_ERROR <> 'NA' AND " +
                                                   "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                   "TRIM(a.COLUMN_NAME) = 'RXCYCLE' AND " +
                                                   "a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION " +
                                          "SELECT a.COLUMN_NAME,'NULLS NOT ALLOWED' AS ERROR_DESC, FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE RXPACKAGE IS NULL OR LEN(TRIM(RXPACKAGE))=0) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                   "a.NOVALUE_ERROR <> 'NA' AND " +
                                                   "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                   "TRIM(a.COLUMN_NAME) = 'RXPACKAGE' AND " +
                                                   "a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION " +
                                          "SELECT a.COLUMN_NAME,'NULLS NOT ALLOWED' AS ERROR_DESC, FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE RX IS NULL OR LEN(TRIM(RX))=0) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                   "a.NOVALUE_ERROR <> 'NA' AND " +
                                                   "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                   "TRIM(a.COLUMN_NAME) = 'RX' AND " +
                                                   "a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION " +
                                          "SELECT a.COLUMN_NAME,'NULLS NOT ALLOWED' AS ERROR_DESC, FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE RXYEAR IS NULL OR LEN(TRIM(RXYEAR))=0) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                   "a.NOVALUE_ERROR <> 'NA' AND " +
                                                   "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                   "TRIM(a.COLUMN_NAME) = 'RXYEAR' AND " +
                                                   "a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION " +
                                          "SELECT a.COLUMN_NAME,'NULLS NOT ALLOWED' AS ERROR_DESC, FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE DBH IS NULL) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                   "a.NOVALUE_ERROR <> 'NA' AND " +
                                                   "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                   "TRIM(a.COLUMN_NAME) = 'DBH' AND " +
                                                   "a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION " +
                                          "SELECT a.COLUMN_NAME,'NULLS NOT ALLOWED' AS ERROR_DESC, FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE TPA IS NULL) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                   "a.NOVALUE_ERROR <> 'NA' AND " +
                                                   "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                   "TRIM(a.COLUMN_NAME) = 'TPA' AND " +
                                                   "a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION " +
                                          "SELECT a.COLUMN_NAME, 'NULLS NOT ALLOWED WHEN DBH >= 5 INCHES' AS ERROR_DESC,FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE DBH IS NOT NULL AND DBH >= 5 AND VOLCFNET IS NULL) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                   "a.NOVALUE_ERROR <> 'NA' AND " +
                                                   "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                   "TRIM(a.COLUMN_NAME) = 'VOLCFNET' AND " +
                                                   "a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION " +
                                          "SELECT a.COLUMN_NAME, 'NULLS NOT ALLOWED' AS ERROR_DESC,FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE VOLTSGRS IS NULL) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                   "a.NOVALUE_ERROR <> 'NA' AND " +
                                                   "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                   "TRIM(a.COLUMN_NAME) = 'VOLTSGRS' AND " +
                                                   "a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION " +
                                          "SELECT a.COLUMN_NAME, 'NULLS NOT ALLOWED WHEN DBH >= 5 INCHES' AS ERROR_DESC,FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE DBH IS NOT NULL AND DBH >= 5 AND VOLCFGRS IS NULL) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                   "a.NOVALUE_ERROR <> 'NA' AND " +
                                                   "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                   "TRIM(a.COLUMN_NAME) = 'VOLCFGRS' AND " +
                                                   "a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION " +
                                          "SELECT a.COLUMN_NAME, 'NULLS NOT ALLOWED' AS ERROR_DESC,FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE DRYBIOT IS NULL) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                   "a.NOVALUE_ERROR <> 'NA' AND " +
                                                   "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                   "TRIM(a.COLUMN_NAME) = 'DRYBIOT' AND " +
                                                   "a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION	" +
                                          "SELECT a.COLUMN_NAME, 'NULLS NOT ALLOWED WHEN DBH >= 5 INCHES' AS ERROR_DESC,FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE DBH IS NOT NULL AND DBH >= 5 AND DRYBIOM IS NULL) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                   "a.NOVALUE_ERROR <> 'NA' AND " +
                                                   "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                   "TRIM(a.COLUMN_NAME) = 'DRYBIOM' AND " +
                                                   "a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION " +
                                          "SELECT a.COLUMN_NAME,'NULLS NOT ALLOWED' AS ERROR_DESC, FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE FVS_TREE_ID IS NULL OR LEN(TRIM(FVS_TREE_ID))=0) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                  "a.NOVALUE_ERROR <> 'NA' AND " +
                                                  "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                  "TRIM(a.COLUMN_NAME) = 'FVS_TREE_ID' AND " +
                                                  "a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION " +
                                          "SELECT a.COLUMN_NAME,'NULLS NOT ALLOWED' AS ERROR_DESC, FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE FVSCREATEDTREE_YN IS NULL OR LEN(TRIM(FVSCREATEDTREE_YN))=0) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                   "a.NOVALUE_ERROR <> 'NA' AND " +
                                                   "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                   "TRIM(a.COLUMN_NAME) = 'FVSCREATEDTREE_YN' AND " +
                                                  "a.RXPACKAGE='" + p_strRxPackage + "' " +
                                          "UNION " +
                                          "SELECT a.COLUMN_NAME,'NULLS NOT ALLOWED' AS ERROR_DESC, FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " WHERE FVS_SPECIES IS NULL OR LEN(TRIM(FVS_SPECIES))=0) FVS_TREE " +
                                             "WHERE a.NOVALUE_ERROR IS NOT NULL AND " +
                                                   "LEN(TRIM(NOVALUE_ERROR)) > 0 AND " +
                                                   "a.NOVALUE_ERROR <> 'NA' AND " +
                                                   "VAL(a.NOVALUE_ERROR) > 0 AND " +
                                                   "TRIM(a.COLUMN_NAME) = 'FVS_SPECIES' AND " +
                                                   "a.RXPACKAGE='" + p_strRxPackage + "')";

                return sqlArray;

            }
            /// <summary>
            /// SQL for post-processing audit of the BIOSUMCALC\FVS_TREE tables. Find required columns with no values.
            /// </summary>
            /// <param name="p_strInsertTable"></param>
            /// <param name="p_strPostAuditSummaryTable"></param>
            /// <param name="p_strFvsTreeTableName"></param>
            /// <param name="p_strFVSTreeFileName"></param>
            /// <returns></returns>
            public static string[] FVSOutputTable_AuditPostSummaryDetailFVS_VALUE_ERROR(string p_strInsertTable, string p_strPostAuditSummaryTable, string p_strFvsTreeTableName, string p_strRxPackage)
            {

                string[] sqlArray = new string[2];

                sqlArray[0] = "DELETE FROM " + p_strInsertTable + " WHERE RXPACKAGE = '" + p_strRxPackage + "'";

                sqlArray[1] = "INSERT INTO  " + p_strInsertTable + " " +
                                    "SELECT * FROM " +
                                         "(SELECT a.COLUMN_NAME,FVS_TREE.RXCYCLE + ' is not a valid cycle. Valid values are 1,2,3 or 4' AS ERROR_DESC, FVS_TREE.* FROM " + p_strPostAuditSummaryTable + " a," +
                                             "(SELECT * " + 
	                                          "FROM " + p_strFvsTreeTableName + " " +
	                                          "WHERE RXCYCLE IS NOT NULL AND LEN(TRIM(RXCYCLE)) > 0 AND " +
	                                                "RXCYCLE NOT IN ('1','2','3','4')) FVS_TREE " +
                                          "WHERE a.VALUE_ERROR IS NOT NULL AND " +
                                                "LEN(TRIM(VALUE_ERROR)) > 0 AND " +
                                                "a.VALUE_ERROR <> 'NA' AND " +
                                                "VAL(a.VALUE_ERROR) > 0 AND " +
                                                "TRIM(a.COLUMN_NAME) = 'RXCYCLE' AND " + 
                                                "a.rxPackage='" + p_strRxPackage + "')";

                return sqlArray;

            }

            /// <summary>
            /// SQL for post-processing audit of the BIOSUMCALC\FVS_TREE tables. List foreign key columns whose values are not found in the foreign tables.
            /// </summary>
            /// <param name="p_strInsertTable"></param>
            /// <param name="p_strPostAuditSummaryTable"></param>
            /// <param name="p_strFvsTreeTableName"></param>
            /// <param name="p_strFVSTreeFileName"></param>
            /// <returns></returns>
            public static string[] FVSOutputTable_AuditPostSummaryDetailFVS_NOTFOUND_ERROR(
                string p_strInsertTable, 
                string p_strPostAuditSummaryTable, 
                string p_strFvsTreeTableName, 
                string p_strRxPackage,
                string p_strCondTable,
                string p_strPlotTable, 
                string p_strTreeTable,
                string p_strRxTable,
                string p_strRxPackageTable,
                string p_strRxPackageWorkTable)
            {

                string[] sqlArray = new string[2];

                sqlArray[0] = "DELETE FROM " + p_strInsertTable + " WHERE rxpackage = '" + p_strRxPackage + "'";

                sqlArray[1] = "INSERT INTO  " + p_strInsertTable + " " +
                                    "SELECT * FROM " +
                                    "(SELECT DISTINCT " +
                                        "'BIOSUM_COND_ID' AS COLUMN_NAME," +
                                        "a.BIOSUM_COND_ID AS NOTFOUND_VALUE," +
                                        "'NOT FOUND IN COND TABLE' AS ERROR_DESC," +
                                        "a.* FROM " + p_strFvsTreeTableName + " a," +
                                            "(SELECT * FROM fvs_tree_biosum_plot_id_work_table a " +
                                             "WHERE a.BIOSUM_COND_ID IS NOT NULL AND LEN(TRIM(a.BIOSUM_COND_ID)) >  0 AND " +
                                                   "NOT EXISTS (SELECT b.BIOSUM_COND_ID FROM cond_biosum_cond_id_work_table  b " +
                                                               "WHERE a.BIOSUM_COND_ID = b.BIOSUM_COND_ID)) biosum_cond_id_not_found " +
                                     "WHERE a.ID = biosum_cond_id_not_found.ID " +
                                     "UNION " +
                                     "SELECT DISTINCT " +
                                       "'BIOSUM_PLOT_ID' AS COLUMN_NAME," +
                                       "biosum_cond_id_not_found_in_plot_table.BIOSUM_PLOT_ID AS NOTFOUND_VALUE," +
                                       "'NOT FOUND IN PLOT TABLE' AS ERROR_DESC," +
                                       "a.* FROM " + p_strFvsTreeTableName + " a," +
                                            "(SELECT * FROM fvs_tree_biosum_plot_id_work_table a " +
                                             "WHERE a.BIOSUM_COND_ID IS NOT NULL AND " +
                                                   "LEN(TRIM(a.BIOSUM_COND_ID)) >  0 AND " +
                                                   "NOT EXISTS (SELECT b.BIOSUM_PLOT_ID FROM plot_biosum_plot_id_work_table  b " +
                                                               "WHERE a.BIOSUM_PLOT_ID = b.BIOSUM_PLOT_ID)) biosum_cond_id_not_found_in_plot_table " +
                                     "WHERE a.ID = biosum_cond_id_not_found_in_plot_table.ID " +
                                     "UNION " +
                                     "SELECT DISTINCT " +
                                       "'FVS_TREE_ID' AS COLUMN_NAME," +
                                       "a.FVS_TREE_ID AS NOTFOUND_VALUE," +
                                       "'NOT FOUND IN TREE TABLE' AS ERROR_DESC," +
                                       "a.* FROM " + p_strFvsTreeTableName + " a," +
                                        "(SELECT * FROM " + p_strFvsTreeTableName + " a " +
                                         "WHERE a.FvsCreatedTree_YN='N' AND " +
                                               "a.FVS_TREE_ID IS NOT NULL AND LEN(TRIM(a.FVS_TREE_ID)) >  0 AND " +
                                               "NOT EXISTS (SELECT b.FVS_TREE_ID FROM tree_fvs_tree_id_work_table b " +
                                                           "WHERE a.fvs_tree_id = b.fvs_tree_id and a.biosum_cond_id = b.biosum_cond_id)) " +
                                                           "fvs_tree_id_not_found_in_tree_table " +
                                     "WHERE a.ID = fvs_tree_id_not_found_in_tree_table.ID " +
                                     "UNION " +
                                     "SELECT DISTINCT " +
                                        "'RX' AS COLUMN_NAME," +
                                        "a.RX AS NOTFOUND_VALUE," +
                                        "'NOT FOUND IN RX TABLE' AS ERROR_DESC," +
                                        "a.* FROM " + p_strFvsTreeTableName + " a " +
                                     "WHERE a.RX NOT IN (SELECT b.RX FROM " + p_strRxTable + " b) " +
                                     "UNION " +
                                     "SELECT DISTINCT " +
                                       "'RXPACKAGE' AS COLUMN_NAME," +
                                       "a.RXPACKAGE AS NOTFOUND_VALUE," +
                                       "'NOT FOUND IN RXPACKAGE TABLE' AS ERROR_DESC," +
                                       "a.* FROM " + p_strFvsTreeTableName + " a " +
                                     "WHERE a.RXPACKAGE NOT IN (SELECT b.RXPACKAGE FROM " + p_strRxPackageTable + " b) " +
                                     "UNION " +
                                     "SELECT DISTINCT " +
                                        "'RXPACKAGE + RXCYCLE + RX' AS COLUMN_NAME," +
                                        "'RXPACKAGE=' + a.RXPACKAGE + ' RXCYCLE=' + a.RXCYCLE + ' RX=' + a.RX  AS NOTFOUND_VALUE," +
                                        "'COMBINATION OF RXPACKAGE, RXCYCLE, AND RX NOT FOUND' AS ERROR_DESC," +
                                        "a.* FROM " + p_strFvsTreeTableName + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " a " +
                                             "WHERE a.RX IS NOT NULL AND LEN(TRIM(a.RX)) >  0 AND " +
                                                   "a.RXPACKAGE IS NOT NULL AND LEN(TRIM(a.RXPACKAGE)) >  0 AND  " +
                                                   "a.RXCYCLE IS NOT NULL AND LEN(TRIM(a.RXCYCLE)) >  0 AND " +
                                                   "NOT EXISTS (SELECT rxp.RX FROM " + p_strRxPackageWorkTable + " rxp " +
                                                               "WHERE TRIM(a.rxpackage) = TRIM(rxp.rxpackage) AND " +
                                                                     "TRIM(a.rxcycle)=TRIM(rxp.rxcycle) AND  " +
                                                                     "TRIM(a.rx)=TRIM(rxp.rx))) not_found_rxpackage_rxcycle_rx_combo " +
                                     "WHERE a.ID = not_found_rxpackage_rxcycle_rx_combo.ID)";

                return sqlArray;

            }
            public static string[] AppendRuntitleToFVSOut(string strTable)
            {
                string[] sqlArray = new string[3];
                sqlArray[0] = $@"alter table {strTable} add column RUNTITLE VARCHAR(255)";
                sqlArray[1] = $@"create index idx_{strTable}_runtitle on {strTable} (RUNTITLE)";
                sqlArray[2] = $@"UPDATE {strTable} SET RUNTITLE = (SELECT RUNTITLE FROM FVS_CASES WHERE CASEID = {strTable}.CASEID)";
                return sqlArray;
            }
            /*
            /// <summary>
            /// SQL for post-processing audit of the BIOSUMCALC\FVS_TREE tables. List foreign key columns whose values are not found in the foreign tables.
            /// </summary>
            /// <param name="p_strInsertTable"></param>
            /// <param name="p_strPostAuditSummaryTable"></param>
            /// <param name="p_strFvsTreeTableName"></param>
            /// <param name="p_strFVSTreeFileName"></param>
            /// <returns></returns>
            public static string[] FVSOutputTable_AuditPostSummaryDetailFVS_VALUE_ERROR(
                string p_strInsertTable,
                string p_strPostAuditSummaryTable,
                string p_strFvsTreeTableName,
                string p_strFVSTreeFileName,
                string p_strCondTable,
                string p_strPlotTable,
                string p_strTreeTable,
                string p_strRxTable,
                string p_strRxPackageTable,
                string p_strRxPackageWorkTable)
            {

                string[] sqlArray = new string[2];

                sqlArray[0] = "DELETE FROM " + p_strInsertTable + " WHERE TRIM(UCASE(FVS_TREE_FILE)) = '" + p_strFVSTreeFileName.ToUpper().Trim() + "'";

                sqlArray[1] = "INSERT INTO  " + p_strInsertTable + " " +
                                    "SELECT * FROM " +
                                    "(SELECT DISTINCT " +
                                       "'" + p_strFVSTreeFileName + "' AS FVS_TREE_FILE," +
                                        "'BIOSUM_COND_ID' AS COLUMN_NAME," +
                                        "a.BIOSUM_COND_ID AS NOTFOUND_VALUE," +
                                        "'NOT FOUND IN COND TABLE' AS ERROR_DESC," +
                                        "a.* FROM " + p_strFvsTreeTableName + " a," +
                                            "(SELECT * FROM fvs_tree_biosum_plot_id_work_table a " +
                                             "WHERE a.BIOSUM_COND_ID IS NOT NULL AND LEN(TRIM(a.BIOSUM_COND_ID)) >  0 AND " +
                                                   "NOT EXISTS (SELECT b.BIOSUM_COND_ID FROM cond_biosum_cond_id_work_table  b " +
                                                               "WHERE a.BIOSUM_COND_ID = b.BIOSUM_COND_ID)) biosum_cond_id_not_found " +
                                     "WHERE a.ID = biosum_cond_id_not_found.ID " +
                                     "UNION " +
                                     "SELECT DISTINCT " +
                                       "'" + p_strFVSTreeFileName + "' AS FVS_TREE_FILE," +
                                       "'BIOSUM_PLOT_ID' AS COLUMN_NAME," +
                                       "biosum_cond_id_not_found_in_plot_table.BIOSUM_PLOT_ID AS NOTFOUND_VALUE," +
                                       "'NOT FOUND IN PLOT TABLE' AS ERROR_DESC," +
                                       "a.* FROM " + p_strFvsTreeTableName + " a," +
                                            "(SELECT * FROM fvs_tree_biosum_plot_id_work_table a " +
                                             "WHERE a.BIOSUM_COND_ID IS NOT NULL AND " +
                                                   "LEN(TRIM(a.BIOSUM_COND_ID)) >  0 AND " +
                                                   "NOT EXISTS (SELECT b.BIOSUM_PLOT_ID FROM plot_biosum_plot_id_work_table  b " +
                                                               "WHERE a.BIOSUM_PLOT_ID = b.BIOSUM_PLOT_ID)) biosum_cond_id_not_found_in_plot_table " +
                                     "WHERE a.ID = biosum_cond_id_not_found_in_plot_table.ID " +
                                     "UNION " +
                                     "SELECT DISTINCT " +
                                       "'" + p_strFVSTreeFileName + "' AS FVS_TREE_FILE," +
                                       "'FVS_TREE_ID' AS COLUMN_NAME," +
                                       "a.FVS_TREE_ID AS NOTFOUND_VALUE," +
                                       "'NOT FOUND IN TREE TABLE' AS ERROR_DESC," +
                                       "a.* FROM " + p_strFvsTreeTableName + " a," +
                                        "(SELECT * FROM " + p_strFvsTreeTableName + " a " +
                                         "WHERE a.FvsCreatedTree_YN='N' AND " +
                                               "a.FVS_TREE_ID IS NOT NULL AND LEN(TRIM(a.FVS_TREE_ID)) >  0 AND " +
                                               "NOT EXISTS (SELECT b.FVS_TREE_ID FROM tree_fvs_tree_id_work_table b " +
                                                           "WHERE a.fvs_tree_id = b.fvs_tree_id and a.biosum_cond_id = b.biosum_cond_id)) fvs_tree_id_not_found_in_tree_table " +
                                     "WHERE a.ID = fvs_tree_id_not_found_in_tree_table.ID " +
                                     "UNION " +
                                     "SELECT DISTINCT " +
                                        "'" + p_strFVSTreeFileName + "' AS FVS_TREE_FILE," +
                                        "'RX' AS COLUMN_NAME," +
                                        "a.RX AS NOTFOUND_VALUE," +
                                        "'NOT FOUND IN RX TABLE' AS ERROR_DESC," +
                                        "a.* FROM " + p_strFvsTreeTableName + " a " +
                                     "WHERE a.RX NOT IN (SELECT b.RX FROM " + p_strRxTable + " b) " +
                                     "UNION " +
                                     "SELECT DISTINCT " +
                                       "'" + p_strFVSTreeFileName + "' AS FVS_TREE_FILE," +
                                       "'RXPACKAGE' AS COLUMN_NAME," +
                                       "a.RXPACKAGE AS NOTFOUND_VALUE," +
                                       "'NOT FOUND IN RXPACKAGE TABLE' AS ERROR_DESC," +
                                       "a.* FROM " + p_strFvsTreeTableName + " a " +
                                     "WHERE a.RXPACKAGE NOT IN (SELECT b.RXPACKAGE FROM " + p_strRxPackageTable + " b) " +
                                     "UNION " +
                                     "SELECT DISTINCT " +
                                        "'" + p_strFVSTreeFileName + "' AS FVS_TREE_FILE," +
                                        "'RXPACKAGE + RXCYCLE + RX' AS COLUMN_NAME," +
                                        "'RXPACKAGE=' + a.RXPACKAGE + ' RXCYCLE=' + a.RXCYCLE + ' RX=' + a.RX  AS NOTFOUND_VALUE," +
                                        "'COMBINATION OF RXPACKAGE, RXCYCLE, AND RX NOT FOUND' AS ERROR_DESC," +
                                        "a.* FROM " + p_strFvsTreeTableName + " a," +
                                            "(SELECT * FROM " + p_strFvsTreeTableName + " a " +
                                             "WHERE a.RX IS NOT NULL AND LEN(TRIM(a.RX)) >  0 AND " +
                                                   "a.RXPACKAGE IS NOT NULL AND LEN(TRIM(a.RXPACKAGE)) >  0 AND  " +
                                                   "a.RXCYCLE IS NOT NULL AND LEN(TRIM(a.RXCYCLE)) >  0 AND " +
                                                   "NOT EXISTS (SELECT rxp.RX FROM " + p_strRxPackageWorkTable + " rxp " +
                                                               "WHERE TRIM(a.rxpackage) = TRIM(rxp.rxpackage) AND " +
                                                                     "TRIM(a.rxcycle)=TRIM(rxp.rxcycle) AND  " +
                                                                     "TRIM(a.rx)=TRIM(rxp.rx))) not_found_rxpackage_rxcycle_rx_combo " +
                                     "WHERE a.ID = not_found_rxpackage_rxcycle_rx_combo.ID)";

                return sqlArray;

            }
            */

            public class FVSInput
            {
		        FVSInput()
		        {
		        }

		        //All the queries necessary to create the FVSIn.accdb FVS_StandInit table using intermediate tables
		        public class StandInit
		        {
		            StandInit()
		            {
		            }

		            public static string CreateDWMFuelbedTypCdToFVSConversionTable()
		            {
		                return "CREATE TABLE Ref_DWM_Fuelbed_Type_Codes (dwm_fuelbed_typcd TEXT(3), fuel_model LONG);";
		            }

		            public static string BulkImportStandDataFromBioSumMaster(string strVariant, string strDestTable,
		                string strCondTableName, string strPlotTableName)
		            {
		                string strInsertIntoStandInit =
		                    "INSERT INTO " + strDestTable +
		                    " (Stand_ID, Variant, Inv_Year, Latitude, Longitude, Location, PV_Code, " +
		                    "Age, Aspect, Slope, ElevFt, Basal_Area_Factor, Inv_Plot_Size, Brk_DBH, " +
		                    "Num_Plots, NonStk_Plots, Sam_Wt, Stk_Pcnt, DG_Trans, DG_Measure, " +
		                    "HTG_Trans, HTG_Measure, Mort_Measure, Forest_Type, State, County) ";
		                string strBioSumWorkTableSelectStmt =
		                    "SELECT c.biosum_cond_id, p.fvs_variant, p.measyear, p.lat, p.lon, " +
		                    "p.fvsloccode, c.habtypcd1, c.stdage, c.aspect, c.slope, p.elev, 0, 1, 999, 1, 0, 1,  " +
		                    "iif(c.landclcd is null, 0, c.landclcd), 1, 10, 1, 5, 5, c.fortypcd, p.statecd, p.countycd ";
		                string strFromTableExpr =
		                    String.Format(
		                        "FROM {0} c INNER JOIN {1} p ON c.biosum_plot_id = p.biosum_plot_id ",
		                        strCondTableName, strPlotTableName);
		                string strFilters = "WHERE c.landclcd = 1 AND ucase(trim(p.fvs_variant)) = \'" +
		                                    strVariant.Trim().ToUpper() + "\'";
		                return strInsertIntoStandInit + strBioSumWorkTableSelectStmt + strFromTableExpr +
		                       strFilters;
		            }

		            public static string UpdateFuelModel(string strDestTable, string strCondTableName)
		            {
		                return String.Format(
		                    "UPDATE ({0} fvs INNER JOIN {1} c ON fvs.Stand_ID=c.biosum_cond_id) " +
		                    "LEFT JOIN {2} ref ON c.dwm_fuelbed_typcd=ref.dwm_fuelbed_typcd " +
		                    "SET fvs.Fuel_Model=IIF(c.dwm_fuelbed_typcd IS NULL, NULL, ref.Fuel_Model);",
		                    strDestTable, strCondTableName, "Ref_DWM_Fuelbed_Type_Codes");
		            }

		            /// <summary>
		            /// Calculate biomass contribution to condition per CWD piece so that they can be summed
		            /// ScalingConstants * SpeciesBulkDensity * SpeciesDecayRatio * PieceTransectDiameter^2 / TotalTransectLengthForCondition
		            /// </summary>
		            /// <param name="strCoarseWoodyDebrisTable"></param>
		            /// <param name="strCondTable"></param>
		            /// <param name="strRefSpecies"></param>
		            /// <param name="strTransectSegmentTable"></param>
		            /// <returns></returns>
		            public static string[] CalculateCoarseWoodyDebrisBiomassTonsPerAcre(string strCoarseWoodyDebrisTable,
		                string strCondTable, string strRefSpecies, string strTransectSegmentTable, string strMinCwdTransectLengthTotal)
		            {
		                string[] strCwdSqlStmts = new string[11];
		                int idx = 0;

		                //Sum horizontal lengths for transects
		                strCwdSqlStmts[idx++] = String.Format(
		                    "SELECT biosum_cond_id, SUM(horiz_length) as CWDTotalLength " +
		                    "INTO CWDTotalLengthWorkTable " +
		                    "FROM {0} " +
		                    "GROUP BY biosum_cond_id;", strTransectSegmentTable);

		                //Create worktable for calculation of cwd piece biomass
		                strCwdSqlStmts[idx++] = String.Format("SELECT cwd.biosum_cond_id, " +
		                                                      "cwd.spcd, " +
		                                                      "cwd.transdia, " +
		                                                      "cwd.inclination, " +
		                                                      "cwd.decaycd, " +
		                                                      "CDbl(0) as BulkDensity, " +
		                                                      "CDbl(0) as DecayRatio, " +
		                                                      "tl.CWDTotalLength, " +
		                                                      "CDbl(0) as Biomass, " +
		                                                      //from NRS-22 Eqn 3.27
		                                                      //.1865972082080956863 = (1.0/2000.0)*(43560.0/144.0)*(3.141592654^2/8.0)
		                                                      "0.186597208 as ScalingConstants " +
		                                                      "INTO CwdPieceWorkTable " +
		                                                      "FROM {0} cwd " +
		                                                      "INNER JOIN CWDTotalLengthWorkTable tl " +
		                                                      "ON cwd.biosum_cond_id=tl.biosum_cond_id;",
		                    strCoarseWoodyDebrisTable);

		                //Set BulkDensity And DecayRatios for each piece for easy multiplication across row
		                strCwdSqlStmts[idx++] = String.Format("UPDATE CwdPieceWorkTable cwd " +
		                                                      "INNER JOIN {0} spc ON cwd.spcd=spc.spcd " +
		                                                      "SET cwd.BulkDensity=spc.CWD_Bulk_Density," +
		                                                      "cwd.DecayRatio=IIF(cwd.decaycd=1, spc.CWD_Decay_Ratio1, " +
		                                                      "IIF(cwd.decaycd=2, spc.CWD_Decay_Ratio2, " +
		                                                      "IIF(cwd.decaycd=3, spc.CWD_Decay_Ratio3, " +
		                                                      "IIF(cwd.decaycd=4, spc.CWD_Decay_Ratio4, " +
		                                                      "IIF(cwd.decaycd=5, spc.CWD_Decay_Ratio5, 0)))));",
		                    strRefSpecies);

		                //TODO: null spcd needs bulkdensity and decayratio
		                strCwdSqlStmts[idx++] = String.Format(
		                    "UPDATE CwdPieceWorkTable cwd, (SELECT * FROM {0} WHERE spcd=998) spc " +
		                    "SET cwd.BulkDensity=spc.CWD_Bulk_Density," +
		                    "cwd.DecayRatio=IIF(cwd.decaycd=1, spc.CWD_Decay_Ratio1, " +
		                    "IIF(cwd.decaycd=2, spc.CWD_Decay_Ratio2, " +
		                    "IIF(cwd.decaycd=3, spc.CWD_Decay_Ratio3, " +
		                    "IIF(cwd.decaycd=4, spc.CWD_Decay_Ratio4, " +
		                    "IIF(cwd.decaycd=5, spc.CWD_Decay_Ratio5, 0))))) " +
		                    "WHERE cwd.spcd IS NULL;",
		                    strRefSpecies);


		                //Calculate the CWD piece's biomass contribution to the condition
		                //The last step is to group these by biosum_cond_id and sum them
		                strCwdSqlStmts[idx++] = "UPDATE CwdPieceWorkTable cwd " +
		                                        "SET cwd.Biomass=cwd.ScalingConstants*cwd.BulkDensity*cwd.DecayRatio*cwd.transdia^2/cwd.CWDTotalLength;";


                        //Update CWD piece's biomass: scale by cosine_inclination or 1 if inclination is null 
                        //Piece inclination is measured in degrees from 0 to 90. Convert to radians for Access cosine function
		                strCwdSqlStmts[idx++] = "UPDATE CwdPieceWorkTable cwd " +
		                                        "SET cwd.Biomass=cwd.Biomass/COS(cwd.inclination * 3.141592654/180)" +
		                                        "WHERE cwd.inclination IS NOT NULL;";

		                //Sum by size and decay bins to match FVS_StandInit columns
		                strCwdSqlStmts[idx++] = "SELECT cwd.biosum_cond_id, " +
		                                        "SUM(IIF((cwd.transdia >= 3 AND cwd.transdia < 6) AND (cwd.decaycd IN (1,2,3)), " +
		                                        "cwd.Biomass, 0)) as fuel_3_6_H, " +
		                                        "SUM(IIF((cwd.transdia >= 6 AND cwd.transdia < 12) AND (cwd.decaycd IN (1,2,3)), " +
		                                        "cwd.Biomass, 0)) as fuel_6_12_H, " +
		                                        "SUM(IIF((cwd.transdia >= 12 AND cwd.transdia < 20) AND (cwd.decaycd IN (1,2,3)), " +
		                                        "cwd.Biomass, 0)) as fuel_12_20_H, " +
		                                        "SUM(IIF((cwd.transdia >= 20 AND cwd.transdia < 35) AND (cwd.decaycd IN (1,2,3)), " +
		                                        "cwd.Biomass, 0)) as fuel_20_35_H, " +
		                                        "SUM(IIF((cwd.transdia >= 35 AND cwd.transdia < 50) AND (cwd.decaycd IN (1,2,3)), " +
		                                        "cwd.Biomass, 0)) as fuel_35_50_H, " +
		                                        "SUM(IIF((cwd.transdia >= 50 AND cwd.transdia < 999) AND (cwd.decaycd IN (1,2,3)), " +
		                                        "cwd.Biomass, 0)) as fuel_gt_50_H, " +
		                                        "SUM(IIF((cwd.transdia >= 3 AND cwd.transdia < 6) AND (cwd.decaycd IN (4,5)), " +
		                                        "cwd.Biomass, 0)) as fuel_3_6_S, " +
		                                        "SUM(IIF((cwd.transdia >= 6 AND cwd.transdia < 12) AND (cwd.decaycd IN (4,5)), " +
		                                        "cwd.Biomass, 0)) as fuel_6_12_S, " +
		                                        "SUM(IIF((cwd.transdia >= 12 AND cwd.transdia < 20) AND (cwd.decaycd IN (4,5)), " +
		                                        "cwd.Biomass, 0)) as fuel_12_20_S, " +
		                                        "SUM(IIF((cwd.transdia >= 20 AND cwd.transdia < 35) AND (cwd.decaycd IN (4,5)), " +
		                                        "cwd.Biomass, 0)) as fuel_20_35_S, " +
		                                        "SUM(IIF((cwd.transdia >= 35 AND cwd.transdia < 50) AND (cwd.decaycd IN (4,5)), " +
		                                        "cwd.Biomass, 0)) as fuel_35_50_S, " +
		                                        "SUM(IIF((cwd.transdia >= 50 AND cwd.transdia < 999) AND (cwd.decaycd IN (4,5)), " +
		                                        "cwd.Biomass, 0)) as fuel_gt_50_S " +
		                                        "INTO DWM_CWD_Aggregates_WorkTable " +
		                                        "FROM CwdPieceWorkTable cwd " +
		                                        "GROUP BY cwd.biosum_cond_id;";

		                //Update CWD columns in FVS_StandInit
		                strCwdSqlStmts[idx++] = "UPDATE FVS_StandInit_WorkTable fvs " +
		                                        "INNER JOIN DWM_CWD_Aggregates_WorkTable cwd ON fvs.Stand_ID=cwd.biosum_cond_id " +
		                                        "SET fvs.fuel_3_6_H=cwd.fuel_3_6_H, " +
		                                        "fvs.fuel_6_12_H=cwd.fuel_6_12_H, " +
		                                        "fvs.fuel_12_20_H=cwd.fuel_12_20_H, " +
		                                        "fvs.fuel_20_35_H=cwd.fuel_20_35_H, " +
		                                        "fvs.fuel_35_50_H=cwd.fuel_35_50_H, " +
		                                        "fvs.fuel_gt_50_H=cwd.fuel_gt_50_H, " +
		                                        "fvs.fuel_3_6_S=cwd.fuel_3_6_S, " +
		                                        "fvs.fuel_6_12_S=cwd.fuel_6_12_S, " +
		                                        "fvs.fuel_12_20_S=cwd.fuel_12_20_S, " +
		                                        "fvs.fuel_20_35_S=cwd.fuel_20_35_S, " +
		                                        "fvs.fuel_35_50_S=cwd.fuel_35_50_S, " +
		                                        "fvs.fuel_gt_50_S=cwd.fuel_gt_50_S;";

		                //Write them to fvs_standinit column
		                strCwdSqlStmts[idx++] = "UPDATE FVS_StandInit_WorkTable fvs " +
		                                        "INNER JOIN CWDTotalLengthWorkTable t ON fvs.stand_id=t.biosum_cond_id " +
		                                        "SET fvs.cwdtotallength=t.CWDTotalLength;";
		                
		                //Nullify CWD where total horizontal length is less than the minimum
		                strCwdSqlStmts[idx++] = String.Format("UPDATE FVS_StandInit_WorkTable fvs " +
		                                        "INNER JOIN CWDTotalLengthWorkTable t ON fvs.stand_id=t.biosum_cond_id " +
		                                        "SET fvs.fuel_3_6_H=NULL, " +
		                                        "fvs.fuel_6_12_H=NULL, " +
		                                        "fvs.fuel_12_20_H=NULL, " +
		                                        "fvs.fuel_20_35_H=NULL, " +
		                                        "fvs.fuel_35_50_H=NULL, " +
		                                        "fvs.fuel_gt_50_H=NULL, " +
		                                        "fvs.fuel_3_6_S=NULL, " +
		                                        "fvs.fuel_6_12_S=NULL, " +
		                                        "fvs.fuel_12_20_S=NULL, " +
		                                        "fvs.fuel_20_35_S=NULL, " +
		                                        "fvs.fuel_35_50_S=NULL, " +
		                                        "fvs.fuel_gt_50_S=NULL " + 
		                                        "WHERE fvs.CwdTotalLength < {0};", strMinCwdTransectLengthTotal);

		                //Set cwd fuels to 0 if TS.horiz_length is non-null and positive
		                //FVS interprets 0s as 0 tons/acre, but null lets it insert default values for the variant.
                        strCwdSqlStmts[idx++] = String.Format("UPDATE FVS_StandInit_WorkTable " +
                                                              "SET Fuel_3_6_H=IIF(Fuel_3_6_H is null, 0, Fuel_3_6_H), " +
                                                              "Fuel_6_12_H=IIF(Fuel_6_12_H is null, 0, Fuel_6_12_H), " +
                                                              "Fuel_12_20_H=IIF(Fuel_12_20_H is null, 0, Fuel_12_20_H), " +
                                                              "Fuel_20_35_H=IIF(Fuel_20_35_H is null, 0, Fuel_20_35_H), " +
                                                              "Fuel_35_50_H=IIF(Fuel_35_50_H is null, 0, Fuel_35_50_H), " +
                                                              "Fuel_gt_50_H=IIF(Fuel_gt_50_H is null, 0, Fuel_gt_50_H), " +
                                                              "Fuel_3_6_S=IIF(Fuel_3_6_S is null, 0, Fuel_3_6_S), " +
                                                              "Fuel_6_12_S=IIF(Fuel_6_12_S is null, 0, Fuel_6_12_S), " +
                                                              "Fuel_12_20_S=IIF(Fuel_12_20_S is null, 0, Fuel_12_20_S), " +
                                                              "Fuel_20_35_S=IIF(Fuel_20_35_S is null, 0, Fuel_20_35_S), " +
                                                              "Fuel_35_50_S=IIF(Fuel_35_50_S is null, 0, Fuel_35_50_S), " +
                                                              "Fuel_gt_50_S=IIF(Fuel_gt_50_S is null, 0, Fuel_gt_50_S) " +
                                                              "WHERE CWDTotalLength > 0;");

		                return strCwdSqlStmts;
		            }

                    /// <summary>
                    /// Calculate biomass for FWD columns in FVS_StandInit_WorkTable. 
                    /// </summary>
                    /// <param name="strDwmFineWoodyDebrisTable"></param>
                    /// <param name="strCondTable"></param>
                    /// <param name="strRefForestTypeTable"></param>
                    /// <param name="strRefForestTypeGroupTable"></param>
                    /// <param name="strMinSmallFwdTransectLengthTotal"></param>
                    /// <param name="strMinLargeFwdTransectLengthTotal"></param>
                    /// <returns></returns>
		            public static string[] CalculateFineWoodyDebrisBiomassTonsPerAcre(string strDwmFineWoodyDebrisTable,
		                string strCondTable, string strMinSmallFwdTransectLengthTotal, string strMinLargeFwdTransectLengthTotal)
		            {
		                string[] strFwdSqlStmts = new string[7];
		                int idx = 0;

                        //Gather variables for NRS-22 FWD equation into temporary table
		                strFwdSqlStmts[idx++] = String.Format(
		                    "SELECT fwd.biosum_cond_id, first(fwd.rsnctcd) as rsncntcd, " +
		                    "sum(fwd.smallct) as smallTotal, first(fwd.small_tl_cond) as smallTL, " +
		                    "sum(fwd.mediumct) as mediumTotal, first(fwd.medium_tl_cond) as mediumTL, " +
		                    "sum(fwd.largect) as largeTotal, first(fwd.large_tl_cond) as largeTL, " +
		                    "first(rftg.fwd_decay_ratio) as decayRatio, first(rftg.fwd_density) as bulkDensity, " +
		                    "first(rftg.fwd_small_qmd) as smallQMD,  " +
		                    "first(rftg.fwd_medium_qmd) as mediumQMD,  " +
		                    "first(rftg.fwd_large_qmd) as largeQMD, " +
		                    "first(small_tl_cond) as SmallMediumTotalLength, " +
		                    "first(large_tl_cond) as LargeTotalLength, " +
		                    "0.186597208 as ScalingConstants " +
		                    "INTO DWM_FWD_Aggregates_WorkTable " +
		                    "FROM (({0} fwd INNER JOIN {1} c ON fwd.biosum_cond_id = c.biosum_cond_id) " +
		                    "INNER JOIN REF_FOREST_TYPE rft ON c.fortypcd = rft.`VALUE`) " +
		                    "INNER JOIN REF_FOREST_TYPE_GROUP rftg ON rft.TYPGRPCD = rftg.`VALUE` " +
		                    "WHERE fwd.rsnctcd < 2 OR fwd.rsnctcd IS NULL " +
		                    "GROUP BY fwd.biosum_cond_id;", strDwmFineWoodyDebrisTable, strCondTable);


                        //Creates a table for scaling FWD TLs because the counts (small, medium, large) are filtered by RsnCtCd
		                strFwdSqlStmts[idx++] = String.Format("SELECT fwd.biosum_cond_id, " +
		                                                      "COUNT(iif(rsnctcd < 2 OR rsnctcd IS NULL, true, null)) as totalAcceptabelRsnCtCd, " +
		                                                      "COUNT(*) as totalFWDRecords, " +
		                                                      "CDbl(totalAcceptabelRsnCtCd)/totalFWDRecords as RsnCtCdScalingFactor " +
		                                                      "INTO Dwm_Fwd_RsnCtCd_WorkTable " +
		                                                      "FROM {0} fwd INNER JOIN FVS_StandInit_WorkTable fvs ON fwd.biosum_cond_id=fvs.stand_id " +
		                                                      "GROUP BY fwd.biosum_cond_id", strDwmFineWoodyDebrisTable);

                        //Multiply the FWD transect lengths for biomass calculations by the proportion of TLs with accepted RsnCtCds.
                        //These RsnCtCd values are currently filtered out (2,3,4). Nulls are kept in.
		                strFwdSqlStmts[idx++] = String.Format("UPDATE DWM_FWD_Aggregates_WorkTable fwd " +
		                                                      "INNER JOIN Dwm_Fwd_RsnCtCd_WorkTable rsn " +
		                                                      "ON fwd.biosum_cond_id=rsn.biosum_cond_id " +
		                                                      "SET fwd.smallTL=fwd.smallTL*rsn.RsnCtCdScalingFactor, " +
		                                                      "fwd.mediumTL=fwd.mediumTL*rsn.RsnCtCdScalingFactor, " +
		                                                      "fwd.largeTL=fwd.largeTL*rsn.RsnCtCdScalingFactor, " +
		                                                      "fwd.SmallMediumTotalLength=fwd.SmallMediumTotalLength*rsn.RsnCtCdScalingFactor, " +
		                                                      "fwd.LargeTotalLength=fwd.LargeTotalLength*rsn.RsnCtCdScalingFactor;");


		                strFwdSqlStmts[idx++] = "UPDATE FVS_StandInit_WorkTable fvs " +
		                                        "INNER JOIN DWM_FWD_Aggregates_WorkTable fwd " +
		                                        "ON fvs.stand_id=fwd.biosum_cond_id " +
		                                        "SET fvs.fuel_0_25_H=IIF(smallTL=0 OR smallTL IS NULL, NULL, " +
		                                        "(fwd.ScalingConstants * (fwd.smallTotal * fwd.smallQMD^2) / fwd.smallTL) * fwd.bulkDensity * fwd.decayRatio), " +
		                                        "fvs.fuel_25_1_H=IIF(mediumTL=0 OR mediumTL IS NULL, NULL, " +
		                                        "(fwd.ScalingConstants * (fwd.mediumTotal * fwd.mediumQMD^2) / fwd.mediumTL) * fwd.bulkDensity * fwd.decayRatio), " +
		                                        "fvs.fuel_1_3_H=IIF(largeTL=0 OR largeTL IS NULL, NULL, " +
		                                        "(fwd.ScalingConstants * (fwd.largeTotal * fwd.largeQMD^2) / fwd.largeTL) * fwd.bulkDensity * fwd.decayRatio), " +
		                                        "fvs.fuel_0_25_S=IIF(smallTL=0 OR smallTL IS NULL, NULL, 0), " +
		                                        "fvs.fuel_25_1_S=IIF(mediumTL=0 OR mediumTL IS NULL, NULL, 0), " +
		                                        "fvs.fuel_1_3_S=IIF(largeTL=0 OR largeTL IS NULL, NULL, 0), " +
		                                        "fvs.SmallMediumTotalLength=fwd.SmallMediumTotalLength," +
		                                        "fvs.LargeTotalLength=fwd.LargeTotalLength;";

		                // Set FWD columns to null if the scaling factor determined by all the DWM_Fine_Woody_Debris.RsnCtCd is 0.
		                strFwdSqlStmts[idx++] = "UPDATE FVS_StandInit_WorkTable fvs " +
		                                        "INNER JOIN Dwm_Fwd_RsnCtCd_WorkTable rsn " +
		                                        "ON fvs.stand_id=rsn.biosum_cond_id " +
		                                        "SET fvs.fuel_0_25_H=NULL, " +
		                                        "fvs.fuel_25_1_H=NULL," +
		                                        "fvs.fuel_1_3_H=NULL, " +
		                                        "fvs.fuel_0_25_S=NULL, " +
		                                        "fvs.fuel_25_1_S=NULL, " +
		                                        "fvs.fuel_1_3_S=NULL, " +
		                                        "fvs.SmallMediumTotalLength=NULL, " +
		                                        "fvs.LargeTotalLength=NULL " +
		                                        "WHERE rsn.RsnCtCdScalingFactor=0;";

		                //Filter out FWD Small and Medium fuels where the minimum Small TL is larger
		                strFwdSqlStmts[idx++] = String.Format("UPDATE FVS_StandInit_WorkTable fvs " +
		                                                      "SET fvs.fuel_0_25_H=NULL, " +
		                                                      "fvs.fuel_25_1_H=NULL, " +
		                                                      "fvs.fuel_0_25_S=NULL, " +
		                                                      "fvs.fuel_25_1_S=NULL " +
		                                                      "WHERE fvs.SmallMediumTotalLength < {0} ",
		                    strMinSmallFwdTransectLengthTotal);

		                //Filter out FWD Large fuels where the minimum Large TL is larger
		                strFwdSqlStmts[idx++] = String.Format("UPDATE FVS_StandInit_WorkTable fvs " +
		                                                      "SET fvs.fuel_1_3_H=NULL, " +
		                                                      "fvs.fuel_1_3_S=NULL " +
		                                                      "WHERE fvs.LargeTotalLength < {0} ",
		                    strMinLargeFwdTransectLengthTotal);
		                return strFwdSqlStmts;
		            }

		            public static string CalculateDuffLitterBiomassTonsPerAcre(string strDwmDuffLitterTable, string strCondTable)
		            {
		                return String.Format(
		                    "SELECT dl.biosum_cond_id, " +
		                    "SUM(IIF(duffdep IS NOT NULL, 1, 0)) as DuffPitCount, " +
		                    "SUM(IIF(littdep IS NOT NULL, 1, 0)) as LitterPitCount, " +
		                    "(AVG(dl.duffdep)*first(rftg.duff_density)*1.815) as fvs_fuel_duff_tonsPerAcre, " +
		                    "(AVG(dl.littdep)*first(rftg.litter_density)*1.815) as fvs_fuel_litter_tonsPerAcre " +
		                    "INTO DWM_DuffLitter_Aggregates_WorkTable " +
		                    "FROM ((({0} dl " +
		                    "INNER JOIN {1} c ON dl.biosum_cond_id = c.biosum_cond_id) " +
		                    "INNER JOIN REF_FOREST_TYPE rft ON c.fortypcd = rft.`VALUE`) " +
		                    "INNER JOIN REF_FOREST_TYPE_GROUP rftg ON rft.TYPGRPCD = rftg.`VALUE`) " +
		                    "INNER JOIN FVS_StandInit_WorkTable fvs on dl.biosum_cond_id=fvs.stand_ID " +
		                    "GROUP BY dl.biosum_cond_id;", strDwmDuffLitterTable, strCondTable);
		            }

		            public static string UpdateFvsStandInitDuffLitterColumns()
		            {
		                return "UPDATE Fvs_StandInit_WorkTable fvs " +
		                       "INNER JOIN DWM_DuffLitter_Aggregates_WorkTable dl ON fvs.stand_id=dl.biosum_cond_id  " +
		                       "SET fvs.fuel_duff = dl.fvs_fuel_duff_tonsPerAcre, " +
		                       "fvs.fuel_litter = dl.fvs_fuel_litter_tonsPerAcre," +
		                       "fvs.DuffPitCount = dl.DuffPitCount," +
		                       "fvs.LitterPitCount = dl.LitterPitCount;";
		            }

		            public static string RemoveDuffYears(string yearsFilter)
		            {
		                return "UPDATE Fvs_StandInit_WorkTable fvs " +
		                       "SET fvs.fuel_duff = null," +
		                       "fvs.DuffPitCount=null " +
		                       "WHERE fvs.inv_year IN (" + yearsFilter + ")";
		            }

		            public static string RemoveLitterYears(string yearsFilter)
		            {
		                return "UPDATE Fvs_StandInit_WorkTable fvs " +
		                       "SET fvs.fuel_litter=null," +
		                       "fvs.LitterPitCount=null " +
		                       "WHERE fvs.inv_year IN (" + yearsFilter + ")";
		            }

		            public static string CreateSiteIndexDataset(string strVariant,
		                string strCondTableName, string strPlotTableName)
		            {
		                string strSQL = "SELECT p.biosum_plot_id, c.biosum_cond_id, p.statecd ," +
		                                "p.countycd, p.plot, p.fvs_variant, p.measyear," +
		                                "c.adforcd,p.elev,c.condid, c.habtypcd1," +
		                                "c.stdage,c.slope,c.aspect,c.ground_land_class_pnw," +
		                                "c.sisp,p.lat,p.lon,c.adforcd,c.habtypcd1, " +
		                                "p.elev,c.landclcd,c.ba_ft2_ac,c.habtypcd1 " +
		                                "FROM " + strCondTableName + " c," +
		                                strPlotTableName + " p " +
		                                "WHERE p.biosum_plot_id = c.biosum_plot_id AND " +
		                                "c.landclcd=1 AND " +
		                                "ucase(trim(p.fvs_variant)) = '" + strVariant.Trim().ToUpper() + "';";
		                return strSQL;
		            }

		            public static string TranslateWorkTableToStandInitTable(string strSourceTable, string strDestTable)
		            {
		                string strInsertIntoStandInit =
		                    "INSERT INTO " + strDestTable;
		                string strBioSumWorkTableSelectStmt =
		                    " SELECT Stand_ID, Variant, Inv_Year, " +
		                    "Latitude, Longitude, Region, Forest, District, Compartment, " +
		                    "Location, Ecoregion, PV_Code, PV_Ref_Code, Age, Aspect, Slope, " +
		                    "Elevation, ElevFt, Basal_Area_Factor, Inv_Plot_Size, Brk_DBH, " +
		                    "Num_Plots, NonStk_Plots, Sam_Wt, Stk_Pcnt, DG_Trans, DG_Measure, " +
		                    "HTG_Trans, HTG_Measure, Mort_Measure, Max_BA, Max_SDI, " +
		                    "Site_Species, Site_Index, Model_Type, Physio_Region, Forest_Type, " +
		                    "State, County, Fuel_Model,[Fuel_0_25_H], [Fuel_25_1_H], [Fuel_1_3_H], " +
		                    "[Fuel_3_6_H], [Fuel_6_12_H], [Fuel_12_20_H], Fuel_20_35_H, Fuel_35_50_H, " +
		                    "Fuel_gt_50_H, Fuel_0_25_S, Fuel_25_1_S, Fuel_1_3_S, Fuel_3_6_S, " +
		                    "Fuel_6_12_S, Fuel_12_20_S, Fuel_20_35_S, Fuel_35_50_S, Fuel_gt_50_S, " +
		                    "Fuel_Litter, Fuel_Duff, SmallMediumTotalLength, " +
		                    "LargeTotalLength, CWDTotalLength, DuffPitCount, LitterPitCount, " +
		                    "Photo_Ref, Photo_code ";
		                string strFromTableExpr = "FROM " + strSourceTable + ";";
		                return strInsertIntoStandInit + strBioSumWorkTableSelectStmt + strFromTableExpr;
		            }

		            public static string InsertSiteIndexSpeciesRow(string strStandID, string strSiteSpecies,
		                string strSiteIndex)
		            {
		                return String.Format(
		                    "UPDATE FVS_StandInit_WorkTable SET Site_Species={1}, Site_Index={2} WHERE STAND_ID={0}; ",
		                    strStandID, strSiteSpecies, strSiteIndex);
		            }

		            public static string DeleteFvsStandInitWorkTable()
		            {
		                return "DROP TABLE FVS_StandInit_WorkTable;";
		            }


		            public static string[] DeleteDwmWorkTables()
		            {
		                string[] strSQL = new string[6];
		                int idx = 0;
		                strSQL[idx++] = "DROP TABLE CWDTotalLengthWorkTable;";
		                strSQL[idx++] = "DROP TABLE CwdPieceWorkTable;";
		                strSQL[idx++] = "DROP TABLE DWM_CWD_Aggregates_WorkTable;";
		                strSQL[idx++] = "DROP TABLE DWM_FWD_Aggregates_WorkTable;";
		                strSQL[idx++] = "DROP TABLE DWM_DuffLitter_Aggregates_WorkTable;";
		                strSQL[idx++] = "DROP TABLE Dwm_Fwd_RsnCtCd_WorkTable;";
		                return strSQL;
		            }

                    public static string PopulateStandInit(string strSourceStandTableAlias, string strCondTable, string strVariant)
                    {
                        string strSQL = "INSERT INTO " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                             " SELECT " + strSourceStandTableAlias + ".*" +
                             " FROM " + strSourceStandTableAlias +
                             " INNER JOIN " + strCondTable + " ON TRIM(" + strCondTable + ".cn) = " + strSourceStandTableAlias + ".STAND_CN" +
                             " WHERE " + strSourceStandTableAlias + ".VARIANT = '" + strVariant + "'" +
                             " AND " + strCondTable + ".landclcd = 1";
                        return strSQL;
                    }

                    public static string UpdateFromCond(string strCondTable, string strVariant)
                    {
                        string strSQL = "UPDATE " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                         " INNER JOIN " + strCondTable + " ON TRIM(" + strCondTable + ".cn) = " + Tables.FIA2FVS.DefaultFvsInputStandTableName + ".STAND_CN" +
                         " SET STAND_ID = trim(biosum_cond_id)," +
                         " SAM_WT = ACRES" +
                         " WHERE " + Tables.FIA2FVS.DefaultFvsInputStandTableName + ".VARIANT = '" + strVariant + "'";
                        return strSQL;
                    }

                    public static string UpdateForestType()
                    {
                        return "UPDATE " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                         " SET FOREST_TYPE = FOREST_TYPE_FIA"; ;
                    }

                    public static string SetFuelModelToNull()
                    {
                        return "UPDATE " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                         " SET FUEL_MODEL = null";
                    }

                    public static string CopyDwmColumns(string strVariant)
                    {
                        string strSQL = "UPDATE " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                         " INNER JOIN FVS_StandInit ON TRIM(FVS_StandInit.STAND_ID) = " + Tables.FIA2FVS.DefaultFvsInputStandTableName + ".STAND_ID" +
                         " SET " + Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_0_25_H = FVS_StandInit.FUEL_0_25_H, " +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_25_1_H = FVS_StandInit.FUEL_25_1_H, " +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_1_3_H = FVS_StandInit.FUEL_1_3_H, " +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_3_6_H = FVS_StandInit.FUEL_3_6_H, " +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_6_12_H = FVS_StandInit.FUEL_6_12_H, " +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_12_20_H = FVS_StandInit.FUEL_12_20_H, " +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_20_35_H = FVS_StandInit.FUEL_20_35_H," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_35_50_H = FVS_StandInit.FUEL_35_50_H," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_GT_50_H = FVS_StandInit.FUEL_GT_50_H," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_0_25_S = FVS_StandInit.FUEL_0_25_S," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_25_1_S = FVS_StandInit.FUEL_25_1_S," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_1_3_S = FVS_StandInit.FUEL_1_3_S," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_3_6_S = FVS_StandInit.FUEL_3_6_S," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_6_12_S = FVS_StandInit.FUEL_6_12_S," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_12_20_S = FVS_StandInit.FUEL_12_20_S," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_20_35_S = FVS_StandInit.FUEL_20_35_S," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_35_50_S = FVS_StandInit.FUEL_35_50_S," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_GT_50_S = FVS_StandInit.FUEL_GT_50_S," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_LITTER = FVS_StandInit.FUEL_LITTER," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".FUEL_DUFF = FVS_StandInit.FUEL_DUFF," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".SmallMediumTotalLength = FVS_StandInit.SmallMediumTotalLength," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".LargeTotalLength = FVS_StandInit.LargeTotalLength," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".CWDTotalLength = FVS_StandInit.CWDTotalLength," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".DuffPitCount = FVS_StandInit.DuffPitCount," +
                         Tables.FIA2FVS.DefaultFvsInputStandTableName + ".LitterPitCount = FVS_StandInit.LitterPitCount" +
                         " WHERE " + Tables.FIA2FVS.DefaultFvsInputStandTableName + ".VARIANT = '" + strVariant + "'";
                        return strSQL;
                    }

                    public static string SetDwmColumnsToNull(string strVariant)
                    {
                        string strSQL = "UPDATE " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                         " SET FUEL_0_25_H = NULL, " +
                         "FUEL_25_1_H = NULL, " +
                         "FUEL_1_3_H = NULL, " +
                         "FUEL_3_6_H = NULL, " +
                         "FUEL_6_12_H = NULL, " +
                         "FUEL_12_20_H = NULL, " +
                         "FUEL_20_35_H = NULL, " +
                         "FUEL_35_50_H = NULL, " +
                         "FUEL_GT_50_H = NULL, " +
                         "FUEL_0_25_S = NULL, " +
                         "FUEL_25_1_S = NULL, " +
                         "FUEL_1_3_S = NULL, " +
                         "FUEL_3_6_S = NULL, " +
                         "FUEL_6_12_S = NULL, " +
                         "FUEL_12_20_S = NULL, " +
                         "FUEL_20_35_S = NULL, " +
                         "FUEL_35_50_S = NULL, " +
                         "FUEL_GT_50_S = NULL, " +
                         "FUEL_LITTER = NULL, " +
                         "FUEL_DUFF = NULL " +
                         " WHERE " + Tables.FIA2FVS.DefaultFvsInputStandTableName + ".VARIANT = '" + strVariant + "'";
                        return strSQL;
                    }
                }

		        //All the queries necessary to create the FVSIn.accdb FVS_TreeInit table using intermediate tables
                public class TreeInit
                {
                    TreeInit()
                    {
                    }

                    public static string BulkImportTreeDataFromBioSumMaster(string strVariant, string strDestTable,
                        string strCondTableName, string strPlotTableName, string strTreeTableName)
                    {
                        string strInsertIntoTreeInit =
                            "INSERT INTO " + strDestTable +
                            " (Stand_ID, Tree_ID, Tree_Count, History, Species, " +
                            "DBH, DG, Htcd, Ht, HtTopK, CrRatio, " +
                            "Damage1, Severity1, Damage2, Severity2, Damage3, Severity3, " +
                            "Prescription, Slope, Aspect, PV_Code, TreeValue, Age, cullbf, mist_cl_cd, " +
                            "fvs_dmg_ag1, fvs_dmg_sv1, fvs_dmg_ag2, fvs_dmg_sv2, fvs_dmg_ag3, fvs_dmg_sv3, TreeCN)  ";
                        string strBioSumWorkTableSelectStmt =
                            "SELECT c.biosum_cond_id, VAL(t.fvs_tree_id) as Tree_ID, t.tpacurr, iif(iif(t.statuscd is null, 0, t.statuscd)=1, 1, 9) as History, t.spcd, " +
                            "t.dia, t.inc10yr, t.htcd, iif(t.ht is null,0,t.ht), iif(t.actualht is null,0,t.actualht), t.cr, " +
                            "0 as Damage1, 0 as Severity1, 0 as Damage2, 0 as Severity2, 0 as Damage3, 0 as Severity3, " +
                            "0 as Prescription, c.slope, c.aspect, c.habtypcd1, 3 as TreeValue, t.bhage as Age, t.cullbf, t.mist_cl_cd, " +
                            "fvs_dmg_ag1, fvs_dmg_sv1, fvs_dmg_ag2, fvs_dmg_sv2, fvs_dmg_ag3, fvs_dmg_sv3, t.cn ";
                        string strFromTableExpr = "FROM " +
                                                  strCondTableName + " c, " + strPlotTableName + " p, " +
                                                  strTreeTableName + " t ";
                        string strFilters =
                            "WHERE t.biosum_cond_id=c.biosum_cond_id AND p.biosum_plot_id=c.biosum_plot_id " +
                            "AND t.dia > 0 AND c.landclcd=1 " +
                            "AND ucase(trim(p.fvs_variant)) = \'" + strVariant.Trim().ToUpper() + "\'";
                        string strSQL = strInsertIntoTreeInit + strBioSumWorkTableSelectStmt + strFromTableExpr +
                                        strFilters;
                        return strSQL;
                    }

                    public static string CreateSpcdConversionTable(string strCondTableName, string strPlotTableName,
                        string strTreeTableName, string strTreeSpeciesTableName)
                    {
                        //Updating FIA Species Codes to FVS Species Codes
                        //Build the temporary species code conversion table
                        string strSelectIntoTempConversionTable =
                            "SELECT DISTINCT p.FVS_VARIANT AS PLOT_FVS_VARIANT, ts.FVS_VARIANT AS TREE_SPECIES_FVS_VARIANT, t.SPCD AS FIA_SPCD, ts.FVS_INPUT_SPCD INTO SPCD_CHANGE_WORK_TABLE ";
                        string strSpcdSources =
                            "FROM ((" + strCondTableName + " AS c INNER JOIN " + strPlotTableName +
                            " AS p ON c.biosum_plot_id = p.biosum_plot_id) INNER JOIN " + strTreeTableName +
                            " AS t ON c.BIOSUM_COND_ID = t.biosum_cond_id) LEFT JOIN " + strTreeSpeciesTableName +
                            " AS ts ON t.SPCD = ts.SPCD ";
                        string strSpcdConversionFilters =
                            "WHERE ts.FVS_VARIANT = p.FVS_VARIANT AND ts.FVS_INPUT_SPCD Is Not Null And ts.FVS_INPUT_SPCD <> t.SPCD;";
                        string strSQL = strSelectIntoTempConversionTable + strSpcdSources + strSpcdConversionFilters;
                        return strSQL;
                    }


                    public static string UpdateFVSSpeciesCodeColumn(string strVariant, string strFVSTreeInitWorkTable)
                    {
                        string strSQL = "UPDATE " + strFVSTreeInitWorkTable +
                                        " AS fvstree INNER JOIN SPCD_CHANGE_WORK_TABLE AS spcdchange ON VAL(fvstree.SPECIES) = spcdchange.FIA_SPCD " +
                                        "SET fvstree.SPECIES = CSTR(spcdchange.FVS_INPUT_SPCD) " +
                                        "WHERE TRIM(spcdchange.PLOT_FVS_VARIANT)=\'" + strVariant.Trim().ToUpper() +
                                        "\'; ";
                        return strSQL;
                    }

                    public static string DeleteCrRatiosForDeadTrees(string strDestTable)
                    {
                        return "UPDATE " + strDestTable + " SET CrRatio=null WHERE History=9;";
                    }

                    public static string RoundCrRatioToSingleDigitCodes(string strDestTable)
                    {
                        /*This is a test to compare predispose and fvsin CrRatio when you round up*/
                        string strSQL = "UPDATE " + strDestTable + " SET " +
                                        "CrRatio=iif(crratio is null, null, iif(len(trim(cstr(crratio)))=0, 0, " +
                                        "iif(0 <= crratio AND crratio <= 10, 1, " +
                                        "iif(crratio <= 20, 2, " +
                                        "iif(crratio <= 30, 3, " +
                                        "iif(crratio <= 40, 4, " +
                                        "iif(crratio <= 50, 5, " +
                                        "iif(crratio <= 60, 6, " +
                                        "iif(crratio <= 70, 7, " +
                                        "iif(crratio <= 80, 8, " +
                                        "iif(crratio <= 100, 9, null)))))))))));";
                        return strSQL;
                    }

                    public static string DeleteHtAndHtTopKForUnknownHtcd(string strDestTable)
                    {
                        return "UPDATE " + strDestTable + " SET Ht=0, HtTopK=0 WHERE Htcd NOT IN (1,2,3,4);";
                    }

                    public static string SetBrokenTopFlag(string strDestTable)
                    {
                        return "UPDATE " + strDestTable +
                               " SET hasBrokenTop = -1 " +
                               "WHERE HtTopK < Ht AND 0 < HtTopK;";
                    }


                    public static string SetHtTopKToZeroIfGteHt(string strDestTable)
                    {
                        return "UPDATE " + strDestTable + " SET HtTopK=0 WHERE Ht <= HtTopK";
                    }


                    public static string SetInferredSeedlingDbh(string strDestTable)
                    {
                        return "UPDATE " + strDestTable +
                               " SET Dbh=0.1 WHERE Tree_Count > 25 AND Dbh <= 0 AND History=1;";
                    }

                    public static string[] DamageCodes(string strDestTable)
                    {
                        string[] strDamageCodeUpdates = new string[15];

                        //use precalculated damage codes if possible
                        strDamageCodeUpdates[0] =
                            "UPDATE " + strDestTable +
                            " SET Damage1=fvs_dmg_ag1, Damage2=fvs_dmg_ag2, Damage3=fvs_dmg_ag3, Severity1=fvs_dmg_sv1, Severity2=fvs_dmg_sv2, Severity3=fvs_dmg_sv3 WHERE History=1 AND fvs_dmg_ag1 is not null;";

                        //Cull board feet
                        strDamageCodeUpdates[1] =
                            "UPDATE " + strDestTable +
                            " SET Damage1=25, Severity1=IIF(cullbf>=100, 99, cullbf) WHERE History=1 AND cullbf>0;";

                        //FVS Mistletoe damage codes
                        string strDamage1_Filter =
                            "History=1 AND iif(mist_cl_cd is null,0,mist_cl_cd) <> 0 AND Damage1 = 0 AND fvs_dmg_ag1 is null;";
                        string strDamage2_Filter =
                            "History=1 AND iif(mist_cl_cd is null,0,mist_cl_cd) <> 0 AND Damage1 NOT IN (0, 30, 31, 32, 33, 34) AND fvs_dmg_ag1 is null;";
                        strDamageCodeUpdates[2] =
                            "UPDATE " + strDestTable +
                            " SET Damage1=31, Severity1=mist_cl_cd WHERE Species=\'108\' AND " +
                            strDamage1_Filter;

                        strDamageCodeUpdates[3] =
                            "UPDATE " + strDestTable +
                            " SET Damage2=31, Severity2=mist_cl_cd WHERE Species=\'108\' AND " +
                            strDamage2_Filter;

                        strDamageCodeUpdates[4] =
                            "UPDATE " + strDestTable +
                            " SET Damage1=32, Severity1=mist_cl_cd WHERE Species=\'073\' AND " +
                            strDamage1_Filter;

                        strDamageCodeUpdates[5] =
                            "UPDATE " + strDestTable +
                            " SET Damage2=32, Severity2=mist_cl_cd WHERE Species=\'073\' AND " +
                            strDamage2_Filter;

                        strDamageCodeUpdates[6] =
                            "UPDATE " + strDestTable +
                            " SET Damage1=33, Severity1=mist_cl_cd WHERE Species=\'202\' AND " +
                            strDamage1_Filter;

                        strDamageCodeUpdates[7] =
                            "UPDATE " + strDestTable +
                            " SET Damage2=33, Severity2=mist_cl_cd WHERE Species=\'202\' AND " +
                            strDamage2_Filter;

                        strDamageCodeUpdates[8] =
                            "UPDATE " + strDestTable +
                            " SET Damage1=34, Severity1=mist_cl_cd WHERE Species=\'122\' AND " +
                            strDamage1_Filter;

                        strDamageCodeUpdates[9] =
                            "UPDATE " + strDestTable +
                            " SET Damage2=34, Severity2=mist_cl_cd WHERE Species=\'122\' AND " +
                            strDamage2_Filter;

                        //default mist_cl_cd damage code if fvs species don't match previous four cases
                        strDamageCodeUpdates[10] =
                            "UPDATE " + strDestTable +
                            " SET Damage1=30, Severity1=mist_cl_cd WHERE Species NOT IN (\'202\',\'108\',\'122\',\'073\') AND " +
                            strDamage1_Filter;

                        strDamageCodeUpdates[11] =
                            "UPDATE " + strDestTable +
                            " SET Damage2=30, Severity2=mist_cl_cd WHERE Species NOT IN (\'202\',\'108\',\'122\',\'073\') AND " +
                            strDamage2_Filter;

                        //broken top 96 added to least priority(?) damage column so fill it last
                        strDamageCodeUpdates[12] =
                            "UPDATE " + strDestTable +
                            " SET Damage1=96 WHERE History=1 AND hasBrokenTop AND Damage1 = 0 AND fvs_dmg_ag1 is null;";

                        //0 means nothing was assigned. 96 means damage2 doesn't need to repeat damage1
                        strDamageCodeUpdates[13] =
                            "UPDATE " + strDestTable +
                            " SET Damage2=96 WHERE History=1 AND hasBrokenTop AND Damage1 NOT IN (0,96) AND Damage2 = 0 AND fvs_dmg_ag1 IS Null;";

                        strDamageCodeUpdates[14] =
                            "UPDATE " + strDestTable +
                            " SET Damage3=96 WHERE History=1 AND hasBrokenTop AND Damage1 NOT IN (0,96) AND Damage2 not in (0,96) and Damage3=0 AND fvs_dmg_ag1 IS Null;";
                        /*END DAMAGE CODES*/
                        return strDamageCodeUpdates;
                    }

                    public static string[] TreeValueClass(string strDestTable)
                    {
                        /*Value Classes: 
                          * All trees (live and dead) initialized to 3. 
                          * If Damage1=25, TreeValue=3 again (redundant). 
                          * Else if Severity > 0, TreeValue=2. 
                          * Else TreeValue=1*/
                        string[] strTreeValueUpdates = new string[2];
                        strTreeValueUpdates[0] = "UPDATE " + strDestTable +
                                                 " SET TreeValue=2 WHERE History=1 AND Damage1<>25 AND Severity1>0;";
                        strTreeValueUpdates[1] = "UPDATE " + strDestTable +
                                                 " SET TreeValue=1 WHERE History=1 AND Damage1<>25 AND Severity1<=0;";
                        return strTreeValueUpdates;
                    }

                    public static string PadSpeciesWithZero(string strDestTable)
                    {
                        //This addresses a problem with FVSOut having incorrect Species Codes being translated into "2TD"
                        //Update the Species column to Trim and "PadLeft" with 0s by Creating a "000" and concatenating with Trim(Species) and taking the rightmost 3 digits
                        //Example: Right("000" & "17      ", 3) => "017". Species=="17      "  because the Species column is width 8 in FVSIn.accdb
                        return "UPDATE " + strDestTable + " SET Species = Right(String(3, \'0\') & Trim(Species), 3);";
                    }

                    public static string TranslateWorkTableToTreeInitTable(string strSourceTable, string strDestTable)
                    {
                        string strInsertIntoTreeInit =
                            "INSERT INTO " + strDestTable;
                        string strBioSumWorkTableSelectStmt =
                            " SELECT Stand_ID, StandPlot_ID, Tree_ID, Tree_Count, History, Species, " +
                            "DBH, DG, Ht, HTG, HtTopK, CrRatio,  " +
                            "Damage1, Severity1, Damage2, Severity2, Damage3, Severity3, " +
                            "TreeValue, Prescription, Age, Slope, Aspect, PV_Code, TopoCode, SitePrep ";
                        string strFromTableExpr = "FROM " + strSourceTable + ";";
                        return strInsertIntoTreeInit + strBioSumWorkTableSelectStmt + strFromTableExpr;
                    }

                    public static string RoundSingleDigitPercentageCrRatiosUpTo10(string strDestTable)
                    {
                        return "UPDATE " + strDestTable + " SET CrRatio=10 WHERE CrRatio<10;";
                    }

                    public static string RoundSingleDigitPercentageCrRatiosDownTo1(string strDestTable)
                    {
                        return "UPDATE " + strDestTable + " SET CrRatio=1 WHERE CrRatio<10;";
                    }

                    public static string DeleteWorkTable()
                    {
                        return "DROP TABLE FVS_TreeInit_WorkTable;";
                    }

                    public static string DeleteSpcdChangeWorkTable()
                    {
                        return "DROP TABLE SPCD_CHANGE_WORK_TABLE;";
                    }

                    public static string PopulateTreeInit(string strSourceTreeTableAlias, string strSourceStandTableAlias,
                        string strCondTable, string strTreeTable, string strVariant)
                    {
                        string strSql = "INSERT INTO " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                         " SELECT " + strSourceTreeTableAlias + ".*" +
                         " FROM " + strSourceTreeTableAlias + ", " + strTreeTable + ", " +
                           strSourceStandTableAlias + ", " + strCondTable +
                         " WHERE " + strCondTable + ".biosum_cond_id = " + strTreeTable + ".biosum_cond_id" +
                         " AND TRIM(" + strTreeTable + ".cn) = " + strSourceTreeTableAlias + ".TREE_CN" +
                         " AND TRIM(" + strCondTable + ".cn) = " + strSourceStandTableAlias + ".STAND_CN" +
                         " AND TRIM(" + strCondTable + ".cn) = " + strSourceTreeTableAlias + ".STAND_CN" +
                         " AND " + strCondTable + ".landclcd = 1 AND " + strTreeTable + ".DIA > 0 AND " +
                         strSourceStandTableAlias + ".VARIANT = '" + strVariant + "'";
                        return strSql;
                    }

                    public static string UpdateFromCond(string strCondTable, string strVariant, string strTargetTable)
                    {
                        string strSQL = "UPDATE(" + strTargetTable +
                            " INNER JOIN cond ON " + strTargetTable + ".STAND_CN = TRIM(" + strCondTable + ".cn))" +
                            " INNER JOIN " + Tables.FIA2FVS.DefaultFvsInputStandTableName + " ON " + strTargetTable + ".STAND_CN = " + Tables.FIA2FVS.DefaultFvsInputStandTableName + ".STAND_CN" +
                            " SET " + strTargetTable + ".STAND_ID = Trim(" + strCondTable + ".BIOSUM_COND_ID)" +
                            " WHERE " + Tables.FIA2FVS.DefaultFvsInputStandTableName + ".VARIANT = '" + strVariant + "'";
                        return strSQL;
                    }

                    public static string UpdateFromTree(string strTreeTable, string strTargetTable)
                    {
                        string strSQL = "UPDATE " + strTargetTable +
                         " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                         " AND " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                         " SET TREE_ID = trim(fvs_tree_id)";
                        return strSQL;
                    }

                    public static string UpdateTreeCount(string strTreeTable, string strTargetTable)
                    {
                        string strSQL = "UPDATE " + strTargetTable +
                         " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                         " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                         " SET TREE_COUNT = tpacurr where statuscd not in (2,3)";
                        return strSQL;
                    }

                    public static string DeleteSeedlings(string strTargetTable)
                    {
                        string strSQL = "DELETE FROM " + strTargetTable +
                         " WHERE DIAMETER = 0.1 AND LEFT(TREE_CN,1) = 'S'";
                        return strSQL;
                    }

                    public static string CopyGrmColumns(string strTargetTable)
                    {
                        string strSQL = "UPDATE " + strTargetTable +
                         " INNER JOIN FVS_TreeInit ON FVS_TreeInit.Tree_ID = " + strTargetTable + ".TREE_ID" +
                         " AND " + strTargetTable + ".STAND_ID = FVS_TreeInit.Stand_ID" +
                         " SET " + strTargetTable + ".DG = FVS_TreeInit.DG, " +
                         strTargetTable + ".HTG = FVS_TreeInit.HTG";
                         return strSQL;
                    }

                    public static string SetGrmColumnsToNull(string strTargetTable)
                    {
                        string strSQL = "UPDATE " + strTargetTable +
                         " SET " + strTargetTable + ".DG = NULL, " +
                         strTargetTable + ".HTG = NULL";
                        return strSQL;
                    }

                    public static string[] UpdateDamageCodesForCull(string strTreeTable, string strTargetTable)
                    {
                        string[] arrDamageCodeUpdates = new string[18];
                        // First pass updates the damage code and severity fields with cull (25) if they are null or = 96, 97
                        arrDamageCodeUpdates[0] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage1 = 25, Severity1 = (" + strTreeTable + ".CULL + " + strTreeTable + ".ROUGHCULL)" +
                            " WHERE (Damage1 IS Null or Damage1 in (96,97)) AND History = 1" +
                            " AND(" + strTreeTable + ".cull + " + strTreeTable + ".roughcull) > 0";
                        arrDamageCodeUpdates[1] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage2 = 25, Severity2 = (" + strTreeTable + ".CULL + " + strTreeTable + ".ROUGHCULL)" +
                            " WHERE Damage1 <> 25 AND (Damage2 IS Null or Damage2 in (96,97)) AND History = 1" +
                            " AND(" + strTreeTable + ".cull + " + strTreeTable + ".roughcull) > 0";
                        arrDamageCodeUpdates[2] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage3 = 25, Severity3 = (" + strTreeTable + ".CULL + " + strTreeTable + ".ROUGHCULL)" +
                            " WHERE Damage1 <> 25 AND Damage2 <> 25 AND (Damage3 IS Null or Damage3 in (96,97)) AND History = 1" +
                            " AND(" + strTreeTable + ".cull + " + strTreeTable + ".roughcull) > 0";
                        // second pass updates the damage code and severity fields with mistletoe (30) if they are null or = 96, 97
                        arrDamageCodeUpdates[3] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage1 = 30, Severity1 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE (Damage1 IS Null OR Damage1 in (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd NOT IN (202, 108, 122, 73)";
                        arrDamageCodeUpdates[4] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage2 = 30, Severity2 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE Damage1 <> 30 AND (Damage2 IS Null OR Damage2 IN (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd NOT IN (202, 108, 122, 73)";
                        arrDamageCodeUpdates[5] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage3 = 30, Severity3 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE Damage1 <> 30 AND Damage2 <> 30 AND (Damage3 IS Null OR Damage3 IN (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd NOT IN (202, 108, 122, 73)";
                        // third pass updates the damage code and severity fields with mistletoe (31) if they are null or = 96, 97
                        arrDamageCodeUpdates[6] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage1 = 31, Severity1 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE (Damage1 IS Null OR Damage1 in (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd = 108";
                        arrDamageCodeUpdates[7] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage2 = 31, Severity2 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE Damage1 <> 31 AND (Damage2 IS Null OR Damage2 IN (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd = 108";
                        arrDamageCodeUpdates[8] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage3 = 31, Severity3 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE Damage1 <> 31 AND Damage2 <> 31 AND (Damage3 IS Null OR Damage3 IN (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd = 108";
                        // fourth pass updates the damage code and severity fields with mistletoe (32) if they are null or = 96, 97
                        arrDamageCodeUpdates[9] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage1 = 32, Severity1 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE (Damage1 IS Null OR Damage1 in (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd = 73";
                        arrDamageCodeUpdates[10] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage2 = 32, Severity2 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE Damage1 <> 32 AND (Damage2 IS Null OR Damage2 IN (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd = 73";
                        arrDamageCodeUpdates[11] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage3 = 32, Severity3 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE Damage1 <> 32 AND Damage2 <> 32 AND (Damage3 IS Null OR Damage3 IN (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd = 73";
                        // fifth pass updates the damage code and severity fields with mistletoe (33) if they are null or = 96, 97
                        arrDamageCodeUpdates[12] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage1 = 33, Severity1 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE (Damage1 IS Null OR Damage1 in (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd = 202";
                        arrDamageCodeUpdates[13] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage2 = 33, Severity2 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE Damage1 <> 33 AND (Damage2 IS Null OR Damage2 IN (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd = 202";
                        arrDamageCodeUpdates[14] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage3 = 33, Severity3 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE Damage1 <> 33 AND Damage2 <> 33 AND (Damage3 IS Null OR Damage3 IN (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd = 202";
                        // Final 6th pass updates the damage code and severity fields with mistletoe (34) if they are null or = 96, 97
                        arrDamageCodeUpdates[15] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage1 = 34, Severity1 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE (Damage1 IS Null OR Damage1 in (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd = 122";
                        arrDamageCodeUpdates[16] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage2 = 34, Severity2 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE Damage1 <> 34 AND (Damage2 IS Null OR Damage2 IN (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd = 122";
                        arrDamageCodeUpdates[17] = "UPDATE " + strTargetTable +
                            " INNER JOIN " + strTreeTable + " ON " + strTargetTable + ".STAND_ID = TRIM(" + strTreeTable + ".biosum_cond_id)" +
                            " AND " + strTargetTable + ".TREE_CN = TRIM(" + strTreeTable + ".cn)" +
                            " SET Damage3 = 34, Severity3 = " + strTreeTable + ".mist_cl_cd" +
                            " WHERE Damage1 <> 34 AND Damage2 <> 34 AND (Damage3 IS Null OR Damage3 IN (96,97)) AND History = 1" +
                            " AND " + strTreeTable + ".mist_cl_cd > 0 AND " + strTreeTable + ".spcd = 122";

                        return arrDamageCodeUpdates;
                    }

                    public static string[] UpdateTreeValue(string strTargetTable)
                    {
                        string[] arrTreeValueUpdates = new string[3];
                        // First pass sets treevalue = 1 for all trees
                        arrTreeValueUpdates[0] = "UPDATE " + strTargetTable +
                            " SET TREEVALUE = 1";
                        // Second pass sets treevalue = 3 for cull trees
                        arrTreeValueUpdates[1] = "UPDATE " + strTargetTable +
                            " SET TREEVALUE = 3 WHERE (Damage1 = 25 AND SEVERITY1 >= 25) OR" +
                            " (Damage2 = 25 AND SEVERITY2 >= 25) OR (Damage3 = 25 AND SEVERITY3 >= 25)";
                        // Third pass sets treevalue = 2 for all other damage codes
                        arrTreeValueUpdates[2] = "UPDATE " + strTargetTable +
                            " SET TREEVALUE = 2 WHERE (Damage1 IS NOT NULL OR Damage2 IS NOT NULL OR Damage3 IS NOT NULL) AND" +
                            " TREEVALUE <> 3";
                        return arrTreeValueUpdates;
                    }

                        public static string UpdateFieldsFromTempTable(string strSourceTable, IList<string> lstFields)
                    {
                        string strSQL = "UPDATE " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                         " INNER JOIN " + strSourceTable + " ON " + strSourceTable + ".TREE_CN = " + Tables.FIA2FVS.DefaultFvsInputTreeTableName + ".TREE_CN" +
                         " AND " + strSourceTable + ".STAND_CN = " + Tables.FIA2FVS.DefaultFvsInputTreeTableName + ".STAND_CN" +
                         " SET ";
                        System.Text.StringBuilder sb = new System.Text.StringBuilder();
                        foreach (var item in lstFields)
                        {
                            sb.Append(Tables.FIA2FVS.DefaultFvsInputTreeTableName + "." + item + "=" + strSourceTable + "." + item);
                            sb.Append(",");
                        }
                        sb.Length--;    // Remove trailing comma
                        strSQL = strSQL + sb.ToString();
                        return strSQL;
                    }

                }

                    //This updates the FVSIn.GroupAddFilesAndKeywords table so that they FVS keywords have the correct DSNIn value
                    public class GroupAddFilesAndKeywords
                {
                    public static string UpdateAllPlots(string strFVSInFileName)
                    {
                        return
                            "UPDATE FVS_GroupAddFilesAndKeywords SET FVS_GroupAddFilesAndKeywords.FVSKeywords = " +
                            "\"Database\" + Chr(13) + Chr(10) + \"DSNIn\" + Chr(13) + Chr(10) + \"" + strFVSInFileName +
                            "\" + Chr(13) + Chr(10) + \"StandSQL\" + Chr(13) + Chr(10) + \"SELECT *\" + Chr(13) + Chr(10) + " +
                            "\"FROM FVS_PlotInit\" + Chr(13) + Chr(10) + \"WHERE StandPlot_ID = '%StandID%'\" + Chr(13) + Chr(10) + " +
                            "\"EndSQL\" + Chr(13) + Chr(10) + \"TreeSQL\" + Chr(13) + Chr(10) + \"SELECT *\" + Chr(13) + Chr(10) + " +
                            "\"FROM FVS_TreeInit\" + Chr(13) + Chr(10) + \"WHERE StandPlot_ID = '%StandID%'\" + Chr(13) + Chr(10) + " +
                            "\"EndSQL\" + Chr(13) + Chr(10) + \"END\" WHERE (FVS_GroupAddFilesAndKeywords.Groups=\"All_Plots\");";
                    }

                    public static string UpdateAllStands(string strFVSInFileName)
                    {
                        return
                            "UPDATE FVS_GroupAddFilesAndKeywords SET FVS_GroupAddFilesAndKeywords.FVSKeywords = " +
                            "\"Database\" + Chr(13) + Chr(10) + \"DSNIn\" + Chr(13) + Chr(10) + \"" + strFVSInFileName +
                            "\" + Chr(13) + Chr(10) + \"StandSQL\" + Chr(13) + Chr(10) + \"SELECT *\" + Chr(13) + Chr(10) + " +
                            "\"FROM FVS_StandInit\" + Chr(13) + Chr(10) + \"WHERE Stand_ID = '%StandID%'\" + Chr(13) + Chr(10) + " +
                            "\"EndSQL\" + Chr(13) + Chr(10) + \"TreeSQL\" + Chr(13) + Chr(10) + \"SELECT *\" + Chr(13) + Chr(10) + " +
                            "\"FROM FVS_TreeInit\" + Chr(13) + Chr(10) + \"WHERE Stand_ID = '%StandID%'\" + Chr(13) + Chr(10) + " +
                            "\"EndSQL\" + Chr(13) + Chr(10) + \"END\" WHERE (FVS_GroupAddFilesAndKeywords.Groups=\"All_Stands\");";
                    }

                }
            }

        }

        public class VolumeAndBiomass
        {
            public VolumeAndBiomass()
            {
            }

            public class FIAPlotInput
            {
                public FIAPlotInput()
                {
                }

                public static string BuildInputTableForVolumeCalculation_Step1(
                    string p_strInputVolumesTable, string p_strFIATreeTable, string p_strColumns, string p_strValues)
                {
                    return
                        $"INSERT INTO {p_strInputVolumesTable} ({p_strColumns}) SELECT {p_strValues} FROM {p_strFIATreeTable} WHERE biosum_status_cd=9 and dia >= 1.0";
                }

                public static string BuildInputTableForVolumeCalculation_Step2(
                    string p_strInputVolumesTable,
                    string p_strFIATreeTable,
                    string p_strFIAPlotTable,
                    string p_strFIACondTable)
                {
                    return "UPDATE " + p_strInputVolumesTable + " i " +
                           "INNER JOIN ((" + p_strFIATreeTable + " t " +
                           "INNER JOIN " + p_strFIACondTable + " c " +
                           "ON t.biosum_cond_id=c.biosum_cond_id) " +
                           "INNER JOIN " + p_strFIAPlotTable + " p " +
                           "ON p.biosum_plot_id=c.biosum_plot_id) " +
                           "ON i.tre_cn=t.cn " +
                           "SET i.spcd=t.spcd," +
                           "i.statuscd=IIF(t.statuscd IS NULL,1,t.statuscd)," +
                           "i.treeclcd=t.treeclcd," +
                           "i.cull=IIF(t.cull IS NULL,0,t.cull)," +
                           "i.roughcull=IIF(t.roughcull IS NULL,0,t.roughcull)," +
                           "i.decaycd=IIF(t.decaycd IS NULL,0,t.decaycd)," +
                           "i.balive=c.balive," +
                           "i.precipitation=p.precipitation";
                }

                public static string BuildInputTableForVolumeCalculation_Step3(
                    string p_strInputVolumesTable,
                    string p_strFIACondTable)
                {
                    return "UPDATE " + p_strInputVolumesTable + " i " +
                           "INNER JOIN " + p_strFIACondTable + " c " +
                           "ON i.cnd_cn=c.biosum_cond_id " +
                           "SET i.vol_loc_grp=IIF(INSTR(1,c.vol_loc_grp,'22') > 0,'S26LEOR',c.vol_loc_grp)," +
                           "i.statuscd=IIF(i.statuscd IS NULL,1,i.statuscd)," +
                           "i.cull=IIF(i.cull IS NULL,0,i.cull)," +
                           "i.roughcull=IIF(i.roughcull IS NULL,0,i.roughcull)," +
                           "i.decaycd=IIF(i.decaycd IS NULL,0,i.decaycd)";
                }

                public static string BuildInputTableForVolumeCalculation_Step4(
                    string p_strCullTable,
                    string p_strInputVolumesTable)
                {
                    return "SELECT tre_cn, IIF(cull IS NOT NULL AND roughcull IS NOT NULL, cull + roughcull," +
                           "IIF(cull IS NOT NULL,cull," +
                           "IIF(roughcull IS NOT NULL, roughcull,0))) AS totalcull " +
                           "INTO " + p_strCullTable + " " +
                           "FROM " + p_strInputVolumesTable;
                }

                public class PNWRS
                {
                    public PNWRS()
                    {
                    }


                    /// <summary>
                    /// Update the TREECLCD column for all trees 
                    /// by comparing SPCD,ROUGHCULL,STATUSCD and TOTALCULL columns
                    /// </summary>
                    /// <param name="p_strCullTable"></param>
                    /// <param name="p_strInputVolumesTable"></param>
                    /// <returns></returns>
                    public static string BuildInputTableForVolumeCalculation_Step5(
                        string p_strCullTable,
                        string p_strInputVolumesTable)
                    {
                        return "UPDATE " + p_strInputVolumesTable + " a " +
                               "INNER JOIN " + p_strCullTable + " b " +
                               "ON a.tre_cn=b.tre_cn " +
                               "SET a.treeclcd=" +
                               "IIF(a.SpCd IN (62,65,66,106,133,138,304,321,322,475,756,758,990),3," +
                               "IIF(a.StatusCd=2,3," +
                               "IIF(b.totalcull < 75,2," +
                               "IIF(a.roughcull > 37.5,3,4))))";
                    }

                    /// <summary>
                    /// Update the TREECLCD column for TREECLCD=3 and dead trees 
                    /// by comparing SPCD,DBH,STATUSCD and DECAYCD columns
                    /// </summary>
                    /// <param name="p_strCullTable"></param>
                    /// <param name="p_strInputVolumesTable"></param>
                    /// <returns></returns>
                    public static string BuildInputTableForVolumeCalculation_Step6(
                        string p_strCullTable,
                        string p_strInputVolumesTable)
                    {
                        return "UPDATE " + p_strInputVolumesTable + " a " +
                               "INNER JOIN " + p_strCullTable + " b " +
                               "ON a.tre_cn=b.tre_cn " +
                               "SET a.treeclcd=" +
                               "IIF(a.DecayCd > 1,4,IIF(a.DIA < 9 AND a.SpCd < 300,4,a.treeclcd)) " +
                               "WHERE a.treeclcd=3 AND a.statuscd=2 AND a.SpCd NOT IN (62,65,66,106,133,138,304,321,322,475,756,758,990)";
                    }
                }

                public static string WriteCalculatedVolumeAndBiomassColumnsToTreeTable(
                    string p_strBiosumCalcOutputTable)
                {
                    return $@"UPDATE {frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName} t 
                                INNER JOIN {Tables.VolumeAndBiomass.BiosumCalcOutputTable} o ON t.CN = o.TRE_CN
                                SET t.volcfgrs=IIF(o.VOLCFGRS_CALC IS NOT NULL, o.VOLCFGRS_CALC, null),
                                t.volcfnet=IIF(o.VOLCFNET_CALC IS NOT NULL, o.VOLCFNET_CALC, null),
                                t.volcfsnd=IIF(o.VOLCFSND_CALC IS NOT NULL, o.VOLCFSND_CALC, null),
                                t.volcsgrs=IIF(o.VOLCSGRS_CALC IS NOT NULL, o.VOLCSGRS_CALC, null),
                                t.voltsgrs=IIF(o.VOLTSGRS_CALC IS NOT NULL,o.VOLTSGRS_CALC,null),
                                t.drybiom=IIF(t.DRYBIOM IS NULL,o.DRYBIOM_CALC,t.DRYBIOM),
                                t.drybiot=IIF(t.DRYBIOT IS NULL,o.DRYBIOT_CALC,t.DRYBIOT),
                                t.drybio_bole=IIF(o.DRYBIO_BOLE_CALC IS NOT NULL, o.DRYBIO_BOLE_CALC, null),
                                t.drybio_top=IIF(o.DRYBIO_TOP_CALC IS NOT NULL, o.DRYBIO_TOP_CALC, null),
                                t.drybio_sapling=IIF(o.DRYBIO_SAPLING_CALC IS NOT NULL, o.DRYBIO_SAPLING_CALC, null),
                                t.drybio_wdld_spp=IIF(o.DRYBIO_WDLD_SPP_CALC IS NOT NULL, o.DRYBIO_WDLD_SPP_CALC, null)";
                }
            }

            public class FVSOut
            {
                public FVSOut()
                {
                }

                /// <summary>
                /// Insert all FVS_TREEs 
                /// </summary>
                /// <param name="p_strInputVolumesTable"></param>
                /// <param name="p_strFvsTreeTable"></param>
                /// <param name="p_strRxPackage"></param>
                /// <returns></returns>
                public static string BuildInputTableForVolumeCalculation_Step1(string p_strInputVolumesTable, string p_strFvsTreeTable, string p_strRxPackage)
                {
                    string strColumns = @"id,biosum_cond_id,invyr,fvs_variant,spcd,dbh,ht,actualht,cr,fvs_tree_id, fvscreatedtree_yn";
                    string strValues = @"id,biosum_cond_id,CINT(rxyear) AS invyr, fvs_variant, IIF(FvsCreatedTree_YN='Y',CINT(fvs_species),-1) AS spcd, dbh,estht,ht,pctcr,fvs_tree_id, fvscreatedtree_yn";
                    return $@"INSERT INTO {p_strInputVolumesTable} ({strColumns}) 
                           SELECT {strValues} 
                           FROM {p_strFvsTreeTable}
                           WHERE rxpackage='{p_strRxPackage.Trim()}' AND DBH >= 1.0";
                }

                /// <summary>
                /// Insert all FVS_TREEs 
                /// </summary>
                /// <param name="p_strInputVolumesTable"></param>
                /// <param name="p_strFvsTreeTable"></param>
                /// <param name="p_strRxPackage"></param>
                /// <returns></returns>
                public static string BuildInputSQLiteTableForVolumeCalculation_Step1(string p_strInputVolumesTable, string p_strFvsTreeTable, 
                    string p_strRxPackage, string p_strFvsVariant)
                {
                    string strColumns = @"id,biosum_cond_id,invyr,fvs_variant,spcd,dbh,ht,actualht,cr,fvs_tree_id, fvscreatedtree_yn";
                    string strValues = @"id,biosum_cond_id, cast(rxyear as integer) as invyr, fvs_variant, " +
                        "CASE WHEN FvsCreatedTree_YN='Y' THEN cast(fvs_species as integer) ELSE -1 END AS spcd, " +
                        "dbh,ht,ht,pctcr,fvs_tree_id, fvscreatedtree_yn";
                    return $@"INSERT INTO {p_strInputVolumesTable} ({strColumns})  
                           SELECT {strValues} 
                           FROM {p_strFvsTreeTable}
                           WHERE rxpackage='{p_strRxPackage.Trim()}' AND fvs_variant='{p_strFvsVariant.Trim()}' and dbh >= 1.0";
                }

                /// <summary>
                /// Insert FVS_TREEs that are not cycle 1 trees
                /// </summary>
                /// <param name="p_strInputVolumesTable"></param>
                /// <param name="p_strFvsTreeTable"></param>
                /// <returns></returns>
                public static string BuildInputTableForVolumeCalculation_Step1(string p_strInputVolumesTable, string p_strFvsTreeTable)
                {
                    string strColumns = "id,biosum_cond_id,invyr,fvs_variant,spcd,dbh,ht,actualht,cr,fvs_tree_id";
                    string values = "id,biosum_cond_id,CINT(rxyear) AS invyr, fvs_variant, IIF(FvsCreatedTree_YN='Y',CINT(fvs_species),-1) AS spcd, dbh,ht,ht,pctcr,fvs_tree_id ";
                    return $@"INSERT INTO {p_strInputVolumesTable} ({strColumns}) SELECT {values} FROM {p_strFvsTreeTable}";
                }

                /// <summary>
                /// Update tree fields with values from the MASTER.TREE records
                /// </summary>
                /// <param name="p_strInputVolumesTable"></param>
                /// <param name="p_strFIATreeTable"></param>
                /// <param name="p_strFIAPlotTable"></param>
                /// <param name="p_strFIACondTable"></param>
                /// <returns></returns>
                public static string BuildInputTableForVolumeCalculation_Step2(string p_strInputVolumesTable, string p_strFIATreeTable, string p_strFIAPlotTable, string p_strFIACondTable)
                {
                    return $@"UPDATE {p_strInputVolumesTable} i 
                                INNER JOIN (({p_strFIATreeTable} t 
                                    INNER JOIN {p_strFIACondTable} c ON t.biosum_cond_id=c.biosum_cond_id)
                                        INNER JOIN {p_strFIAPlotTable} p ON p.biosum_plot_id=c.biosum_plot_id) 
                                ON i.fvs_tree_id = t.fvs_tree_id and i.biosum_cond_id = t.biosum_cond_id 
                            SET i.spcd=t.spcd,
                                i.statuscd=IIF(t.statuscd IS NULL,1,t.statuscd),
                                i.treeclcd=t.treeclcd,
                                i.cull=IIF(t.cull IS NULL,0,t.cull),
                                i.roughcull=IIF(t.roughcull IS NULL,0,t.roughcull),
                                i.decaycd=IIF(t.decaycd IS NULL,0,t.decaycd),
                                i.balive=c.balive,
                                i.precipitation=p.precipitation,
                                i.ecosubcd = p.ecosubcd,
                                i.stdorgcd = c.stdorgcd,
                                i.actualht = IIF(t.actualht <> t.ht, i.ht - t.ht + t.actualht, i.actualht)";
                }

                /// <summary>
                /// Update biosum_volumes_input fields (balive, precipitation)
                /// </summary>
                /// <param name="p_strInputVolumesTable"></param>
                /// <param name="p_strFIATreeTable"></param>
                /// <param name="p_strFIAPlotTable"></param>
                /// <param name="p_strFIACondTable"></param>
                /// <returns></returns>
                public static string BuildInputTableForVolumeCalculation_Step2a(string p_strInputVolumesTable, string p_strFIAPlotTable, string p_strFIACondTable, string p_strFvsOutSummaryTable)
                {
                    return $@"UPDATE (({p_strFIAPlotTable} p INNER JOIN {p_strFIACondTable} c ON p.biosum_plot_id = c.biosum_plot_id) 
                                INNER JOIN {p_strInputVolumesTable} i ON c.biosum_cond_id = i.biosum_cond_id) 
                                INNER JOIN {p_strFvsOutSummaryTable} s ON i.invyr = s.Year AND i.biosum_cond_id = s.StandID
                                SET i.balive=s.ba, i.precipitation=p.precipitation
                                WHERE i.fvscreatedtree_yn='Y'";
                }

                /// <summary>
                /// Update biosum_volumes_input fields (precipitation)
                /// </summary>
                /// <param name="p_strInputVolumesTable"></param>
                /// <param name="p_strFIAPlotTable"></param>
                /// <param name="p_strFIACondTable"></param>
                /// <returns></returns>
                public static string BuildInputSQLiteTableForVolumeCalculation_Step2a(string p_strInputVolumesTable, string p_strFIAPlotTable, string p_strFIACondTable)
                {
                    return $@"UPDATE (({p_strFIAPlotTable} p INNER JOIN {p_strFIACondTable} c ON p.biosum_plot_id = c.biosum_plot_id) 
                                INNER JOIN {p_strInputVolumesTable} i ON c.biosum_cond_id = i.biosum_cond_id) 
                                SET i.precipitation=p.precipitation
                                WHERE i.fvscreatedtree_yn='Y'";
                }

                /// <summary>
                /// Update biosum_volumes_input field (balive) directly from FVSOut.db for FVS-created trees
                /// </summary>
                /// <param name="p_strInputVolumesTable"></param>
                /// <param name="p_strRunTitle"></param>
                      /// <returns></returns>
                public static string BuildInputSQLiteTableForVolumeCalculation_Step1a(string p_strInputVolumesTable, string p_strRunTitle)
                {
                    return $@"UPDATE {p_strInputVolumesTable} SET balive = 
                            (SELECT s.ba FROM fvsout.FVS_Summary AS s, fvsout.FVS_Cases c
                            WHERE {p_strInputVolumesTable}.biosum_cond_id = s.StandId and biosum_volumes_input.invyr = s.Year
                            and s.CaseID = c.CaseID and c.RunTitle = '{p_strRunTitle}'
                            and FvsCreatedTree_YN = 'Y') 
                            WHERE EXISTS (SELECT s.ba
                                FROM fvsout.FVS_Summary AS s, fvsout.FVS_Cases c
                                WHERE biosum_volumes_input.biosum_cond_id = s.StandId and biosum_volumes_input.invyr = s.Year 
                                and s.CaseID = c.CaseID and c.RunTitle = '{p_strRunTitle}' and FvsCreatedTree_YN = 'Y')";
                }

                /// <summary>
                /// Update tree fields to have values other than null. Also assign the VOL_LOC_GRP value from the condition table.
                /// </summary>
                /// <param name="p_strInputVolumesTable"></param>
                /// <param name="p_strFIACondTable"></param>
                /// <returns></returns>
                public static string BuildInputTableForVolumeCalculation_Step3(string p_strInputVolumesTable, string p_strFIACondTable)
                {
                    return $@"UPDATE {p_strInputVolumesTable} i INNER JOIN {p_strFIACondTable} c 
                                ON i.biosum_cond_id=c.biosum_cond_id 
                            SET i.vol_loc_grp=c.vol_loc_grp,
                                i.statuscd=IIF(i.statuscd IS NULL,1,i.statuscd),
                                i.cull=IIF(i.cull IS NULL,0,i.cull),
                                i.roughcull=IIF(i.roughcull IS NULL,0,i.roughcull),
                                i.decaycd=IIF(i.decaycd IS NULL,0,i.decaycd)";
                }

                /// <summary>
                /// Create and Populate a CULL work table with TOTALCULL (cull + roughcull)
                /// </summary>
                /// <returns></returns>
                public static string BuildInputTableForVolumeCalculation_Step4(string p_strCullTable, string p_strInputVolumesTable)
                {
                    //TODO: does inputVolumesTable have an id column? yes. What is it set to? How is id populated? FVS_TREE.ID?
                    string columns = "id, IIF(cull IS NOT NULL AND roughcull IS NOT NULL, cull + roughcull, IIF(cull IS NOT NULL,cull, IIF(roughcull IS NOT NULL, roughcull,0))) AS totalcull";
                    return $@"SELECT {columns} INTO {p_strCullTable} FROM {p_strInputVolumesTable}";
                }

                public class PNWRS
                {
                    public PNWRS()
                    {
                    }

                    /// <summary>
                    /// Update the TREECLCD column for all trees 
                    /// by comparing SPCD,ROUGHCULL,STATUSCD and TOTALCULL columns
                    /// </summary>
                    /// <param name="p_strCullTable"></param>
                    /// <param name="p_strInputVolumesTable"></param>
                    /// <returns></returns>
                    public static string BuildInputTableForVolumeCalculation_Step5(string p_strCullTable, string p_strInputVolumesTable)
                    {
                        return $@"UPDATE {p_strInputVolumesTable} a INNER JOIN {p_strCullTable} b ON a.id=b.id
                                   SET a.treeclcd= IIF(a.SpCd IN (62,65,66,106,133,138,304,321,322,475,756,758,990),3,
                                       IIF(a.StatusCd=2,3,
                                           IIF(b.totalcull < 75,2,
                                               IIF(a.roughcull > 37.5,3,4))))";
                    }

                    /// <summary>
                    /// Update the TREECLCD column for TREECLCD=3 and dead trees 
                    /// by comparing SPCD,DBH,STATUSCD and DECAYCD columns
                    /// </summary>
                    /// <param name="p_strCullTable"></param>
                    /// <param name="p_strInputVolumesTable"></param>
                    /// <returns></returns>
                    public static string BuildInputTableForVolumeCalculation_Step6(string p_strCullTable, string p_strInputVolumesTable)
                    {
                        return $@"UPDATE {p_strInputVolumesTable} a INNER JOIN {p_strCullTable} b ON a.id=b.id
                               SET a.treeclcd=IIF(a.DecayCd > 1,4,IIF(a.dbh < 9 AND a.SpCd < 300,4,a.treeclcd)) 
                               WHERE a.treeclcd=3 AND a.statuscd=2 AND a.SpCd NOT IN (62,65,66,106,133,138,304,321,322,475,756,758,990)";
                    }
                }

                /// <summary>
                /// Insert into the MS Access Biosum Volume table
                /// the formatted data in the input volumes table.
                /// This extra step is needed before importing to 
                /// Oracle because the performance of formatting of data from Access to 
                /// the Oracle Linked table is slow.
                /// </summary>
                /// <param name="p_strInputVolumesTable"></param>
                /// <param name="p_strOracleBiosumVolumesTable"></param>
                /// <returns></returns>
                public static string BuildInputTableForVolumeCalculation_Step7(string p_strInputVolumesTable, string p_strBiosumVolumesTable)
                {
                    string strColumns = "STATECD,COUNTYCD,PLOT,INVYR,VOL_LOC_GRP,TREE,SPCD,DIA,HT," +
                                        "ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL,DECAYCD,TOTAGE,TRE_CN,CND_CN,PLT_CN, DIAHTCD, BALIVE, PRECIPITATION";

                    string strValues =
                        "CINT(MID(BIOSUM_COND_ID,6,2)) AS STATECD,CINT(MID(BIOSUM_COND_ID,12,3)) AS COUNTYCD,CINT(MID(BIOSUM_COND_ID,16,5)) AS PLOT," +
                        "INVYR,VOL_LOC_GRP,ID AS TREE,SPCD,DBH AS DIA,HT,ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL,DECAYCD,TOTAGE," +
                        "CSTR(ID) AS TRE_CN,BIOSUM_COND_ID AS CND_CN,MID(BIOSUM_COND_ID,1,LEN(BIOSUM_COND_ID)-1) AS PLT_CN, DIAHTCD, BALIVE, PRECIPITATION";

                    return $@"INSERT INTO {p_strBiosumVolumesTable} ({strColumns}) SELECT {strValues} FROM {p_strInputVolumesTable}";
                }

                public static string BuildInputSQLiteBiosumCalcTable_Step7(string p_strInputVolumesTable, string p_strBiosumVolumesTable)
                {
                    string strColumns = "STATECD,COUNTYCD,PLOT,INVYR,VOL_LOC_GRP,TREE,SPCD,DIA,HT," +
                                        "ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL,DECAYCD,TOTAGE,TRE_CN,CND_CN,PLT_CN, DIAHTCD, BALIVE, PRECIPITATION, STANDING_DEAD_CD";

                    string strValues =
                        "CAST((SUBSTR(BIOSUM_COND_ID,6,2)) AS INTEGER) AS STATECD, CAST((SUBSTR(BIOSUM_COND_ID,12,3)) AS INTEGER) AS COUNTYCD,CAST((SUBSTR(BIOSUM_COND_ID,16,5)) AS INTEGER) AS PLOT," +
                        "INVYR, TRIM(VOL_LOC_GRP),ID AS TREE,SPCD,DBH AS DIA,ROUND(HT),ROUND(ACTUALHT),CR,STATUSCD,TREECLCD,ROUGHCULL,CULL,DECAYCD,TOTAGE," +
                        "CAST(ID AS TEXT) AS TRE_CN,BIOSUM_COND_ID AS CND_CN,SUBSTR(BIOSUM_COND_ID,1,LENGTH(BIOSUM_COND_ID)-1) AS PLT_CN, DIAHTCD, BALIVE, PRECIPITATION, STANDING_DEAD_CD";

                    return $@"INSERT INTO {p_strBiosumVolumesTable} ({strColumns}) SELECT {strValues} FROM {p_strInputVolumesTable}";
                }


                /// <summary>
                /// Insert records into the fcs oracle linked table
                /// </summary>
                /// <param name="p_strBiosumVolumesTable"></param>
                /// <param name=""></param>
                /// <returns></returns>
                public static string BuildInputTableForVolumeCalculation_Step8(string p_strBiosumVolumesTable, string p_strOracleFCSBiosumVolumesLinkedTable)
                {
                    string strColumns = @"STATECD,COUNTYCD,PLOT,INVYR,VOL_LOC_GRP,TREE,SPCD,DIA,HT,ACTUALHT,CR,
                                        STATUSCD,TREECLCD,ROUGHCULL,CULL,DECAYCD,TOTAGE,TRE_CN,CND_CN,PLT_CN";
                    return $@"INSERT INTO {p_strOracleFCSBiosumVolumesLinkedTable} ({strColumns}) 
                                SELECT {strColumns} FROM {p_strBiosumVolumesTable}";
                }

                /// <summary>
                /// Update the FVS_TREE table with the volumes and biomass
                /// values that Oracle returned
                /// </summary>
                /// <param name="p_strFvsTreeTable"></param>
                /// <param name="p_strOracleBiosumVolumesTable"></param>
                /// <returns></returns>
                public static string BuildInputTableForVolumeCalculation_Step9(string p_strFvsTreeTable, string p_strBiosumCalcOutputTable)
                {
                    return $@"UPDATE {p_strFvsTreeTable} f
                           INNER JOIN {p_strBiosumCalcOutputTable} o
                           ON f.id = o.tree 
                           SET f.drybio_bole=IIF(o.DRYBIO_BOLE_CALC IS NOT NULL, o.DRYBIO_BOLE_CALC, null),
                           f.drybio_sapling=IIF(o.DRYBIO_SAPLING_CALC IS NOT NULL, o.DRYBIO_SAPLING_CALC, null),
                           f.drybio_top=IIF(o.DRYBIO_TOP_CALC IS NOT NULL, o.DRYBIO_TOP_CALC, null),
                           f.drybio_wdld_spp=IIF(o.DRYBIO_WDLD_SPP_CALC IS NOT NULL, o.DRYBIO_WDLD_SPP_CALC, null),
                           f.drybiom=o.DRYBIOM_CALC,
                           f.drybiot=o.DRYBIOT_CALC,
                           f.volcfgrs=o.VOLCFGRS_CALC,
                           f.volcfnet=o.VOLCFNET_CALC,
                           f.volcfsnd=IIF(o.VOLCFSND_CALC IS NOT NULL, o.VOLCFSND_CALC, null),
                           f.volcsgrs=o.VOLCSGRS_CALC,
                           f.voltsgrs=o.VOLTSGRS_CALC ";
                }

                /// <summary>
                /// Update the FVS_TREE table with the volumes and biomass
                /// values that FCS returned
                /// </summary>
                /// <param name="p_strFvsTreeTable"></param>
                /// <param name="p_strBiosumCalcOutputTable"></param>
                /// <returns></returns>
                public static string BuildInputSQLiteTableForVolumeCalculation_Step9(string p_strFvsTreeTable, string p_strBiosumCalcOutputTable)
                {
                    return $@"UPDATE {p_strFvsTreeTable} 
                           SET (drybio_bole, drybio_sapling, drybio_top, drybio_wdld_spp,
                               drybiom, drybiot, volcfgrs) 
                               = (select drybio_bole_calc, drybio_sapling_calc, drybio_top_calc, drybio_wdld_spp_calc,
                                 drybiom_calc, drybiot_calc, volcfgrs_calc FROM FCS.{p_strBiosumCalcOutputTable}
                                 WHERE {p_strFvsTreeTable}.id = FCS.{p_strBiosumCalcOutputTable}.tree)";
                }

                /// <summary>
                /// Update the FVS_TREE table with the volumes and biomass
                /// values that FCS returned
                /// </summary>
                /// <param name="p_strFvsTreeTable"></param>
                /// <param name="p_strBiosumCalcOutputTable"></param>
                /// <returns></returns>
                public static string BuildInputSQLiteTableForVolumeCalculation_Step10(string p_strFvsTreeTable, string p_strBiosumCalcOutputTable)
                {
                    return $@"UPDATE {p_strFvsTreeTable} 
                           SET (volcfnet, volcfsnd, volcsgrs, voltsgrs,
                               standing_dead_cd, statuscd, decaycd) 
                               = (select volcfnet_calc, volcfsnd_calc, volcsgrs_calc, voltsgrs_calc,
                                 standing_dead_cd, statuscd, decaycd FROM FCS.{p_strBiosumCalcOutputTable}
                                 WHERE {p_strFvsTreeTable}.id = FCS.{p_strBiosumCalcOutputTable}.tree)";
                    // Note: standing_dead_cd, statuscd, decaycd aren't calculated by FCS but this was an easy way to populate it from master.tree
                }

                /// <summary>
                /// Update the FVS_TREE table with the volumes and biomass for Oracle XE
                /// values that Oracle returned
                /// </summary>
                /// <param name="p_strFvsTreeTable"></param>
                /// <param name="p_strOracleBiosumVolumesTable"></param>
                /// <returns></returns>
                public static string BuildInputTableForVolumeCalculationXE_Step9(string p_strFvsTreeTable, string p_strOracleBiosumVolumesTable)
                {
                    return $@"UPDATE {p_strFvsTreeTable} f INNER JOIN {p_strOracleBiosumVolumesTable} o ON f.id = o.tree 
                            SET f.volcsgrs=o.VOLCSGRS_CALC,
                                f.volcfgrs=o.VOLCFGRS_CALC,
                                f.volcfnet=o.VOLCFNET_CALC,
                                f.drybiot=o.DRYBIOT_CALC,
                                f.drybiom=o.DRYBIOM_CALC,
                                f.voltsgrs=o.VOLTSGRS_CALC";
                }

                public static string[] PopulateRxFields(string p_strFvsTreeTable, string p_strSeqNumTable)
                {
                    string[] arrQueries = new string[2];
                    arrQueries[0] = $@"UPDATE {p_strFvsTreeTable}                            
                                                     SET rxcycle = (SELECT CASE WHEN CYCLE1_PRE_YN ='Y' THEN '1' 
                                                     WHEN CYCLE1_POST_YN ='Y' THEN '1'
                                                     WHEN CYCLE2_PRE_YN ='Y' THEN '2' 
                                                     WHEN CYCLE2_POST_YN ='Y' THEN '2'
                                                     WHEN CYCLE3_PRE_YN ='Y' THEN '3' 
                                                     WHEN CYCLE3_POST_YN ='Y' THEN '3'
                                                     WHEN CYCLE4_PRE_YN ='Y' THEN '4' 
                                                     WHEN CYCLE4_POST_YN ='Y' THEN '4'  
                                                     ELSE NULL END
                                                     FROM FVS_SUMMARY_PREPOST_SEQNUM_MATRIX m WHERE 
                                                     m.STANDID = cutlist_fia_trees_work_table.biosum_cond_id and
                                                     cutlist_fia_trees_work_table.rxpackage = m.RXPACKAGE and 
                                                     cutlist_fia_trees_work_table.Year = m.YEAR )
                                                     WHERE EXISTS ( SELECT * from FVS_SUMMARY_PREPOST_SEQNUM_MATRIX m 
                                                     WHERE m.STANDID = cutlist_fia_trees_work_table.biosum_cond_id and
                                                     cutlist_fia_trees_work_table.rxpackage = m.RXPACKAGE)";

                    return arrQueries;
                }
            }
        }

        public class TravelTime
		{
            public string m_strTravelTimeTable;
            private bool _bLoadDataSources = true;
            private Queries _oQueries=null;
			public TravelTime()
			{
			}
            public Queries ReferenceQueries
			{
				get {return _oQueries;}
				set {_oQueries=value;}
			}
            public bool LoadDatasource
            {
                get { return _bLoadDataSources; }
                set { _bLoadDataSources = value; }
            }
			

            public void LoadDatasources()
            {
                m_strTravelTimeTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("PLOT");


                if (this.m_strTravelTimeTable.Trim().Length == 0)
                {

                    MessageBox.Show("!!Could Not Locate Travel Times Table!!", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    ReferenceQueries.m_intError = -1;
                    return;
                }
            }
		}
		public class Audit
		{
			

			public Audit()
			{
			}

		}
		public class FIAPlot
		{
			public string m_strPlotTable;
			public string m_strTreeTable;
			public string m_strCondTable;
			public string m_strSiteIndexTable;
			public string m_strTreeRegionalBiomassTable;
			public string m_strPopEvalTable;
			public string m_strPopEstnUnitTable;
			public string m_strPopStratumTable;
			public string m_strPPSATable;
            public string m_strBiosumPopStratumAdjustmentFactorsTable;


			private Queries _oQueries=null;	
			private bool _bLoadDataSources=true;
			string m_strSql="";

			public FIAPlot()
			{
			}

			public Queries ReferenceQueries
			{
				get {return _oQueries;}
				set {_oQueries=value;}
			}
			public bool LoadDatasource
			{
				get {return _bLoadDataSources;}
				set {_bLoadDataSources=value;}
			}


			public void LoadDatasources()
			{
				m_strPlotTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("PLOT");
				m_strTreeTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("TREE");
				m_strCondTable= ReferenceQueries.m_oDataSource.getValidDataSourceTableName("CONDITION");
				
			
				if (this.m_strPlotTable.Trim().Length == 0) 
				{
					
					MessageBox.Show("!!Could Not Locate Plot Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strCondTable.Trim().Length == 0)
				{
					MessageBox.Show("!!Could Not Locate Condition Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				if (this.m_strTreeTable.Trim().Length == 0 && this._oQueries._strScenarioType!="optimizer")
				{
					MessageBox.Show("!!Could Not Locate Tree Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}
				
			
			}
            
            static public string[] FIADBPlotInput_CalculateAdjustmentFactorsSQL(
                string p_strPpsaTable,
                string p_strPopEstUnitTable,
                string p_strPopStratumTable,
                string p_strPopEvalTable,
                string p_strFIADBPlotTable,
                string p_strFIADBCondTable,
                string p_strRsCd,
                string p_strEvalId,
                string p_strCondProp)
            {
                string[] strSQL = new string[21];

                //
                //CREATE THE BIOSUM PLOT TABLE WITH 
                //ONLY THE CURRENTLY SELECTED EVALUATION
                //
                strSQL[0] = "CREATE TABLE BIOSUM_PLOT AS " +
                            "SELECT p.* " + 
                            "FROM " + p_strFIADBPlotTable + " p, " + 
                                "(SELECT DISTINCT PLT_CN " + 
                                 "FROM " + p_strPpsaTable + " " + 
                                 "WHERE RSCD=" + p_strRsCd + " AND EVALID=" + p_strEvalId + ") ppsa " + 
                            "WHERE p.CN = ppsa.PLT_CN";
                //
                //CREATE BIOSUM_PPSA
                //
                strSQL[1] = "CREATE TABLE BIOSUM_PPSA AS " +
                            "SELECT ppsa.*,p.CYCLE,p.SUBCYCLE " + 
                            "FROM " + p_strPpsaTable + " ppsa " + 
                            "INNER JOIN BIOSUM_PLOT p " + 
                            "ON ppsa.PLT_CN=p.CN " + 
                            "WHERE RSCD=" + p_strRsCd + " AND EVALID=" + p_strEvalId;
                //
                //CREATE BIOSUM COND TABLE
                //
                strSQL[2] = "CREATE TABLE BIOSUM_COND AS " +
                            "SELECT c.* FROM " + p_strFIADBCondTable + " c," + 
                            "(SELECT CN FROM BIOSUM_PLOT) p " + 
                            "WHERE c.PLT_CN = p.CN";

                //change hazardous condition to sampled
                strSQL[3] = "UPDATE BIOSUM_COND SET cond_status_cd = 1 " +
                            "WHERE COND_NONSAMPLE_REASN_CD = 5";

                //update condition satatus to NONSAMPLED if the condition proportion is less than .25
                strSQL[4] = "UPDATE BIOSUM_COND SET cond_status_cd = 5 " +
                            "WHERE cond_status_cd = 1 AND condprop_unadj < ." + p_strCondProp;

                //join pop_estn_unit,pop_stratum,pop_eval tables into biosum_eus_temp
                strSQL[5] = "CREATE TABLE BIOSUM_EUS_TEMP AS " +
                            "SELECT pe.rscd, pe.evalid,ps.estn_unit,ps.stratumcd," +
                                   "pe.eval_descr,peu.estn_unit_descr,peu.arealand_eu," +
                                   "peu.areatot_eu , ps.p1pointcnt, ps.p2pointcnt," +
                                   "peu.p1pntcnt_eu as p1pointcnt_eu,peu.area_used," +
                                   "ps.adj_factor_macr,ps.adj_factor_subp," +
                                   "ps.adj_factor_micr,ps.expns,pe.LAND_ONLY " +
                            "FROM " + p_strPopEstUnitTable + " peu," +
                                      p_strPopEvalTable + " pe," + p_strPopStratumTable + " ps " +
                                      "WHERE  ((pe.rscd=" + p_strRsCd + " AND pe.EVALID=" + p_strEvalId + ")) AND " +
                                              "(pe.rscd = ps.rscd AND pe.evalid = ps.evalid) AND " +
                                              "(ps.rscd = peu.rscd AND ps.evalid = peu.evalid AND " +
                                               "ps.estn_unit = peu.estn_unit)";

                //
                //SUM UP UNADJUSTED FACTORS FOR DENIED ACCESS
                //
                strSQL[6] = "CREATE TABLE BIOSUM_PPSA_DENIED_ACCESS AS " +
                            "SELECT DISTINCT ppsa.evalid, ppsa.estn_unit, ppsa.statecd, ppsa.stratumcd, ppsa.plot," +
                                            "ppsa.countycd, ppsa.subcycle, ppsa.cycle, ppsa.unitcd," +
                            " SUM(CASE WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD IN(1, 2, 3, 4) THEN 0" +
                            " WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD NOT IN(1, 2, 3, 4) THEN c.MACRPROP_UNADJ" +
                            " WHEN eus.LAND_ONLY <> 'N' AND c.COND_STATUS_CD IN(1, 2, 3) THEN 0" +
                            " ELSE c.MACRPROP_UNADJ END) AS denied_macr," +
                            " SUM(CASE WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD IN(1, 2, 3, 4) THEN 0" +
                            " WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD NOT IN(1, 2, 3, 4) THEN c.MICRPROP_UNADJ" +
                            " WHEN eus.LAND_ONLY <> 'N' AND c.COND_STATUS_CD IN(1, 2, 3) THEN 0" +
                            " ELSE c.MICRPROP_UNADJ END) AS denied_micr," +
                            " SUM(CASE WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD IN(1, 2, 3, 4) THEN 0" +
                            " WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD NOT IN(1, 2, 3, 4) THEN c.SUBPPROP_UNADJ" +
                            " WHEN eus.LAND_ONLY <> 'N' AND c.COND_STATUS_CD IN(1, 2, 3) THEN 0" +
                            " ELSE c.SUBPPROP_UNADJ END) AS denied_subp," +
                            " SUM(CASE WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD IN(1, 2, 3, 4) THEN 0" +
                            " WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD NOT IN(1, 2, 3, 4) THEN c.CONDPROP_UNADJ" +
                            " WHEN eus.LAND_ONLY <> 'N' AND c.COND_STATUS_CD IN(1, 2, 3) THEN 0" +
                            " ELSE c.CONDPROP_UNADJ END) AS denied_cond" +
                            " FROM BIOSUM_PPSA ppsa," +
                                 "BIOSUM_COND c," +
                                 "BIOSUM_EUS_TEMP eus " +
                            "WHERE (ppsa.plt_cn = c.plt_cn) AND " +
                                  "(ppsa.rscd = eus.rscd AND " +
                                   "ppsa.evalid = eus.evalid AND " +
                                   "ppsa.stratumcd = eus.stratumcd AND " +
                                   "ppsa.estn_unit = eus.estn_unit) " +
                            "GROUP BY ppsa.evalid, ppsa.estn_unit, ppsa.statecd, ppsa.stratumcd," +
                                     "ppsa.plot, ppsa.countycd, ppsa.subcycle, ppsa.cycle, ppsa.unitcd";

                //DELETE DENIED_COND=1
                strSQL[7] = "DELETE FROM BIOSUM_PPSA_DENIED_ACCESS WHERE DENIED_COND =  1";
                //JOIN THE 2 TABLES
                strSQL[8] = "CREATE TABLE BIOSUM_PPSA_TEMP AS " +
                            "SELECT ppsa.*," +
                                   "denied.denied_macr," +
                                   "denied.denied_micr," +
                                   "denied.denied_subp," +
                                   "denied.denied_cond " +
                            "FROM BIOSUM_PPSA_DENIED_ACCESS denied," +
                                 "BIOSUM_PPSA ppsa " +
                           "WHERE ppsa.evalid = denied.evalid AND " +
                                 "ppsa.estn_unit = denied.estn_unit AND " +
                                 "ppsa.stratumcd = denied.stratumcd AND " +
                                 "ppsa.plot = denied.plot AND " +
                                 "ppsa.statecd = denied.statecd AND " +
                                 "ppsa.countycd  = denied.countycd AND " +
                                 "ppsa.subcycle  = denied.subcycle AND " +
                                 "ppsa.cycle     = denied.cycle AND " +
                                 "ppsa.unitcd    = denied.unitcd";
                //
                //CALCULATE ADJUSTMENTS
                //
                strSQL[9] = "CREATE TABLE BIOSUM_EUS_ACCESS AS" +
                                " SELECT DISTINCT " +
                                "eus.rscd, eus.evalid, eus.estn_unit, eus.stratumcd," +
                                "eus.arealand_eu, eus.areatot_eu," +
                                "eus.area_used, rowcount.p2pointcnt_man as p2pointcnt_man," +
                                "eus.p1pointcnt, eus.p1pointcnt_eu, eus.p2pointcnt," +
                                " SUM(c.MACRPROP_UNADJ *" +
                                " CASE WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD IN(1, 2, 3, 4) THEN 1" +
                                " WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD NOT IN(1, 2, 3, 4) THEN 0" +
                                " WHEN eus.LAND_ONLY <> 'N' AND c.COND_STATUS_CD IN(1, 2, 3) THEN 1" +
                                " ELSE 0 END) / SUM(c.macrprop_unadj) as pmh_macr," +
                                " SUM(c.MICRPROP_UNADJ *" +
                                " CASE WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD IN(1, 2, 3, 4) THEN 1" +
                                " WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD NOT IN(1, 2, 3, 4) THEN 0" +
                                " WHEN eus.LAND_ONLY <> 'N' AND c.COND_STATUS_CD IN(1, 2, 3) THEN 1 " +
                                " ELSE 0 END) / SUM(c.MICRPROP_unadj) as pmh_micr," +
                                " SUM(c.SUBPPROP_UNADJ *" +
                                " CASE WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD IN(1, 2, 3, 4) THEN 1" +
                                " WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD NOT IN(1, 2, 3, 4) THEN 0" +
                                " WHEN eus.LAND_ONLY <> 'N' AND c.COND_STATUS_CD IN(1, 2, 3) THEN 1" +
                                " ELSE 0 END) / SUM(c.SUBPPROP_unadj) as pmh_sub," +
                                " SUM(c.CONDPROP_UNADJ *" +
                                " CASE WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD IN(1, 2, 3, 4) THEN 1" +
                                " WHEN eus.LAND_ONLY = 'N' AND c.COND_STATUS_CD NOT IN(1, 2, 3, 4) THEN 0" +
                                " WHEN eus.LAND_ONLY <> 'N' AND c.COND_STATUS_CD IN(1, 2, 3) THEN 1" +
                                " ELSE 0 END) / SUM(c.CONDPROP_unadj) as pmh_cond " +
                            "FROM BIOSUM_COND c," +
                                 "BIOSUM_PPSA_TEMP ppsa," +
                                 "BIOSUM_EUS_TEMP  eus," +
                                "(SELECT eus_count.rscd," +
                                        "eus_count.evalid," +
                                        "eus_count.estn_unit," +
                                        "eus_count.stratumcd," +
                                        "COUNT(ppsa_count.PLT_CN) as p2pointcnt_man " +
                                 "FROM BIOSUM_PPSA_TEMP ppsa_count," +
                                      "BIOSUM_EUS_TEMP eus_count " +
                                 "WHERE (ppsa_count.rscd = eus_count.rscd AND " +
                                        "ppsa_count.evalid = eus_count.evalid AND " +
                                        "ppsa_count.stratumcd = eus_count.stratumcd AND " +
                                        "ppsa_count.estn_unit = eus_count.estn_unit) " +
                                 "GROUP BY eus_count.rscd," +
                                          "eus_count.evalid," +
                                          "eus_count.estn_unit," +
                                          "eus_count.stratumcd " +
                                ") AS ROWCOUNT " +
                            "WHERE (ppsa.plt_cn = c.plt_cn) AND " +
                                  "(ppsa.rscd = eus.rscd AND " +
                                   "ppsa.evalid = eus.evalid AND " +
                                   "ppsa.stratumcd = eus.stratumcd AND " +
                                   "ppsa.estn_unit = eus.estn_unit) AND " +
                                  "(eus.rscd = rowcount.rscd AND " +
                                   "eus.evalid = rowcount.evalid AND " +
                                   "eus.estn_unit = rowcount.estn_unit AND " +
                                   "eus.stratumcd = rowcount.stratumcd) " +
                           "GROUP BY eus.rscd," +
                                    "eus.evalid," +
                                    "eus.estn_unit," +
                                    "eus.stratumcd," +
                                    "eus.arealand_eu," +
                                    "eus.areatot_eu," +
                                    "eus.area_used," +
                                    "eus.p1pointcnt," +
                                    "eus.p1pointcnt_eu," +
                                    "eus.p2pointcnt," +
                                    "p2pointcnt_man";

                strSQL[10] = "ALTER TABLE BIOSUM_EUS_ACCESS ADD COLUMN STRATUM_AREA DOUBLE";
                strSQL[11] = "ALTER TABLE BIOSUM_EUS_ACCESS ADD COLUMN DOUBLE_SAMPLING INTEGER";
                //
                //CALCULATE STRATUM AREA
                //
                strSQL[12] = "UPDATE BIOSUM_EUS_ACCESS " +
                                "SET double_sampling = " +
                                "CASE WHEN p1pointcnt_eu is null OR p1pointcnt is null OR p1pointcnt = p1pointcnt_eu THEN 0 " +
                                "ELSE 1 " +
                                "END, " +
                                "stratum_area = " +
                                "CASE WHEN p1pointcnt_eu is NOT null AND p1pointcnt_eu > 0 THEN area_used*p1pointcnt / p1pointcnt_eu " +
                                "ELSE 0 " +
                                "END";
                //
                //MERGE BIOSUM_EUS_ACCESS WITH BIOSUM_EUS_TEMP INTO  biosum_pop_stratum_adjustment_factors
                //
                strSQL[13] = "CREATE TABLE biosum_pop_stratum_adjustment_factors AS " +
                            "SELECT a.rscd,a.evalid," +
                                  "a.estn_unit,a.stratumcd," +
                                  "a.p2pointcnt_man,a.stratum_area," +
                                  "a.double_sampling," +
                                  "a.pmh_macr,a.pmh_micr," +
                                  "a.pmh_sub,a.pmh_cond," +
                                  "b.eval_descr,b.estn_unit_descr," +
                                  "b.adj_factor_macr, b.adj_factor_subp," +
                                  "b.adj_factor_micr, b.expns " +
                          "FROM BIOSUM_EUS_ACCESS a " +
                          "INNER JOIN BIOSUM_EUS_TEMP b " +
                          "ON a.rscd = b.rscd AND " +
                             "a.evalid = b.evalid AND " +
                             "a.estn_unit = b.estn_unit AND " +
                             "a.stratumcd = b.stratumcd";

                strSQL[14] = "ALTER TABLE biosum_pop_stratum_adjustment_factors ADD COLUMN stratum_cn CHAR(34)";
                //
                //UPDATE THE  biosum_pop_stratum_adjustment_factors TABLE 
                //WITH THE KEY COLUMN FROM THE POP_STRATUM TABLE
                //
                strSQL[15] = "UPDATE biosum_pop_stratum_adjustment_factors" +
                    " SET (STRATUM_CN) =" +
                    " (SELECT POP_STRATUM.cn FROM " + p_strPopStratumTable +
                    " WHERE " + p_strPopStratumTable + ".RSCD = biosum_pop_stratum_adjustment_factors.RSCD" +
                    " AND " + p_strPopStratumTable + ".EVALID = biosum_pop_stratum_adjustment_factors.EVALID" +
                    " AND " + p_strPopStratumTable + ".ESTN_UNIT = biosum_pop_stratum_adjustment_factors.ESTN_UNIT" +
                    " AND " + p_strPopStratumTable + ".STRATUMCD = biosum_pop_stratum_adjustment_factors.STRATUMCD" +
                    " AND biosum_pop_stratum_adjustment_factors.RSCD = " + p_strRsCd +
                    " AND biosum_pop_stratum_adjustment_factors.EVALID = " + p_strEvalId + ")" +
                    " WHERE EXISTS( " +
                    " SELECT * FROM POP_STRATUM" +
                    " WHERE " + p_strPopStratumTable + ".RSCD = biosum_pop_stratum_adjustment_factors.RSCD" +
                    " AND " + p_strPopStratumTable + ".EVALID = biosum_pop_stratum_adjustment_factors.EVALID" +
                    " AND " + p_strPopStratumTable + ".ESTN_UNIT = biosum_pop_stratum_adjustment_factors.ESTN_UNIT" +
                    " AND " + p_strPopStratumTable + ".STRATUMCD = biosum_pop_stratum_adjustment_factors.STRATUMCD" +
                    " AND biosum_pop_stratum_adjustment_factors.RSCD = " + p_strRsCd +
                    " AND biosum_pop_stratum_adjustment_factors.EVALID = " + p_strEvalId + ")";

                //
                //CLEAN UP
                //
                strSQL[16] = "DROP TABLE BIOSUM_PPSA";
                strSQL[17] = "DROP TABLE BIOSUM_EUS_TEMP";
                strSQL[18] = "DROP TABLE BIOSUM_PPSA_DENIED_ACCESS";
                strSQL[19] = "DROP TABLE BIOSUM_PPSA_TEMP";
                strSQL[20] = "DROP TABLE BIOSUM_EUS_ACCESS";
                return strSQL;
            }
             
			
		}
		public class Processor
		{
            private Queries _oQueries = null;
            private bool _bLoadDataSources = true;
            public string m_strAdditionalHarvestCostsTable;
			public Processor()
			{
			}
            public Queries ReferenceQueries
            {
                get { return _oQueries; }
                set { _oQueries = value; }
            }
            public bool LoadDatasource
            {
                get { return _bLoadDataSources; }
                set { _bLoadDataSources = value; }
            }
            public void LoadDatasources()
            {
              
            }
           

			public static string AuditFvsOut_SelectIntoUnionOfFVSTreeTables(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strIntoTable,RxPackageItem_Collection p_oRxPackageItem_Collection, string[] p_strFVSVariantsArray,string p_strColumnList)
			{
               
				string strSql="SELECT DISTINCT " + p_strColumnList +  " " + 
					          "INTO " + p_strIntoTable + " " + 
					          "FROM (";
				int x,y;
				for (x=0;x<=p_strFVSVariantsArray.Length-1;x++)
				{
                    for (y = 0; y <= p_oRxPackageItem_Collection.Count - 1; y++)
                    {
                        if (p_oAdo.TableExist(p_oConn, "fvs_tree_IN_" + p_strFVSVariantsArray[x].Trim() + "_P" + p_oRxPackageItem_Collection.Item(y).RxPackageId + "_TREE_CUTLIST"))
                        {
                            strSql = strSql + "SELECT " + p_strColumnList + " " +
                                              "FROM fvs_tree_IN_" + p_strFVSVariantsArray[x].Trim() + "_P" + p_oRxPackageItem_Collection.Item(y).RxPackageId + "_TREE_CUTLIST UNION ";
                        }
                    }
				}
				if (strSql.IndexOf("UNION",0) > 0) strSql = strSql.Substring(0,strSql.Length - 6) + ")";
               
				return strSql;
				
			}
            public static List<string> AuditFvsOut_SelectIntoUnionOfFVSTreeTablesUsingListArray(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strIntoTable, string p_strColumnList)
            {
                List<string> strList = new List<string>();
                string strSql = "";

                // This is an update from the original implementation when all the cut lists were in separate databases
                // No more need to generate multiple sql statements by variant, rxpackage
                if (p_oAdo.TableExist(p_oConn, Tables.FVS.DefaultFVSCutTreeTableName))
                {
                    strSql = $@"SELECT count(*) FROM {Tables.FVS.DefaultFVSCutTreeTableName}";
                    long lngCount = p_oAdo.getRecordCount(p_oAdo.m_OleDbConnection, strSql, Tables.FVS.DefaultFVSCutTreeTableName);
                    if (lngCount > 0)
                    {
                        strSql = $@"SELECT DISTINCT { p_strColumnList}
                        INTO {p_strIntoTable} FROM {Tables.FVS.DefaultFVSCutTreeTableName}";
                        strList.Add(strSql);
                    }                           
                }
                return strList;
            }
			
		}
        public class ProcessorScenarioRun
        {
            private string strSQL = "";
            private string _strScenarioId = "";
            public string ScenarioId
            {
                get { return _strScenarioId; }
                set { _strScenarioId = value; }
            }
            public ProcessorScenarioRun()
            {
            }
            public static string InitializeBinsTable(string p_strTreeBinsTableName,
                                                     string p_strClass)
            {
                return "UPDATE " + p_strTreeBinsTableName + " b " +
                        "SET b." + p_strClass + "_Util_Logs_count = 0," +
                        "b." + p_strClass + "_Util_Logs_TPA = 0," +
                        "b." + p_strClass + "_Util_Chips_count = 0," +
                        "b." + p_strClass + "_Util_Chips_TPA = 0," +
                        "b." + p_strClass + "_Util_Logs_merch_wt = 0," +
                        "b." + p_strClass + "_Util_Logs_merch_vol = 0," +
                        "b." + p_strClass + "_Util_Logs_biomass_wt = 0," +
                        "b." + p_strClass + "_Util_Logs_biomass_vol = 0," +
                        "b." + p_strClass + "_Util_Chips_merch_wt = 0," +
                        "b." + p_strClass + "_Util_Chips_merch_vol = 0," +
                        "b." + p_strClass + "_Util_Chips_biomass_wt = 0," +
                        "b." + p_strClass + "_Util_Chips_biomass_vol = 0," +
                        "b." + p_strClass + "_NonUtil_Logs_count = 0," +
                        "b." + p_strClass + "_NonUtil_Logs_TPA = 0," +
                        "b." + p_strClass + "_NonUtil_Chips_count = 0," +
                        "b." + p_strClass + "_NonUtil_Chips_TPA = 0," +
                        "b." + p_strClass + "_NonUtil_Logs_merch_wt = 0," +
                        "b." + p_strClass + "_NonUtil_Logs_merch_vol = 0," +
                        "b." + p_strClass + "_NonUtil_Logs_biomass_wt = 0," +
                        "b." + p_strClass + "_NonUtil_Logs_biomass_vol = 0," +
                        "b." + p_strClass + "_NonUtil_Chips_merch_wt = 0," +
                        "b." + p_strClass + "_NonUtil_Chips_merch_vol = 0," +
                        "b." + p_strClass + "_NonUtil_Chips_biomass_wt = 0," +
                        "b." + p_strClass + "_NonUtil_Chips_biomass_vol = 0";



            }

            public static string InitializeHwdBinsTable(string p_strTreeBinsTableName,
                                                     string p_strClass)
            {
                return "UPDATE " + p_strTreeBinsTableName + " b " +
                        "SET b.HWD_" + p_strClass + "_Util_Logs_count = 0," +
                        "b.HWD_" + p_strClass + "_Util_Logs_TPA = 0," +
                        "b.HWD_" + p_strClass + "_Util_Chips_count = 0," +
                        "b.HWD_" + p_strClass + "_Util_Chips_TPA = 0," +
                        "b.HWD_" + p_strClass + "_Util_Logs_merch_wt = 0," +
                        "b.HWD_" + p_strClass + "_Util_Logs_merch_vol = 0," +
                        "b.HWD_" + p_strClass + "_Util_Logs_biomass_wt = 0," +
                        "b.HWD_" + p_strClass + "_Util_Logs_biomass_vol = 0," +
                        "b.HWD_" + p_strClass + "_Util_Chips_merch_wt = 0," +
                        "b.HWD_" + p_strClass + "_Util_Chips_merch_vol = 0," +
                        "b.HWD_" + p_strClass + "_Util_Chips_biomass_wt = 0," +
                        "b.HWD_" + p_strClass + "_Util_Chips_biomass_vol = 0," + 
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_count = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_TPA = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_count = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_TPA = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_merch_wt = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_merch_vol = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_biomass_wt = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Logs_biomass_vol = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_merch_wt = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_merch_vol = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_biomass_wt = 0," +
                        "b.HWD_" + p_strClass + "_NonUtil_Chips_biomass_vol = 0";



            }

            public static string SumBinByPlotRxSpcGrpDbhGrp(string p_strFromTableName, string p_strIntoTableName)
            {
                return "SELECT BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE, species_group, diam_group," +
                        " Sum(BC_Util_Logs_count) AS BC_Util_Logs_count," +
                        " Sum(BC_Util_Logs_TPA) AS BC_Util_Logs_TPA," +
                        " Sum(BC_Util_Logs_merch_wt) AS BC_Util_Logs_merch_wt," +
                        " Sum(BC_Util_Logs_merch_vol) AS BC_Util_Logs_merch_vol," +
                        " Sum(BC_Util_Logs_biomass_vol) AS BC_Util_Logs_biomass_vol," +
                        " Sum(BC_Util_Logs_biomass_wt) AS BC_Util_Logs_biomass_wt," +
                        " Sum(BC_Util_Chips_count) AS BC_Util_Chips_count," +
                        " Sum(BC_Util_Chips_TPA) AS BC_Util_Chips_TPA," +
                        " Sum(BC_Util_Chips_merch_wt) AS BC_Util_Chips_merch_wt," +
                        " Sum(BC_Util_Chips_merch_vol) AS BC_Util_Chips_merch_vol," +
                        " Sum(BC_Util_Chips_biomass_wt) AS BC_Util_Chips_biomass_wt," +
                        " Sum(BC_Util_Chips_biomass_vol) AS BC_Util_Chips_biomass_vol," +
                        " Sum(BC_NonUtil_Logs_count) AS BC_NonUtil_Logs_count," +
                        " Sum(BC_NonUtil_Logs_TPA) AS BC_NonUtil_Logs_TPA," +
                        " Sum(BC_NonUtil_Logs_merch_wt) AS BC_NonUtil_Logs_merch_wt," +
                        " Sum(BC_NonUtil_Logs_merch_vol) AS BC_NonUtil_Logs_merch_vol," +
                        " Sum(BC_NonUtil_Logs_biomass_vol) AS BC_NonUtil_Logs_biomass_vol," +
                        " Sum(BC_NonUtil_Logs_biomass_wt) AS BC_NonUtil_Logs_biomass_wt," +
                        " Sum(BC_NonUtil_Chips_count) AS BC_NonUtil_Chips_count," +
                        " Sum(BC_NonUtil_Chips_TPA) AS BC_NonUtil_Chips_TPA," +
                        " Sum(BC_NonUtil_Chips_merch_wt) AS BC_NonUtil_Chips_merch_wt," +
                        " Sum(BC_NonUtil_Chips_merch_vol) AS BC_NonUtil_Chips_merch_vol," +
                        " Sum(BC_NonUtil_Chips_biomass_wt) AS BC_NonUtil_Chips_biomass_wt," +
                        " Sum(BC_NonUtil_Chips_biomass_vol) AS BC_NonUtil_Chips_biomass_vol," +
                        " Sum(CHIPS_Util_Logs_count) AS CHIPS_Util_Logs_count," +
                        " Sum(CHIPS_Util_Logs_TPA) AS CHIPS_Util_Logs_TPA," +
                        " Sum(CHIPS_Util_Logs_merch_wt) AS CHIPS_Util_Logs_merch_wt," +
                        " Sum(CHIPS_Util_Logs_merch_vol) AS CHIPS_Util_Logs_merch_vol," +
                        " Sum(CHIPS_Util_Chips_count) AS CHIPS_Util_Chips_count," +
                        " Sum(CHIPS_Util_Chips_TPA) AS CHIPS_Util_Chips_TPA," +
                        " Sum(CHIPS_Util_Chips_merch_wt) AS CHIPS_Util_Chips_merch_wt," +
                        " Sum(CHIPS_Util_Chips_merch_vol) AS CHIPS_Util_Chips_merch_vol," +
                        " Sum(CHIPS_Util_Logs_biomass_wt) AS CHIPS_Util_Logs_biomass_wt," +
                        " Sum(CHIPS_Util_Logs_biomass_vol) AS CHIPS_Util_Logs_biomass_vol," +
                        " Sum(CHIPS_Util_Chips_biomass_wt) AS CHIPS_Util_Chips_biomass_wt," +
                        " Sum(CHIPS_Util_Chips_biomass_vol) AS CHIPS_Util_Chips_biomass_vol," +
                		" Sum(CHIPS_NonUtil_Logs_count) AS CHIPS_NonUtil_Logs_count," +
                        " Sum(CHIPS_NonUtil_Logs_TPA) AS CHIPS_NonUtil_Logs_TPA," +
                        " Sum(CHIPS_NonUtil_Logs_merch_wt) AS CHIPS_NonUtil_Logs_merch_wt," +
                        " Sum(CHIPS_NonUtil_Logs_merch_vol) AS CHIPS_NonUtil_Logs_merch_vol," +
                        " Sum(CHIPS_NonUtil_Chips_count) AS CHIPS_NonUtil_Chips_count," +
                        " Sum(CHIPS_NonUtil_Chips_TPA) AS CHIPS_NonUtil_Chips_TPA," +
                        " Sum(CHIPS_NonUtil_Chips_merch_wt) AS CHIPS_NonUtil_Chips_merch_wt," +
                        " Sum(CHIPS_NonUtil_Chips_merch_vol) AS CHIPS_NonUtil_Chips_merch_vol," +
                        " Sum(CHIPS_NonUtil_Logs_biomass_wt) AS CHIPS_NonUtil_Logs_biomass_wt," +
                        " Sum(CHIPS_NonUtil_Logs_biomass_vol) AS CHIPS_NonUtil_Logs_biomass_vol," +
                        " Sum(CHIPS_NonUtil_Chips_biomass_wt) AS CHIPS_NonUtil_Chips_biomass_wt," +
                        " Sum(CHIPS_NonUtil_Chips_biomass_vol) AS CHIPS_NonUtil_Chips_biomass_vol," +
                        " Sum(SMLOGS_Util_Logs_count) AS SMLOGS_Util_Logs_count," +
                        " Sum(SMLOGS_Util_Logs_TPA) AS SMLOGS_Util_Logs_TPA," +
                        " Sum(SMLOGS_Util_Logs_merch_wt) AS SMLOGS_Util_Logs_merch_wt," +
                        " Sum(SMLOGS_Util_Logs_merch_vol) AS SMLOGS_Util_Logs_merch_vol," +
                        " Sum(SMLOGS_Util_Logs_biomass_wt) AS SMLOGS_Util_Logs_biomass_wt," +
                        " Sum(SMLOGS_Util_Logs_biomass_vol) AS SMLOGS_Util_Logs_biomass_vol," +
                        " Sum(SMLOGS_Util_Chips_count) AS SMLOGS_Util_Chips_count," +
                        " Sum(SMLOGS_Util_Chips_TPA) AS SMLOGS_Util_Chips_TPA," +
                        " Sum(SMLOGS_Util_Chips_merch_wt) AS SMLOGS_Util_Chips_merch_wt," +
                        " Sum(SMLOGS_Util_Chips_merch_vol) AS SMLOGS_Util_Chips_merch_vol," +
                        " Sum(SMLOGS_Util_Chips_biomass_wt) AS SMLOGS_Util_Chips_biomass_wt," +
                        " Sum(SMLOGS_Util_Chips_biomass_vol) AS SMLOGS_Util_Chips_biomass_vol," +
                        " Sum(SMLOGS_NonUtil_Logs_count) AS SMLOGS_NonUtil_Logs_count," +
                        " Sum(SMLOGS_NonUtil_Logs_TPA) AS SMLOGS_NonUtil_Logs_TPA," +
                        " Sum(SMLOGS_NonUtil_Logs_merch_wt) AS SMLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(SMLOGS_NonUtil_Logs_merch_vol) AS SMLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(SMLOGS_NonUtil_Logs_biomass_wt) AS SMLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(SMLOGS_NonUtil_Logs_biomass_vol) AS SMLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(SMLOGS_NonUtil_Chips_count) AS SMLOGS_NonUtil_Chips_count," +
                        " Sum(SMLOGS_NonUtil_Chips_TPA) AS SMLOGS_NonUtil_Chips_TPA," +
                        " Sum(SMLOGS_NonUtil_Chips_merch_wt) AS SMLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(SMLOGS_NonUtil_Chips_merch_vol) AS SMLOGS_NonUtil_Chips_merch_vol," +
                        " Sum(SMLOGS_NonUtil_Chips_biomass_wt) AS SMLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(SMLOGS_NonUtil_Chips_biomass_vol) AS SMLOGS_NonUtil_Chips_biomass_vol," +
                        " Sum(LGLOGS_Util_Logs_count) AS LGLOGS_Util_Logs_count," +
                        " Sum(LGLOGS_Util_Logs_TPA) AS LGLOGS_Util_Logs_TPA," +
                        " Sum(LGLOGS_Util_Logs_merch_wt) AS LGLOGS_Util_Logs_merch_wt," +
                        " Sum(LGLOGS_Util_Logs_merch_vol) AS LGLOGS_Util_Logs_merch_vol," +
                        " Sum(LGLOGS_Util_Logs_biomass_wt) AS LGLOGS_Util_Logs_biomass_wt," +
                        " Sum(LGLOGS_Util_Logs_biomass_vol) AS LGLOGS_Util_Logs_biomass_vol," +
                        " Sum(LGLOGS_Util_Chips_count) AS LGLOGS_Util_Chips_count," +
                        " Sum(LGLOGS_Util_Chips_TPA) AS LGLOGS_Util_Chips_TPA," +
                        " Sum(LGLOGS_Util_Chips_merch_wt) AS LGLOGS_Util_Chips_merch_wt," +
                        " Sum(LGLOGS_Util_Chips_merch_vol) AS LGLOGS_Util_Chips_merch_vol, " +
                        " Sum(LGLOGS_Util_Chips_biomass_wt) AS LGLOGS_Util_Chips_biomass_wt," +
                        " Sum(LGLOGS_Util_Chips_biomass_vol) AS LGLOGS_Util_Chips_biomass_vol," +
                     	" Sum(LGLOGS_NonUtil_Logs_count) AS LGLOGS_NonUtil_Logs_count," +
                        " Sum(LGLOGS_NonUtil_Logs_TPA) AS LGLOGS_NonUtil_Logs_TPA," +
                        " Sum(LGLOGS_NonUtil_Logs_merch_wt) AS LGLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(LGLOGS_NonUtil_Logs_merch_vol) AS LGLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(LGLOGS_NonUtil_Logs_biomass_wt) AS LGLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(LGLOGS_NonUtil_Logs_biomass_vol) AS LGLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(LGLOGS_NonUtil_Chips_count) AS LGLOGS_NonUtil_Chips_count," +
                        " Sum(LGLOGS_NonUtil_Chips_TPA) AS LGLOGS_NonUtil_Chips_TPA," +
                        " Sum(LGLOGS_NonUtil_Chips_merch_wt) AS LGLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(LGLOGS_NonUtil_Chips_merch_vol) AS LGLOGS_NonUtil_Chips_merch_vol, " +
                        " Sum(LGLOGS_NonUtil_Chips_biomass_wt) AS LGLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(LGLOGS_NonUtil_Chips_biomass_vol) AS LGLOGS_NonUtil_Chips_biomass_vol " +
                        " INTO " + p_strIntoTableName + " " +
                        " FROM " + p_strFromTableName + " " +
                        " GROUP BY BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE, species_group, diam_group" +
                        " HAVING (((species_group) Is Not Null))";
            }
            public static string SumHwdBinByPlotRxSpcGrpDbhGrp(string p_strFromTableName, string p_strIntoTableName)
            {
                return "SELECT BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE, species_group, diam_group," +
                        " Sum(HWD_BC_Util_Logs_count) AS HWD_BC_Util_Logs_count," +
                        " Sum(HWD_BC_Util_Logs_TPA) AS HWD_BC_Util_Logs_TPA," +
                        " Sum(HWD_BC_Util_Logs_merch_wt) AS HWD_BC_Util_Logs_merch_wt," +
                        " Sum(HWD_BC_Util_Logs_merch_vol) AS HWD_BC_Util_Logs_merch_vol," +
                        " Sum(HWD_BC_Util_Logs_biomass_vol) AS HWD_BC_Util_Logs_biomass_vol," +
                        " Sum(HWD_BC_Util_Logs_biomass_wt) AS HWD_BC_Util_Logs_biomass_wt," +
                        " Sum(HWD_BC_Util_Chips_count) AS HWD_BC_Util_Chips_count," +
                        " Sum(HWD_BC_Util_Chips_TPA) AS HWD_BC_Util_Chips_TPA," +
                        " Sum(HWD_BC_Util_Chips_merch_wt) AS HWD_BC_Util_Chips_merch_wt," +
                        " Sum(HWD_BC_Util_Chips_merch_vol) AS HWD_BC_Util_Chips_merch_vol," +
                        " Sum(HWD_BC_Util_Chips_biomass_wt) AS HWD_BC_Util_Chips_biomass_wt," +
                        " Sum(HWD_BC_Util_Chips_biomass_vol) AS HWD_BC_Util_Chips_biomass_vol," +
                        " Sum(HWD_BC_NonUtil_Logs_count) AS HWD_BC_NonUtil_Logs_count," +
                        " Sum(HWD_BC_NonUtil_Logs_TPA) AS HWD_BC_NonUtil_Logs_TPA," +
                        " Sum(HWD_BC_NonUtil_Logs_merch_wt) AS HWD_BC_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_BC_NonUtil_Logs_merch_vol) AS HWD_BC_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_BC_NonUtil_Logs_biomass_vol) AS HWD_BC_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_BC_NonUtil_Logs_biomass_wt) AS HWD_BC_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_BC_NonUtil_Chips_count) AS HWD_BC_NonUtil_Chips_count," +
                        " Sum(HWD_BC_NonUtil_Chips_TPA) AS HWD_BC_NonUtil_Chips_TPA," +
                        " Sum(HWD_BC_NonUtil_Chips_merch_wt) AS HWD_BC_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_BC_NonUtil_Chips_merch_vol) AS HWD_BC_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_BC_NonUtil_Chips_biomass_wt) AS HWD_BC_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_BC_NonUtil_Chips_biomass_vol) AS HWD_BC_NonUtil_Chips_biomass_vol," +
                        " Sum(HWD_CHIPS_Util_Logs_count) AS HWD_CHIPS_Util_Logs_count," +
                        " Sum(HWD_CHIPS_Util_Logs_TPA) AS HWD_CHIPS_Util_Logs_TPA," +
                        " Sum(HWD_CHIPS_Util_Logs_merch_wt) AS HWD_CHIPS_Util_Logs_merch_wt," +
                        " Sum(HWD_CHIPS_Util_Logs_merch_vol) AS HWD_CHIPS_Util_Logs_merch_vol," +
                        " Sum(HWD_CHIPS_Util_Chips_count) AS HWD_CHIPS_Util_Chips_count," +
                        " Sum(HWD_CHIPS_Util_Chips_TPA) AS HWD_CHIPS_Util_Chips_TPA," +
                        " Sum(HWD_CHIPS_Util_Chips_merch_wt) AS HWD_CHIPS_Util_Chips_merch_wt," +
                        " Sum(HWD_CHIPS_Util_Chips_merch_vol) AS HWD_CHIPS_Util_Chips_merch_vol," +
                        " Sum(HWD_CHIPS_Util_Logs_biomass_wt) AS HWD_CHIPS_Util_Logs_biomass_wt," +
                        " Sum(HWD_CHIPS_Util_Logs_biomass_vol) AS HWD_CHIPS_Util_Logs_biomass_vol," +
                        " Sum(HWD_CHIPS_Util_Chips_biomass_wt) AS HWD_CHIPS_Util_Chips_biomass_wt," +
                        " Sum(HWD_CHIPS_Util_Chips_biomass_vol) AS HWD_CHIPS_Util_Chips_biomass_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_count) AS HWD_CHIPS_NonUtil_Logs_count," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_TPA) AS HWD_CHIPS_NonUtil_Logs_TPA," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_merch_wt) AS HWD_CHIPS_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_merch_vol) AS HWD_CHIPS_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_count) AS HWD_CHIPS_NonUtil_Chips_count," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_TPA) AS HWD_CHIPS_NonUtil_Chips_TPA," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_merch_wt) AS HWD_CHIPS_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_merch_vol) AS HWD_CHIPS_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_biomass_wt) AS HWD_CHIPS_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_biomass_vol) AS HWD_CHIPS_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_biomass_wt) AS HWD_CHIPS_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_biomass_vol) AS HWD_CHIPS_NonUtil_Chips_biomass_vol," +
                        " Sum(HWD_SMLOGS_Util_Logs_count) AS HWD_SMLOGS_Util_Logs_count," +
                        " Sum(HWD_SMLOGS_Util_Logs_TPA) AS HWD_SMLOGS_Util_Logs_TPA," +
                        " Sum(HWD_SMLOGS_Util_Logs_merch_wt) AS HWD_SMLOGS_Util_Logs_merch_wt," +
                        " Sum(HWD_SMLOGS_Util_Logs_merch_vol) AS HWD_SMLOGS_Util_Logs_merch_vol," +
                        " Sum(HWD_SMLOGS_Util_Logs_biomass_wt) AS HWD_SMLOGS_Util_Logs_biomass_wt," +
                        " Sum(HWD_SMLOGS_Util_Logs_biomass_vol) AS HWD_SMLOGS_Util_Logs_biomass_vol," +
                        " Sum(HWD_SMLOGS_Util_Chips_count) AS HWD_SMLOGS_Util_Chips_count," +
                        " Sum(HWD_SMLOGS_Util_Chips_TPA) AS HWD_SMLOGS_Util_Chips_TPA," +
                        " Sum(HWD_SMLOGS_Util_Chips_merch_wt) AS HWD_SMLOGS_Util_Chips_merch_wt," +
                        " Sum(HWD_SMLOGS_Util_Chips_merch_vol) AS HWD_SMLOGS_Util_Chips_merch_vol," +
                        " Sum(HWD_SMLOGS_Util_Chips_biomass_wt) AS HWD_SMLOGS_Util_Chips_biomass_wt," +
                        " Sum(HWD_SMLOGS_Util_Chips_biomass_vol) AS HWD_SMLOGS_Util_Chips_biomass_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_count) AS HWD_SMLOGS_NonUtil_Logs_count," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_TPA) AS HWD_SMLOGS_NonUtil_Logs_TPA," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_merch_wt) AS HWD_SMLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_merch_vol) AS HWD_SMLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_biomass_wt) AS HWD_SMLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_biomass_vol) AS HWD_SMLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_count) AS HWD_SMLOGS_NonUtil_Chips_count," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_TPA) AS HWD_SMLOGS_NonUtil_Chips_TPA," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_merch_wt) AS HWD_SMLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_merch_vol) AS HWD_SMLOGS_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_biomass_wt) AS HWD_SMLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_biomass_vol) AS HWD_SMLOGS_NonUtil_Chips_biomass_vol," +
                        " Sum(HWD_LGLOGS_Util_Logs_count) AS HWD_LGLOGS_Util_Logs_count," +
                        " Sum(HWD_LGLOGS_Util_Logs_TPA) AS HWD_LGLOGS_Util_Logs_TPA," +
                        " Sum(HWD_LGLOGS_Util_Logs_merch_wt) AS HWD_LGLOGS_Util_Logs_merch_wt," +
                        " Sum(HWD_LGLOGS_Util_Logs_merch_vol) AS HWD_LGLOGS_Util_Logs_merch_vol," +
                        " Sum(HWD_LGLOGS_Util_Logs_biomass_wt) AS HWD_LGLOGS_Util_Logs_biomass_wt," +
                        " Sum(HWD_LGLOGS_Util_Logs_biomass_vol) AS HWD_LGLOGS_Util_Logs_biomass_vol," +
                        " Sum(HWD_LGLOGS_Util_Chips_count) AS HWD_LGLOGS_Util_Chips_count," +
                        " Sum(HWD_LGLOGS_Util_Chips_TPA) AS HWD_LGLOGS_Util_Chips_TPA," +
                        " Sum(HWD_LGLOGS_Util_Chips_merch_wt) AS HWD_LGLOGS_Util_Chips_merch_wt," +
                        " Sum(HWD_LGLOGS_Util_Chips_merch_vol) AS HWD_LGLOGS_Util_Chips_merch_vol," +
                        " Sum(HWD_LGLOGS_Util_Chips_biomass_wt) AS HWD_LGLOGS_Util_Chips_biomass_wt," +
                        " Sum(HWD_LGLOGS_Util_Chips_biomass_vol) AS HWD_LGLOGS_Util_Chips_biomass_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_count) AS HWD_LGLOGS_NonUtil_Logs_count," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_TPA) AS HWD_LGLOGS_NonUtil_Logs_TPA," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_merch_wt) AS HWD_LGLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_merch_vol) AS HWD_LGLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_biomass_wt) AS HWD_LGLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_biomass_vol) AS HWD_LGLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_count) AS HWD_LGLOGS_NonUtil_Chips_count," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_TPA) AS HWD_LGLOGS_NonUtil_Chips_TPA," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_merch_wt) AS HWD_LGLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_merch_vol) AS HWD_LGLOGS_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_biomass_wt) AS HWD_LGLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_biomass_vol) AS HWD_LGLOGS_NonUtil_Chips_biomass_vol " +
                        " INTO " + p_strIntoTableName + " " +
                        " FROM " + p_strFromTableName + " " +
                        " GROUP BY BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE, species_group, diam_group" +
                        " HAVING (((species_group) Is Not Null))";
            }
            public static string SumBinByPlotRx(string p_strFromTableName, string p_strIntoTableName)
            {
                return "SELECT BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE," +
                        " Sum(BC_Util_Logs_count) AS BC_Util_Logs_count," +
                        " Sum(BC_Util_Logs_TPA) AS BC_Util_Logs_TPA," +
                        " Sum(BC_Util_Logs_merch_wt) AS BC_Util_Logs_merch_wt," +
                        " Sum(BC_Util_Logs_merch_vol) AS BC_Util_Logs_merch_vol," +
                        " Sum(BC_Util_Logs_biomass_vol) AS BC_Util_Logs_biomass_vol," +
                        " Sum(BC_Util_Logs_biomass_wt) AS BC_Util_Logs_biomass_wt," +
                        " Sum(BC_Util_Chips_count) AS BC_Util_Chips_count," +
                        " Sum(BC_Util_Chips_TPA) AS BC_Util_Chips_TPA," +
                        " Sum(BC_Util_Chips_merch_wt) AS BC_Util_Chips_merch_wt," +
                        " Sum(BC_Util_Chips_merch_vol) AS BC_Util_Chips_merch_vol," +
                        " Sum(BC_Util_Chips_biomass_wt) AS BC_Util_Chips_biomass_wt," +
                        " Sum(BC_Util_Chips_biomass_vol) AS BC_Util_Chips_biomass_vol," +
                        " Sum(BC_NonUtil_Logs_count) AS BC_NonUtil_Logs_count," +
                        " Sum(BC_NonUtil_Logs_TPA) AS BC_NonUtil_Logs_TPA," +
                        " Sum(BC_NonUtil_Logs_merch_wt) AS BC_NonUtil_Logs_merch_wt," +
                        " Sum(BC_NonUtil_Logs_merch_vol) AS BC_NonUtil_Logs_merch_vol," +
                        " Sum(BC_NonUtil_Logs_biomass_vol) AS BC_NonUtil_Logs_biomass_vol," +
                        " Sum(BC_NonUtil_Logs_biomass_wt) AS BC_NonUtil_Logs_biomass_wt," +
                        " Sum(BC_NonUtil_Chips_count) AS BC_NonUtil_Chips_count," +
                        " Sum(BC_NonUtil_Chips_TPA) AS BC_NonUtil_Chips_TPA," +
                        " Sum(BC_NonUtil_Chips_merch_wt) AS BC_NonUtil_Chips_merch_wt," +
                        " Sum(BC_NonUtil_Chips_merch_vol) AS BC_NonUtil_Chips_merch_vol," +
                        " Sum(BC_NonUtil_Chips_biomass_wt) AS BC_NonUtil_Chips_biomass_wt," +
                        " Sum(BC_NonUtil_Chips_biomass_vol) AS BC_NonUtil_Chips_biomass_vol," +
                        " Sum(CHIPS_Util_Logs_count) AS CHIPS_Util_Logs_count," +
                        " Sum(CHIPS_Util_Logs_TPA) AS CHIPS_Util_Logs_TPA," +
                        " Sum(CHIPS_Util_Logs_merch_wt) AS CHIPS_Util_Logs_merch_wt," +
                        " Sum(CHIPS_Util_Logs_merch_vol) AS CHIPS_Util_Logs_merch_vol," +
                        " Sum(CHIPS_Util_Chips_count) AS CHIPS_Util_Chips_count," +
                        " Sum(CHIPS_Util_Chips_TPA) AS CHIPS_Util_Chips_TPA," +
                        " Sum(CHIPS_Util_Chips_merch_wt) AS CHIPS_Util_Chips_merch_wt," +
                        " Sum(CHIPS_Util_Chips_merch_vol) AS CHIPS_Util_Chips_merch_vol," +
                        " Sum(CHIPS_Util_Logs_biomass_wt) AS CHIPS_Util_Logs_biomass_wt," +
                        " Sum(CHIPS_Util_Logs_biomass_vol) AS CHIPS_Util_Logs_biomass_vol," +
                        " Sum(CHIPS_Util_Chips_biomass_wt) AS CHIPS_Util_Chips_biomass_wt," +
                        " Sum(CHIPS_Util_Chips_biomass_vol) AS CHIPS_Util_Chips_biomass_vol," +
                        " Sum(CHIPS_NonUtil_Logs_count) AS CHIPS_NonUtil_Logs_count," +
                        " Sum(CHIPS_NonUtil_Logs_TPA) AS CHIPS_NonUtil_Logs_TPA," +
                        " Sum(CHIPS_NonUtil_Logs_merch_wt) AS CHIPS_NonUtil_Logs_merch_wt," +
                        " Sum(CHIPS_NonUtil_Logs_merch_vol) AS CHIPS_NonUtil_Logs_merch_vol," +
                        " Sum(CHIPS_NonUtil_Chips_count) AS CHIPS_NonUtil_Chips_count," +
                        " Sum(CHIPS_NonUtil_Chips_TPA) AS CHIPS_NonUtil_Chips_TPA," +
                        " Sum(CHIPS_NonUtil_Chips_merch_wt) AS CHIPS_NonUtil_Chips_merch_wt," +
                        " Sum(CHIPS_NonUtil_Chips_merch_vol) AS CHIPS_NonUtil_Chips_merch_vol," +
                        " Sum(CHIPS_NonUtil_Logs_biomass_wt) AS CHIPS_NonUtil_Logs_biomass_wt," +
                        " Sum(CHIPS_NonUtil_Logs_biomass_vol) AS CHIPS_NonUtil_Logs_biomass_vol," +
                        " Sum(CHIPS_NonUtil_Chips_biomass_wt) AS CHIPS_NonUtil_Chips_biomass_wt," +
                        " Sum(CHIPS_NonUtil_Chips_biomass_vol) AS CHIPS_NonUtil_Chips_biomass_vol," +
                        " Sum(SMLOGS_Util_Logs_count) AS SMLOGS_Util_Logs_count," +
                        " Sum(SMLOGS_Util_Logs_TPA) AS SMLOGS_Util_Logs_TPA," +
                        " Sum(SMLOGS_Util_Logs_merch_wt) AS SMLOGS_Util_Logs_merch_wt," +
                        " Sum(SMLOGS_Util_Logs_merch_vol) AS SMLOGS_Util_Logs_merch_vol," +
                        " Sum(SMLOGS_Util_Logs_biomass_wt) AS SMLOGS_Util_Logs_biomass_wt," +
                        " Sum(SMLOGS_Util_Logs_biomass_vol) AS SMLOGS_Util_Logs_biomass_vol," +
                        " Sum(SMLOGS_Util_Chips_count) AS SMLOGS_Util_Chips_count," +
                        " Sum(SMLOGS_Util_Chips_TPA) AS SMLOGS_Util_Chips_TPA," +
                        " Sum(SMLOGS_Util_Chips_merch_wt) AS SMLOGS_Util_Chips_merch_wt," +
                        " Sum(SMLOGS_Util_Chips_merch_vol) AS SMLOGS_Util_Chips_merch_vol," +
                        " Sum(SMLOGS_Util_Chips_biomass_wt) AS SMLOGS_Util_Chips_biomass_wt," +
                        " Sum(SMLOGS_Util_Chips_biomass_vol) AS SMLOGS_Util_Chips_biomass_vol," +
						" Sum(SMLOGS_NonUtil_Logs_count) AS SMLOGS_NonUtil_Logs_count," +
                        " Sum(SMLOGS_NonUtil_Logs_TPA) AS SMLOGS_NonUtil_Logs_TPA," +
                        " Sum(SMLOGS_NonUtil_Logs_merch_wt) AS SMLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(SMLOGS_NonUtil_Logs_merch_vol) AS SMLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(SMLOGS_NonUtil_Logs_biomass_wt) AS SMLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(SMLOGS_NonUtil_Logs_biomass_vol) AS SMLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(SMLOGS_NonUtil_Chips_count) AS SMLOGS_NonUtil_Chips_count," +
                        " Sum(SMLOGS_NonUtil_Chips_TPA) AS SMLOGS_NonUtil_Chips_TPA," +
                        " Sum(SMLOGS_NonUtil_Chips_merch_wt) AS SMLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(SMLOGS_NonUtil_Chips_merch_vol) AS SMLOGS_NonUtil_Chips_merch_vol," +
                        " Sum(SMLOGS_NonUtil_Chips_biomass_wt) AS SMLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(SMLOGS_NonUtil_Chips_biomass_vol) AS SMLOGS_NonUtil_Chips_biomass_vol," +
                        " Sum(LGLOGS_Util_Logs_count) AS LGLOGS_Util_Logs_count," +
                        " Sum(LGLOGS_Util_Logs_TPA) AS LGLOGS_Util_Logs_TPA," +
                        " Sum(LGLOGS_Util_Logs_merch_wt) AS LGLOGS_Util_Logs_merch_wt," +
                        " Sum(LGLOGS_Util_Logs_merch_vol) AS LGLOGS_Util_Logs_merch_vol," +
                        " Sum(LGLOGS_Util_Logs_biomass_wt) AS LGLOGS_Util_Logs_biomass_wt," +
                        " Sum(LGLOGS_Util_Logs_biomass_vol) AS LGLOGS_Util_Logs_biomass_vol," +
                        " Sum(LGLOGS_Util_Chips_count) AS LGLOGS_Util_Chips_count," +
                        " Sum(LGLOGS_Util_Chips_TPA) AS LGLOGS_Util_Chips_TPA," +
                        " Sum(LGLOGS_Util_Chips_merch_wt) AS LGLOGS_Util_Chips_merch_wt," +
                        " Sum(LGLOGS_Util_Chips_merch_vol) AS LGLOGS_Util_Chips_merch_vol, " +
                        " Sum(LGLOGS_Util_Chips_biomass_wt) AS LGLOGS_Util_Chips_biomass_wt," +
                        " Sum(LGLOGS_Util_Chips_biomass_vol) AS LGLOGS_Util_Chips_biomass_vol," +
                        " Sum(LGLOGS_NonUtil_Logs_count) AS LGLOGS_NonUtil_Logs_count," +
                        " Sum(LGLOGS_NonUtil_Logs_TPA) AS LGLOGS_NonUtil_Logs_TPA," +
                        " Sum(LGLOGS_NonUtil_Logs_merch_wt) AS LGLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(LGLOGS_NonUtil_Logs_merch_vol) AS LGLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(LGLOGS_NonUtil_Logs_biomass_wt) AS LGLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(LGLOGS_NonUtil_Logs_biomass_vol) AS LGLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(LGLOGS_NonUtil_Chips_count) AS LGLOGS_NonUtil_Chips_count," +
                        " Sum(LGLOGS_NonUtil_Chips_TPA) AS LGLOGS_NonUtil_Chips_TPA," +
                        " Sum(LGLOGS_NonUtil_Chips_merch_wt) AS LGLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(LGLOGS_NonUtil_Chips_merch_vol) AS LGLOGS_NonUtil_Chips_merch_vol, " +
                        " Sum(LGLOGS_NonUtil_Chips_biomass_wt) AS LGLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(LGLOGS_NonUtil_Chips_biomass_vol) AS LGLOGS_NonUtil_Chips_biomass_vol " +
                        " INTO " + p_strIntoTableName + " " +
                        " FROM " + p_strFromTableName + " " +
                        " GROUP BY BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE";
                       
            }
            public static string SumHwdBinByPlotRx(string p_strFromTableName, string p_strIntoTableName)
            {
                return "SELECT BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE," +
                        " Sum(HWD_BC_Util_Logs_count) AS HWD_BC_Util_Logs_count," +
                        " Sum(HWD_BC_Util_Logs_TPA) AS HWD_BC_Util_Logs_TPA," +
                        " Sum(HWD_BC_Util_Logs_merch_wt) AS HWD_BC_Util_Logs_merch_wt," +
                        " Sum(HWD_BC_Util_Logs_merch_vol) AS HWD_BC_Util_Logs_merch_vol," +
                        " Sum(HWD_BC_Util_Logs_biomass_vol) AS HWD_BC_Util_Logs_biomass_vol," +
                        " Sum(HWD_BC_Util_Logs_biomass_wt) AS HWD_BC_Util_Logs_biomass_wt," +
                        " Sum(HWD_BC_Util_Chips_count) AS HWD_BC_Util_Chips_count," +
                        " Sum(HWD_BC_Util_Chips_TPA) AS HWD_BC_Util_Chips_TPA," +
                        " Sum(HWD_BC_Util_Chips_merch_wt) AS HWD_BC_Util_Chips_merch_wt," +
                        " Sum(HWD_BC_Util_Chips_merch_vol) AS HWD_BC_Util_Chips_merch_vol," +
                        " Sum(HWD_BC_Util_Chips_biomass_wt) AS HWD_BC_Util_Chips_biomass_wt," +
                        " Sum(HWD_BC_Util_Chips_biomass_vol) AS HWD_BC_Util_Chips_biomass_vol," +
                        " Sum(HWD_BC_NonUtil_Logs_count) AS HWD_BC_NonUtil_Logs_count," +
                        " Sum(HWD_BC_NonUtil_Logs_TPA) AS HWD_BC_NonUtil_Logs_TPA," +
                        " Sum(HWD_BC_NonUtil_Logs_merch_wt) AS HWD_BC_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_BC_NonUtil_Logs_merch_vol) AS HWD_BC_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_BC_NonUtil_Logs_biomass_vol) AS HWD_BC_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_BC_NonUtil_Logs_biomass_wt) AS HWD_BC_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_BC_NonUtil_Chips_count) AS HWD_BC_NonUtil_Chips_count," +
                        " Sum(HWD_BC_NonUtil_Chips_TPA) AS HWD_BC_NonUtil_Chips_TPA," +
                        " Sum(HWD_BC_NonUtil_Chips_merch_wt) AS HWD_BC_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_BC_NonUtil_Chips_merch_vol) AS HWD_BC_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_BC_NonUtil_Chips_biomass_wt) AS HWD_BC_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_BC_NonUtil_Chips_biomass_vol) AS HWD_BC_NonUtil_Chips_biomass_vol," +
                        " Sum(HWD_CHIPS_Util_Logs_count) AS HWD_CHIPS_Util_Logs_count," +
                        " Sum(HWD_CHIPS_Util_Logs_TPA) AS HWD_CHIPS_Util_Logs_TPA," +
                        " Sum(HWD_CHIPS_Util_Logs_merch_wt) AS HWD_CHIPS_Util_Logs_merch_wt," +
                        " Sum(HWD_CHIPS_Util_Logs_merch_vol) AS HWD_CHIPS_Util_Logs_merch_vol," +
                        " Sum(HWD_CHIPS_Util_Chips_count) AS HWD_CHIPS_Util_Chips_count," +
                        " Sum(HWD_CHIPS_Util_Chips_TPA) AS HWD_CHIPS_Util_Chips_TPA," +
                        " Sum(HWD_CHIPS_Util_Chips_merch_wt) AS HWD_CHIPS_Util_Chips_merch_wt," +
                        " Sum(HWD_CHIPS_Util_Chips_merch_vol) AS HWD_CHIPS_Util_Chips_merch_vol," +
                        " Sum(HWD_CHIPS_Util_Logs_biomass_wt) AS HWD_CHIPS_Util_Logs_biomass_wt," +
                        " Sum(HWD_CHIPS_Util_Logs_biomass_vol) AS HWD_CHIPS_Util_Logs_biomass_vol," +
                        " Sum(HWD_CHIPS_Util_Chips_biomass_wt) AS HWD_CHIPS_Util_Chips_biomass_wt," +
                        " Sum(HWD_CHIPS_Util_Chips_biomass_vol) AS HWD_CHIPS_Util_Chips_biomass_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_count) AS HWD_CHIPS_NonUtil_Logs_count," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_TPA) AS HWD_CHIPS_NonUtil_Logs_TPA," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_merch_wt) AS HWD_CHIPS_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_merch_vol) AS HWD_CHIPS_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_count) AS HWD_CHIPS_NonUtil_Chips_count," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_TPA) AS HWD_CHIPS_NonUtil_Chips_TPA," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_merch_wt) AS HWD_CHIPS_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_merch_vol) AS HWD_CHIPS_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_biomass_wt) AS HWD_CHIPS_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Logs_biomass_vol) AS HWD_CHIPS_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_biomass_wt) AS HWD_CHIPS_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_CHIPS_NonUtil_Chips_biomass_vol) AS HWD_CHIPS_NonUtil_Chips_biomass_vol," +
                        " Sum(HWD_SMLOGS_Util_Logs_count) AS HWD_SMLOGS_Util_Logs_count," +
                        " Sum(HWD_SMLOGS_Util_Logs_TPA) AS HWD_SMLOGS_Util_Logs_TPA," +
                        " Sum(HWD_SMLOGS_Util_Logs_merch_wt) AS HWD_SMLOGS_Util_Logs_merch_wt," +
                        " Sum(HWD_SMLOGS_Util_Logs_merch_vol) AS HWD_SMLOGS_Util_Logs_merch_vol," +
                        " Sum(HWD_SMLOGS_Util_Logs_biomass_wt) AS HWD_SMLOGS_Util_Logs_biomass_wt," +
                        " Sum(HWD_SMLOGS_Util_Logs_biomass_vol) AS HWD_SMLOGS_Util_Logs_biomass_vol," +
                        " Sum(HWD_SMLOGS_Util_Chips_count) AS HWD_SMLOGS_Util_Chips_count," +
                        " Sum(HWD_SMLOGS_Util_Chips_TPA) AS HWD_SMLOGS_Util_Chips_TPA," +
                        " Sum(HWD_SMLOGS_Util_Chips_merch_wt) AS HWD_SMLOGS_Util_Chips_merch_wt," +
                        " Sum(HWD_SMLOGS_Util_Chips_merch_vol) AS HWD_SMLOGS_Util_Chips_merch_vol," +
                        " Sum(HWD_SMLOGS_Util_Chips_biomass_wt) AS HWD_SMLOGS_Util_Chips_biomass_wt," +
                        " Sum(HWD_SMLOGS_Util_Chips_biomass_vol) AS HWD_SMLOGS_Util_Chips_biomass_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_count) AS HWD_SMLOGS_NonUtil_Logs_count," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_TPA) AS HWD_SMLOGS_NonUtil_Logs_TPA," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_merch_wt) AS HWD_SMLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_merch_vol) AS HWD_SMLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_biomass_wt) AS HWD_SMLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Logs_biomass_vol) AS HWD_SMLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_count) AS HWD_SMLOGS_NonUtil_Chips_count," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_TPA) AS HWD_SMLOGS_NonUtil_Chips_TPA," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_merch_wt) AS HWD_SMLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_merch_vol) AS HWD_SMLOGS_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_biomass_wt) AS HWD_SMLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_SMLOGS_NonUtil_Chips_biomass_vol) AS HWD_SMLOGS_NonUtil_Chips_biomass_vol," +
                        " Sum(HWD_LGLOGS_Util_Logs_count) AS HWD_LGLOGS_Util_Logs_count," +
                        " Sum(HWD_LGLOGS_Util_Logs_TPA) AS HWD_LGLOGS_Util_Logs_TPA," +
                        " Sum(HWD_LGLOGS_Util_Logs_merch_wt) AS HWD_LGLOGS_Util_Logs_merch_wt," +
                        " Sum(HWD_LGLOGS_Util_Logs_merch_vol) AS HWD_LGLOGS_Util_Logs_merch_vol," +
                        " Sum(HWD_LGLOGS_Util_Logs_biomass_wt) AS HWD_LGLOGS_Util_Logs_biomass_wt," +
                        " Sum(HWD_LGLOGS_Util_Logs_biomass_vol) AS HWD_LGLOGS_Util_Logs_biomass_vol," +
                        " Sum(HWD_LGLOGS_Util_Chips_count) AS HWD_LGLOGS_Util_Chips_count," +
                        " Sum(HWD_LGLOGS_Util_Chips_TPA) AS HWD_LGLOGS_Util_Chips_TPA," +
                        " Sum(HWD_LGLOGS_Util_Chips_merch_wt) AS HWD_LGLOGS_Util_Chips_merch_wt," +
                        " Sum(HWD_LGLOGS_Util_Chips_merch_vol) AS HWD_LGLOGS_Util_Chips_merch_vol," +
                        " Sum(HWD_LGLOGS_Util_Chips_biomass_wt) AS HWD_LGLOGS_Util_Chips_biomass_wt," +
                        " Sum(HWD_LGLOGS_Util_Chips_biomass_vol) AS HWD_LGLOGS_Util_Chips_biomass_vol, " +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_count) AS HWD_LGLOGS_NonUtil_Logs_count," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_TPA) AS HWD_LGLOGS_NonUtil_Logs_TPA," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_merch_wt) AS HWD_LGLOGS_NonUtil_Logs_merch_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_merch_vol) AS HWD_LGLOGS_NonUtil_Logs_merch_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_biomass_wt) AS HWD_LGLOGS_NonUtil_Logs_biomass_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Logs_biomass_vol) AS HWD_LGLOGS_NonUtil_Logs_biomass_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_count) AS HWD_LGLOGS_NonUtil_Chips_count," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_TPA) AS HWD_LGLOGS_NonUtil_Chips_TPA," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_merch_wt) AS HWD_LGLOGS_NonUtil_Chips_merch_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_merch_vol) AS HWD_LGLOGS_NonUtil_Chips_merch_vol," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_biomass_wt) AS HWD_LGLOGS_NonUtil_Chips_biomass_wt," +
                        " Sum(HWD_LGLOGS_NonUtil_Chips_biomass_vol) AS HWD_LGLOGS_NonUtil_Chips_biomass_vol " +
                        " INTO " + p_strIntoTableName + " " +
                        " FROM " + p_strFromTableName + " " +
                        " GROUP BY BIOSUM_COND_ID, RXPACKAGE,RX,RXCYCLE";
                       
            }
            public static string InitializeOutputTableValues(string p_strOutputTableName)
            {
               
                return
                    "UPDATE " + p_strOutputTableName + " o " +
                    "SET o.[CHIPS TPA] = 0," + 
                        "o.[CHIPS Average Vol (ft3)] = 0," + 
                        "o.[CHIPS Average Weight (tons)] = 0," +
                        "o.[CHIPS Average Density (lbs/ft3)] = 0," + 
                        "o.[CHIPS Hwd Proportion] = 0," + 
                        "o.[CHIPS Chip Fraction] = 0," +
                        "o.[CHIPS utilized logs (ft3)] = 0," +
                        "o.[CHIPS utilized chips (tons)] = 0," +
                        "o.[SMLOGS TPA] = 0," +
                        "o.[SMLOGS Average Vol (ft3)] = 0," +
                        "o.[SMLOGS Average Weight (tons)] = 0," +
                        "o.[SMLOGS Average Density (lbs/ft3)] = 0," +
                        "o.[SMLOGS Hwd Proportion] = 0," +
                        "o.[SMLOGS Chip Fraction] = 0," +
                        "o.[SMLOGS utilized logs (ft3)] = 0," +
                        "o.[SMLOGS utilized chips (tons)] = 0," +
                        "o.[LGLOGS TPA] = 0," +
                        "o.[LGLOGS Average Vol (ft3)] = 0," +
                        "o.[LGLOGS Average Weight (tons)] = 0," +
                        "o.[LGLOGS Average Density (lbs/ft3)] =0," +
                        "o.[LGLOGS Hwd Proportion] =0," +
                        "o.[LGLOGS Chip Fraction] = 0," +
                        "o.[LGLOGS utilized logs (ft3)] = 0," +
                        "o.[LGLOGS utilized chips (tons)] = 0," +
                        "o.[TOTAL TPA] = 0," +
                        "o.[TOTAL Average Vol (ft3)] = 0," +
                        "o.[TOTAL Average Weight (tons)] = 0," +
                        "o.[TOTAL Average Density (lbs/ft3)] = 0," +
                        "o.[TOTAL Hwd Proportion] = 0," +
                        "o.[TOTAL Chip Fraction] = 0," +
                        "o.[TOTAL utilized logs (ft3)] = 0," +
                        "o.[TOTAL utilized chips (tons)] = 0," +
                        "o.[BRUSH CUT utilized logs (ft3)] = 0," +
                        "o.[BRUSH CUT utilized chips (tons)] = 0," +
                        "o.[BRUSH CUT not utilized TPA] = 0," +
                        "o.[BRUSH CUT not utilized Average Vol (ft3)] = 0," +
                        "o.[BRUSH CUT not utilized Average Weight (tons)] = 0, " +
                        "o.[BRUSH CUT not utilized Average Density (lbs/ft3)] = 0, " +
                        "o.[BRUSH CUT not utilized Hwd Proportion] = 0," +
                        "o.[BRUSH CUT not utilized Chip Fraction] = 0," +
                        "o.[BRUSH CUT not utilized logs (ft3)] = 0," +
                        "o.[BRUSH CUT not utilized chips (tons)] = 0," + 
                        "o.gis_yard_dist_ft=1," + 
                        "o.slope=0," + 
                        "o.elev=0";

           
            }
            public static string UpdateOutputTableFromBinTables(string p_strOutputTableName, string p_strBinTotalsTableName, string p_strHwdBinTotalsTableName)
            {
                return
                    "UPDATE (" + p_strOutputTableName + " o " +
                    "INNER JOIN " + p_strBinTotalsTableName + " b " +
                    "ON (o.biosum_cond_id = b.biosum_cond_id AND " +
                        "o.rxpackage = b.rxpackage AND " +
                        "o.rx = b.rx AND " +
                        "o.rxcycle = b.rxcycle)) " +
                    "INNER JOIN " + p_strHwdBinTotalsTableName + " h " +
                    "ON (b.biosum_cond_id = h.biosum_cond_id AND " +
                        "b.rxpackage = h.rxpackage AND " +
                        "b.rx = h.rx AND " +
                        "b.rxcycle = h.rxcycle) " +
                    "SET o.[CHIPS TPA] = b.CHIPS_Util_Chips_TPA + b.SMLOGS_Util_Chips_TPA + b.LGLOGS_Util_Chips_TPA," +
                        "o.[CHIPS Average Vol (ft3)] = " +
                          "IIF(b.CHIPS_Util_Chips_TPA+ b.SMLOGS_Util_Chips_TPA + b.LGLOGS_Util_Chips_TPA > 0," +
                          "(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Chips_biomass_vol+b.LGLOGS_Util_Chips_biomass_vol)/(b.CHIPS_Util_Chips_TPA+ b.SMLOGS_Util_Chips_TPA + b.LGLOGS_Util_Chips_TPA),0)," +
                        "o.[CHIPS Average Weight (tons)] = " +
                          "IIF(b.CHIPS_Util_Chips_TPA+ b.SMLOGS_Util_Chips_TPA + b.LGLOGS_Util_Chips_TPA>0," +
                          "(b.CHIPS_Util_Chips_biomass_wt+b.SMLOGS_Util_Chips_biomass_wt+b.LGLOGS_Util_Chips_biomass_wt)/(b.CHIPS_Util_Chips_TPA+ b.SMLOGS_Util_Chips_TPA + b.LGLOGS_Util_Chips_TPA),0)," +
                        "o.[CHIPS Average Density (lbs/ft3)] = IIF(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Chips_biomass_vol+b.LGLOGS_Util_Chips_biomass_vol=0,0,((b.CHIPS_Util_Chips_biomass_wt+b.SMLOGS_Util_Chips_biomass_wt+b.LGLOGS_Util_Chips_biomass_wt)*2000)/(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Chips_biomass_vol+b.LGLOGS_Util_Chips_biomass_vol))," +
                        "o.[CHIPS Hwd Proportion] = IIF(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Chips_biomass_vol+b.LGLOGS_Util_Chips_biomass_vol=0,0,(h.HWD_CHIPS_Util_Chips_biomass_vol+h.HWD_SMLOGS_Util_Chips_biomass_vol+h.HWD_LGLOGS_Util_Chips_biomass_vol)/(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Chips_biomass_vol+b.LGLOGS_Util_Chips_biomass_vol))," +
                         "o.[CHIPS Chip Fraction] = IIf(o.[CHIPS Average Vol (ft3)]=0,0,100)," +
                        "o.[CHIPS utilized logs (ft3)] = b.CHIPS_Util_Logs_biomass_vol+b.SMLOGS_Util_Chips_biomass_vol+b.LGLOGS_Util_Chips_biomass_vol," +
                        "o.[CHIPS utilized chips (tons)] = b.CHIPS_Util_Chips_biomass_wt+b.SMLOGS_Util_Chips_biomass_wt+b.LGLOGS_Util_Chips_biomass_wt," +
                        "o.[SMLOGS TPA] = b.SMLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA," +
                        "o.[SMLOGS Average Vol (ft3)] = IIF(b.SMLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA=0,0,(b.SMLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol)/(b.SMLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA))," +
                        "o.[SMLOGS Average Weight (tons)] = IIF(b.SMLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA=0,0,(b.SMLOGS_Util_Logs_merch_wt+b.SMLOGS_Util_Chips_merch_wt)/(b.SMLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA))," +
                        "o.[SMLOGS Average Density (lbs/ft3)] = IIF(b.SMLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol=0,0,((b.SMLOGS_Util_Logs_merch_wt+b.SMLOGS_Util_Chips_merch_wt)*2000)/(b.SMLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol))," +
                        "o.[SMLOGS Hwd Proportion] = IIF(b.SMLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol=0,0,(h.HWD_SMLOGS_Util_Logs_merch_vol+h.HWD_SMLOGS_Util_Chips_merch_vol)/(b.SMLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol))," +
                        "o.[SMLOGS Chip Fraction] = 0," +
                        "o.[SMLOGS utilized logs (ft3)] = b.SMLOGS_Util_Logs_merch_vol," +
                        "o.[SMLOGS utilized chips (tons)] = b.SMLOGS_Util_Chips_biomass_wt + b.SMLOGS_Util_Chips_merch_wt," +
                        "o.[LGLOGS TPA] = b.LGLOGS_Util_Logs_TPA + b.LGLOGS_Util_Chips_TPA," +
                        "o.[LGLOGS Average Vol (ft3)] = IIF(b.LGLOGS_Util_Logs_TPA+b.LGLOGS_Util_Chips_TPA=0,0,(b.LGLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Chips_merch_vol)/(b.LGLOGS_Util_Logs_TPA+b.LGLOGS_Util_Chips_TPA))," +
                        "o.[LGLOGS Average Weight (tons)] = IIF(b.LGLOGS_Util_Logs_TPA+b.LGLOGS_Util_Chips_TPA=0,0,(b.LGLOGS_Util_Logs_merch_wt+b.LGLOGS_Util_Chips_merch_wt)/(b.LGLOGS_Util_Logs_TPA+b.LGLOGS_Util_Chips_TPA))," +
                        "o.[LGLOGS Average Density (lbs/ft3)] =IIF(b.LGLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Chips_merch_vol=0,0,((b.LGLOGS_Util_Logs_merch_wt+b.LGLOGS_Util_Chips_merch_wt)*2000)/(b.LGLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Chips_merch_vol))," +
                        "o.[LGLOGS Hwd Proportion] =IIF(b.LGLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Chips_merch_vol=0,0,(h.HWD_LGLOGS_Util_Logs_merch_vol+h.HWD_LGLOGS_Util_Chips_merch_vol)/(b.LGLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Chips_merch_vol))," +
                        "o.[LGLOGS Chip Fraction] = 0," +
                        "o.[LGLOGS utilized logs (ft3)] = b.LGLOGS_Util_Logs_merch_vol," +
                        "o.[LGLOGS utilized chips (tons)] = b.LGLOGS_Util_Chips_biomass_wt + b.LGLOGS_Util_Chips_merch_wt," +
                        "o.[TOTAL TPA] = b.CHIPS_Util_Chips_TPA+b.SMLOGS_Util_Logs_TPA+b.LGLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA+b.LGLOGS_Util_Chips_TPA," +
                        "o.[TOTAL Average Vol (ft3)] = IIF(b.CHIPS_Util_Chips_TPA+b.SMLOGS_Util_Logs_TPA+b.LGLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA+b.LGLOGS_Util_Chips_TPA=0,0,(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol+b.LGLOGS_Util_Chips_merch_vol)/" +
                                "(b.CHIPS_Util_Chips_TPA+b.SMLOGS_Util_Logs_TPA+b.LGLOGS_Util_Logs_TPA+b.SMLOGS_Util_Chips_TPA+b.LGLOGS_Util_Chips_TPA))," +
                        "o.[TOTAL Average Weight (tons)] = IIF(b.CHIPS_Util_Chips_TPA+b.SMLOGS_Util_Logs_TPA+b.LGLOGS_Util_Logs_TPA=0,0,(b.CHIPS_Util_Chips_biomass_wt+b.SMLOGS_Util_Logs_merch_wt+b.LGLOGS_Util_Logs_merch_wt)/" +
                                "(b.CHIPS_Util_Chips_TPA+b.SMLOGS_Util_Logs_TPA+b.LGLOGS_Util_Logs_TPA))," +
                        "o.[TOTAL Average Density (lbs/ft3)] = IIF(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol+b.LGLOGS_Util_Chips_merch_vol=0,0,((b.CHIPS_Util_Chips_biomass_wt+b.SMLOGS_Util_Logs_merch_wt+b.LGLOGS_Util_Logs_merch_wt+b.SMLOGS_Util_Chips_merch_wt+b.LGLOGS_Util_Chips_merch_wt)*2000)/" +
                                "(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol+b.LGLOGS_Util_Chips_merch_vol))," +
                        "o.[TOTAL Hwd Proportion] = IIF(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol+b.LGLOGS_Util_Chips_merch_vol=0,0,(h.HWD_CHIPS_Util_Chips_biomass_vol+h.HWD_SMLOGS_Util_Logs_merch_vol+h.HWD_LGLOGS_Util_Logs_merch_vol+h.HWD_SMLOGS_Util_Chips_merch_vol+h.HWD_LGLOGS_Util_Chips_merch_vol)/" +
                                "(b.CHIPS_Util_Chips_biomass_vol+b.SMLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Logs_merch_vol+b.SMLOGS_Util_Chips_merch_vol+b.LGLOGS_Util_Chips_merch_vol))," +
                        "o.[TOTAL Chip Fraction] = 0," +
                        "o.[TOTAL utilized logs (ft3)] = b.SMLOGS_Util_Logs_merch_vol+b.LGLOGS_Util_Logs_merch_vol," +
                        "o.[TOTAL utilized chips (tons)] = b.CHIPS_Util_Chips_biomass_wt+b.SMLOGS_Util_Chips_biomass_wt+b.LGLOGS_Util_Chips_biomass_wt+b.SMLOGS_Util_Chips_merch_wt+b.LGLOGS_Util_Chips_merch_wt," +
                        "o.[BRUSH CUT utilized logs (ft3)] = b.BC_Util_Logs_biomass_vol," +
                        "o.[BRUSH CUT utilized chips (tons)] = b.BC_Util_Chips_biomass_wt," +
                        "o.[BRUSH CUT not utilized TPA] = b.BC_NonUtil_Chips_TPA + b.BC_NonUtil_Logs_TPA," +
                        "o.[BRUSH CUT not utilized Average Vol (ft3)] = " +
                                "IIF(b.BC_NonUtil_Chips_TPA+ b.BC_NonUtil_Logs_TPA > 0," +
                                    "(b.BC_NonUtil_Chips_biomass_vol)/(b.BC_NonUtil_Chips_TPA+ b.BC_NonUtil_Logs_TPA),0)," +
                        "o.[BRUSH CUT not utilized Average Weight (tons)] = " +
                                "IIF(b.BC_NonUtil_Chips_TPA+ b.BC_NonUtil_Logs_TPA>0," +
                                   "(b.BC_NonUtil_Chips_biomass_wt+b.BC_NonUtil_Logs_biomass_wt)/(b.BC_NonUtil_Chips_TPA + b.BC_NonUtil_Logs_TPA),0)," +
                        "o.[BRUSH CUT not utilized Average Density (lbs/ft3)] = " +
                                "IIF(b.BC_NonUtil_Chips_biomass_vol+b.BC_NonUtil_Logs_biomass_vol=0,0,((b.BC_NonUtil_Chips_biomass_wt+b.BC_NonUtil_Logs_biomass_wt)*2000)/(b.BC_NonUtil_Chips_biomass_vol+b.BC_NonUtil_Logs_biomass_vol))," +
                        "o.[BRUSH CUT not utilized Hwd Proportion] = " +
                                "IIF(b.BC_NonUtil_Chips_biomass_vol+b.BC_NonUtil_Logs_biomass_vol=0,0,(h.HWD_BC_NonUtil_Chips_biomass_vol+h.HWD_BC_NonUtil_Logs_biomass_vol)/(b.BC_NonUtil_Chips_biomass_vol+b.BC_NonUtil_Logs_biomass_vol))," +
                        "o.[BRUSH CUT not utilized Chip Fraction] = IIf(o.[BRUSH CUT not utilized Average Vol (ft3)]=0,0,100)," +
                        "o.[BRUSH CUT not utilized logs (ft3)] = b.BC_NonUtil_Logs_biomass_vol," +
                        "o.[BRUSH CUT not utilized chips (tons)] = b.BC_NonUtil_Chips_biomass_wt";

            }
            public static string SumBinSumTableByPlotRxSpcGrpDbhGrp(string p_strIntoTableName, string p_strBinSumTableName, string p_strHwdBinSumTableName)
            {
                return
                    "SELECT b.BIOSUM_COND_ID,b.RXPACKAGE, b.RX, b.RXCYCLE, " +
                           "b.species_group, b.diam_group," +
                           "SUM(b.[BC_Util_Logs_merch_wt] + b.[CHIPS_Util_Logs_merch_wt] " +
                           "+ b.[SMLOGS_Util_Logs_merch_wt] + b.[LGLOGS_Util_Logs_merch_wt] " +
                           ") " +
                           "AS MERCH_WT_GT," +
                           "SUM(b.[BC_Util_Logs_merch_vol] + b.[CHIPS_Util_Logs_merch_vol] " +
                           "+ b.[SMLOGS_Util_Logs_merch_vol] + b.[LGLOGS_Util_Logs_merch_vol] " +
                           ") " +
                           "AS MERCH_VOL_CF," +
                           "SUM(b.[BC_Util_Logs_biomass_wt] + b.[CHIPS_Util_Logs_biomass_wt] " +
                           "+ b.[SMLOGS_Util_Logs_biomass_wt] + b.[LGLOGS_Util_Logs_biomass_wt] " +
                           "+ b.[SMLOGS_Util_Chips_biomass_wt] + b.[LGLOGS_Util_Chips_biomass_wt] " +
                           "+ b.[CHIPS_Util_Chips_biomass_wt]+ b.[SMLOGS_Util_Chips_merch_wt]) " +
                           "AS CHIP_WT_GT," +
                           "SUM(b.[BC_Util_Logs_biomass_vol] + b.[CHIPS_Util_Logs_biomass_vol] " +
                           "+ b.[SMLOGS_Util_Logs_biomass_vol] + b.[LGLOGS_Util_Logs_biomass_vol] " +
                           "+ b.[SMLOGS_Util_Chips_biomass_vol] + b.[LGLOGS_Util_Chips_biomass_vol] " +
                           "+ b.[CHIPS_Util_Chips_biomass_vol] + b.[LGLOGS_Util_Chips_merch_vol]) " +
                           "AS CHIP_VOL_CF," +
                           "SUM(b.[BC_NonUtil_Logs_merch_wt] + b.[CHIPS_NonUtil_Logs_merch_wt] " +
                           "+ b.[SMLOGS_NonUtil_Logs_merch_wt] + b.[LGLOGS_NonUtil_Logs_merch_wt] " +
                           ") " +
                           "AS NOT_UTILIZED_MERCH_WT_GT," +
                           "SUM(b.[BC_NonUtil_Logs_merch_vol] + b.[CHIPS_NonUtil_Logs_merch_vol] " +
                           "+ b.[SMLOGS_NonUtil_Logs_merch_vol] + b.[LGLOGS_NonUtil_Logs_merch_vol] " +
                           ") " +
                           "AS NOT_UTILIZED_MERCH_VOL_CF," +
                           "SUM(b.[BC_NonUtil_Logs_biomass_wt] + b.[CHIPS_NonUtil_Logs_biomass_wt] " +
                           "+ b.[SMLOGS_NonUtil_Logs_biomass_wt] + b.[LGLOGS_NonUtil_Logs_biomass_wt] " +
                           "+ b.[SMLOGS_NonUtil_Chips_biomass_wt] + b.[LGLOGS_NonUtil_Chips_biomass_wt] " +
                           "+ b.[SMLOGS_NonUtil_Chips_merch_wt]) " +
                           "AS NOT_UTILIZED_CHIP_WT_GT," +
                           "SUM(b.[BC_NonUtil_Logs_biomass_vol] + b.[CHIPS_NonUtil_Logs_biomass_vol] " +
                           "+ b.[SMLOGS_NonUtil_Logs_biomass_vol] + b.[LGLOGS_NonUtil_Logs_biomass_vol] " +
                           "+ b.[SMLOGS_NonUtil_Chips_biomass_vol] + b.[LGLOGS_NonUtil_Chips_biomass_vol] " +
                           "+ b.[LGLOGS_NonUtil_Chips_merch_vol]) " +
                           "AS NOT_UTILIZED_CHIP_VOL_CF, " +
                           "SUM(b.[BC_NonUtil_Chips_merch_wt] + b.[BC_NonUtil_Chips_biomass_wt]) " +
                           "AS BC_WT_GT, " +
                           "SUM(b.[BC_NonUtil_Chips_merch_vol] + b.[BC_NonUtil_Chips_biomass_vol]) " +
                           "AS BC_VOL_CF " +
                           "INTO " + p_strIntoTableName + " " +
                           "FROM " + p_strBinSumTableName + " b " +
                           "WHERE (((b.diam_group) Is Not Null)) " +
                           "GROUP BY b.BIOSUM_COND_ID, b.rxpackage, b.RX, b.rxcycle, b.species_group, b.diam_group";

            }
           
            public static string AppendToSumSoftwoodBinSumTableByPlotRxSpcGrpDbhGrp(string p_strTable, string p_strSoftwoodBinSumTableName,string p_strHardwoodBinSumTableName)
            {
                return "INSERT INTO " + p_strTable + " " +
                       "SELECT b.BIOSUM_COND_ID,b.RXPACKAGE, b.RX, b.RXCYCLE, b.species_group," +
                              "b.diam_group," +
                               "SUM(b.[BC_Util_Logs_merch_wt] + " +
                                   "b.[CHIPS_Util_Logs_merch_wt] + " +
                                   "b.[SMLOGS_Util_Logs_merch_wt] + " +
                                   "b.[LGLOGS_Util_Logs_merch_wt]) AS MERCH_WT_GT," +
                               "SUM(b.[BC_Util_Logs_merch_vol] + " +
                                   "b.[CHIPS_Util_Logs_merch_vol] + " +
                                   "b.[SMLOGS_Util_Logs_merch_vol] + " +
                                   "b.[LGLOGS_Util_Logs_merch_vol]) AS MERCH_VOL_CF," +
                               "SUM(b.[BC_Util_Logs_biomass_wt] + " +
                                   "b.[CHIPS_Util_Logs_biomass_wt] + " +
                                   "b.[SMLOGS_Util_Logs_biomass_wt] + " +
                                   "b.[LGLOGS_Util_Logs_biomass_wt] + " +
                                   "b.[SMLOGS_Util_Chips_biomass_wt] + " +
                                   "b.[LGLOGS_Util_Chips_biomass_wt] + " +
                                   "b.[SMLOGS_Util_Chips_merch_wt]) AS CHIP_WT_GT," +
                               "SUM(b.[BC_Util_Logs_biomass_vol] + b.[CHIPS_Util_Logs_biomass_vol] + " +
                                   "b.[SMLOGS_Util_Logs_biomass_vol] + " +
                                   "b.[LGLOGS_Util_Logs_biomass_vol] + " +
                                   "b.[SMLOGS_Util_Chips_biomass_vol] + " +
                                   "b.[LGLOGS_Util_Chips_biomass_vol]  + " +
                                   "b.[LGLOGS_Util_Chips_merch_vol]) AS CHIP_VOL_CF " +
                                "SUM(b.[BC_NonUtil_Logs_merch_wt] + " +
                                   "b.[CHIPS_NonUtil_Logs_merch_wt] + " +
                                   "b.[SMLOGS_NonUtil_Logs_merch_wt] + " +
                                   "b.[LGLOGS_NonUtil_Logs_merch_wt]) AS NOT_UTILIZED_MERCH_WT_GT," +
                                "SUM(b.[BC_NonUtil_Logs_merch_vol] + " +
                                   "b.[CHIPS_NonUtil_Logs_merch_vol] + " +
                                   "b.[SMLOGS_NonUtil_Logs_merch_vol] + " +
                                   "b.[LGLOGS_NonUtil_Logs_merch_vol]) AS NOT_UTILIZED_MERCH_VOL_CF," +
                                "SUM(b.[BC_NonUtil_Logs_biomass_wt] + " +
                                   "b.[CHIPS_NonUtil_Logs_biomass_wt] + " +
                                   "b.[SMLOGS_NonUtil_Logs_biomass_wt] + " +
                                   "b.[LGLOGS_NonUtil_Logs_biomass_wt] + " +
                                   "b.[SMLOGS_NonUtil_Chips_biomass_wt] + " +
                                   "b.[LGLOGS_NonUtil_Chips_biomass_wt] + " +
                                   "b.[SMLOGS_NonUtil_Chips_merch_wt]) AS NOT_UTILIZED_CHIP_WT_GT," +
                                "SUM(b.[BC_NonUtil_Logs_biomass_vol] + b.[CHIPS_NonUtil_Logs_biomass_vol] + " +
                                   "b.[SMLOGS_NonUtil_Logs_biomass_vol] + " +
                                   "b.[LGLOGS_NonUtil_Logs_biomass_vol] + " +
                                   "b.[SMLOGS_NonUtil_Chips_biomass_vol] + " +
                                   "b.[LGLOGS_NonUtil_Chips_biomass_vol]  + " +
                                   "b.[LGLOGS_NonUtil_Chips_merch_vol]) AS NOT_UTILIZED_CHIP_VOL_CF " +
                       "FROM " + p_strSoftwoodBinSumTableName + " b " +
                       "GROUP BY b.BIOSUM_COND_ID, b.rxpackage, b.RX, b.rxcycle," +
                                "b.species_group, b.diam_group";
                     
            }
            public static string AppendToSumHardwoodBinSumTableByPlotRxSpcGrpDbhGrp(string p_strTable, string p_strSoftwoodBinSumTableName, string p_strHardwoodBinSumTableName)
            {
                return "INSERT INTO " + p_strTable + " " +
                       "SELECT h.BIOSUM_COND_ID,h.RXPACKAGE, h.RX, h.RXCYCLE, h.species_group," +
                              "h.diam_group," +
                              "SUM(h.[HWD_BC_Util_Logs_merch_wt] + " +
                                  "h.[HWD_CHIPS_Util_Logs_merch_wt] + " +
                                  "h.[HWD_SMLOGS_Util_Logs_merch_wt] + " +
                                  "h.[HWD_LGLOGS_Util_Logs_merch_wt]) AS MERCH_WT_GT," +
                              "SUM(h.[HWD_BC_Util_Logs_merch_vol] + " +
                                  "h.[HWD_CHIPS_Util_Logs_merch_vol] + " +
                                  "h.[HWD_SMLOGS_Util_Logs_merch_vol] + " +
                                  "h.[HWD_LGLOGS_Util_Logs_merch_vol]) AS MERCH_VOL_CF," +
                             "SUM(h.[HWD_BC_Util_Logs_biomass_wt] + " +
                                 "h.[HWD_CHIPS_Util_Logs_biomass_wt] + " +
                                 "h.[HWD_SMLOGS_Util_Logs_biomass_wt] + " +
                                 "h.[HWD_LGLOGS_Util_Logs_biomass_wt] + " +
                                 "h.[HWD_SMLOGS_Util_Chips_biomass_wt] + " +
                                 "h.[HWD_LGLOGS_Util_Chips_biomass_wt] + " +
                                 "h.[HWD_SMLOGS_Util_Chips_merch_wt]) AS CHIP_WT_GT," +
                             "SUM(h.[HWD_BC_Util_Logs_biomass_vol] + " +
                                 "h.[HWD_CHIPS_Util_Logs_biomass_vol] + " +
                                 "h.[HWD_SMLOGS_Util_Logs_biomass_vol] + " +
                                 "h.[HWD_LGLOGS_Util_Logs_biomass_vol] + " +
                                 "h.[HWD_SMLOGS_Util_Chips_biomass_vol] + " +
                                 "h.[HWD_LGLOGS_Util_Chips_biomass_vol] +  " +
                                 "h.[HWD_LGLOGS_Util_Chips_merch_vol]) AS CHIP_VOL_CF," +
                             "SUM(h.[HWD_BC_NonUtil_Logs_merch_wt] + " +
                                 "h.[HWD_CHIPS_NonUtil_Logs_merch_wt] + " +
                                 "h.[HWD_SMLOGS_NonUtil_Logs_merch_wt] + " +
                                 "h.[HWD_LGLOGS_NonUtil_Logs_merch_wt]) AS NOT_UTILIZED_MERCH_WT_GT," +
                             "SUM(h.[HWD_BC_NonUtil_Logs_merch_vol] + " +
                                  "h.[HWD_CHIPS_NonUtil_Logs_merch_vol] + " +
                                  "h.[HWD_SMLOGS_NonUtil_Logs_merch_vol] + " +
                                  "h.[HWD_LGLOGS_NonUtil_Logs_merch_vol]) AS NOT_UTILIZED_MERCH_VOL_CF," +
                             "SUM(h.[HWD_BC_NonUtil_Logs_biomass_wt] + " +
                                 "h.[HWD_CHIPS_NonUtil_Logs_biomass_wt] + " +
                                 "h.[HWD_SMLOGS_NonUtil_Logs_biomass_wt] + " +
                                 "h.[HWD_LGLOGS_NonUtil_Logs_biomass_wt] + " +
                                 "h.[HWD_SMLOGS_NonUtil_Chips_biomass_wt] + " +
                                 "h.[HWD_LGLOGS_NonUtil_Chips_biomass_wt] + " +
                                 "h.[HWD_SMLOGS_NonUtil_Chips_merch_wt]) AS NOT_UTILIZED_CHIP_WT_GT," +
                             "SUM(h.[HWD_BC_NonUtil_Logs_biomass_vol] + " +
                                 "h.[HWD_CHIPS_NonUtil_Logs_biomass_vol] + " +
                                 "h.[HWD_SMLOGS_NonUtil_Logs_biomass_vol] + " +
                                 "h.[HWD_LGLOGS_NonUtil_Logs_biomass_vol] + " +
                                 "h.[HWD_SMLOGS_NonUtil_Chips_biomass_vol] + " +
                                 "h.[HWD_LGLOGS_NonUtil_Chips_biomass_vol] +  " +
                                 "h.[HWD_LGLOGS_NonUtil_Chips_merch_vol]) AS NOT_UTILIZED_CHIP_VOL_CF," +
                       "FROM " + p_strHardwoodBinSumTableName + " h " +
                        "GROUP BY h.BIOSUM_COND_ID, h.rxpackage, h.RX, h.rxcycle," +
                                "h.species_group, h.diam_group";
                       
            }
            public static string AppendToTreeVolVal(string p_strSourceTable, string p_strDestTable,string p_strDateTimeCreated)
            {
                return
                    "INSERT INTO " + p_strDestTable + " " +
                        "(biosum_cond_id,rxpackage,rx,rxcycle,species_group,diam_group, merch_wt_gt," +
                         "merch_vol_cf, chip_wt_gt, chip_vol_cf, bc_wt_gt, bc_vol_cf, DateTimeCreated) " +
                         "SELECT DISTINCT t.BIOSUM_COND_ID, t.rxpackage,t.rx,t.rxcycle," +
                                        "t.species_group,t.diam_group," +
                                        "t.MERCH_WT_GT,t.MERCH_VOL_CF," +
                                        "t.CHIP_WT_GT,t.CHIP_VOL_CF," +
                                        "t.BC_WT_GT,t.BC_VOL_CF," +
                                        "'" + p_strDateTimeCreated + "' " + 
                         "FROM " + p_strSourceTable + " t " +
                         "WHERE (((t.MERCH_WT_GT)<>0) OR ((t.MERCH_VOL_CF)<>0) OR " +
                                "((t.CHIP_WT_GT)<>0) OR ((t.CHIP_VOL_CF)<>0) OR " +
                                "((t.BC_WT_GT)<>0) OR ((t.BC_VOL_CF)<>0))";

            }
            
            public static string UpdateTreeVolValWithMerchChipMarketValues(
                string p_strScenarioMerchChipMarketValuesTable,
                string p_strScenarioCostRevenueEscalatorValuesTable,
                string p_strScenarioId,
                string p_strTreeVolValTable)
            {
                return "UPDATE " + p_strTreeVolValTable + " t " +
                       "INNER JOIN (" + p_strScenarioMerchChipMarketValuesTable + " v " +
                                   "INNER JOIN " + p_strScenarioCostRevenueEscalatorValuesTable + " e " + 
                                   "ON v.scenario_id = e.scenario_id) " + 
                       "ON t.species_group = v.species_group AND t.diam_group=v.diam_group " + 
                       "SET t.merch_val_dpa = " + 
                                "IIF(t.RXCycle='1',t.merch_vol_cf * v.merch_value," + 
                                "IIF(t.RXCycle='2',t.merch_vol_cf * (v.merch_value * e.EscalatorMerchWoodRevenue_Cycle2)," + 
                                "IIF(t.RXCycle='3',t.merch_vol_cf * (v.merch_value * e.EscalatorMerchWoodRevenue_Cycle3)," + 
                                "IIF(t.RXCycle='4',t.merch_vol_cf * (v.merch_value * e.EscalatorMerchWoodRevenue_Cycle4),0))))," + 
                           "t.chip_mkt_val_pgt = " +
                                "IIF(t.RXCycle='1',v.chip_value," +
                                "IIF(t.RXCycle='2',(v.chip_value * e.EscalatorEnergyWoodRevenue_Cycle2)," +
                                "IIF(t.RXCycle='3',(v.chip_value * e.EscalatorEnergyWoodRevenue_Cycle3)," +
                                "IIF(t.RXCycle='4',(v.chip_value * e.EscalatorEnergyWoodRevenue_Cycle4),0))))," + 
                           "t.chip_val_dpa = " + 
                                "IIF(t.RXCycle='1',t.chip_wt_gt * v.chip_value," + 
                                "IIF(t.RXCycle='2',t.chip_wt_gt * (v.chip_value * e.EscalatorEnergyWoodRevenue_Cycle2)," + 
                                "IIF(t.RXCycle='3',t.chip_wt_gt * (v.chip_value * e.EscalatorEnergyWoodRevenue_Cycle3)," + 
                                "IIF(t.RXCycle='4',t.chip_wt_gt * (v.chip_value * e.EscalatorEnergyWoodRevenue_Cycle4),0))))," + 
                          "merch_to_chipbin_YN = " + 
                                "IIF(v.wood_bin='M','N','Y') " + 
                       "WHERE TRIM(UCASE(v.scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'";

            }
           
           
            public static string PopulateFRCSInputTable(string p_strDestTableName,string p_strSourceTableName)
            {
                return "SELECT o.biosum_cond_id + o.rxpackage + o.rx + o.rxcycle AS Stand," +
                                "o.Slope AS [Percent Slope]," +
                                "o.gis_yard_dist_ft AS [One-way Yarding Distance]," +
                                "Null AS BLANK," +
                                "o.elev AS [Project Elevation]," +
                                "o.[Harvesting system] AS [Harvesting System]," +
                                "o.[CHIPS TPA] AS [Chip tree per acre]," +
                                "o.[CHIPS Chip Fraction] AS [Residue fraction for chip trees]," +
                                "o.[CHIPS Average Vol (ft3)] AS [Chip trees average volume(ft3)]," +
                                "o.[CHIPS Average Density (lbs/ft3)]," +
                                "o.[CHIPS Hwd Proportion]," +
                                "o.[SMLOGS TPA] AS [Small log trees per acre]," +
                                "o.[SMLOGS Chip Fraction] AS [Small log trees residue fraction]," +
                                "o.[SMLOGS Average Vol (ft3)] AS [Small log trees average volume(ft3)]," +
                                "o.[SMLOGS Average Density (lbs/ft3)] AS [Small log trees average density(lbs/ft3)]," +
                                "o.[SMLOGS Hwd Proportion] AS [Small log trees hardwood proportion]," +
                                "o.[LGLOGS TPA] AS [Large log trees per acre]," +
                                "o.[LGLOGS Chip Fraction] AS [Large log trees residue fraction]," +
                                "o.[LGLOGS Average Vol (ft3)] AS [Large log trees average vol(ft3)]," +
                                "o.[LGLOGS Average Density (lbs/ft3)] AS [Large log trees average density(lbs/ft3)]," +
                                "o.[LGLOGS Hwd Proportion] AS [Large log trees hardwood proportion]," +
                                "Null AS BLANK1," +
                                "Null AS BLANK2," +
                                "Null AS BLANK3," +
                                "Null AS BLANK4," +
                                "Null AS BLANK5," +
                                "Null AS [Costs per green ton]," +
                                "Null AS [Costs per cubic foot]," +
                                "Null AS [Costs per cubic foot1]," +
                                "o.rxpackage + o.rx + o.rxcycle AS RxPackage_Rx_RxCycle " +
                        "INTO " + p_strDestTableName + " " +
                        "FROM " + p_strSourceTableName + " o ";
            }
            public static string AppendToFRCSInputTable(string p_strDestTableName, string p_strSourceTableName)
            {
                return "INSERT INTO " + p_strDestTableName + " " +
                                "(Stand," +
                                "[Percent Slope]," +
                                "[One-way Yarding Distance]," +
                                "BLANK," +
                                "[Project Elevation]," +
                                "[Harvesting System]," +
                                "[Chip tree per acre]," +
                                "[Residue fraction for chip trees]," +
                                "[Chip trees average volume(ft3)]," +
                                "[CHIPS Average Density (lbs/ft3)]," +
                                "[CHIPS Hwd Proportion]," +
                                "[Small log trees per acre]," +
                                "[Small log trees residue fraction]," +
                                "[Small log trees average volume(ft3)]," +
                                "[Small log trees average density(lbs/ft3)]," +
                                "[Small log trees hardwood proportion]," +
                                "[Large log trees per acre]," +
                                "[Large log trees residue fraction]," +
                                "[Large log trees average vol(ft3)]," +
                                "[Large log trees average density(lbs/ft3)]," +
                                "[Large log trees hardwood proportion]," +
                                "BLANK1," +
                                "BLANK2," +
                                "BLANK3," +
                                "BLANK4," +
                                "BLANK5," +
                                "[Costs per green ton]," +
                                "[Costs per cubic foot]," +
                                "[Costs per cubic foot1]," +
                                "RxPackage_Rx_RxCycle) " +
                        "SELECT Stand," +
                                "[Percent Slope]," +
                                "[One-way Yarding Distance]," +
                                "BLANK," +
                                "[Project Elevation]," +
                                "[Harvesting System]," +
                                "[Chip tree per acre]," +
                                "[Residue fraction for chip trees]," +
                                "[Chip trees average volume(ft3)]," +
                                "[CHIPS Average Density (lbs/ft3)]," +
                                "[CHIPS Hwd Proportion]," +
                                "[Small log trees per acre]," +
                                "[Small log trees residue fraction]," +
                                "[Small log trees average volume(ft3)]," +
                                "[Small log trees average density(lbs/ft3)]," +
                                "[Small log trees hardwood proportion]," +
                                "[Large log trees per acre]," +
                                "[Large log trees residue fraction]," +
                                "[Large log trees average vol(ft3)]," +
                                "[Large log trees average density(lbs/ft3)]," +
                                "[Large log trees hardwood proportion]," +
                                "BLANK1," +
                                "BLANK2," +
                                "BLANK3," +
                                "BLANK4," +
                                "BLANK5," +
                                "[Costs per green ton]," +
                                "[Costs per cubic foot]," +
                                "[Costs per cubic foot1]," +
                                "RxPackage_Rx_RxCycle " +
                        "FROM " + p_strSourceTableName; 
                               
            }
            public static string PopulateOPCOSTInputTable(string p_strDestTableName, string p_strSourceTableName)
            {
                return "SELECT o.biosum_cond_id + o.rxpackage + o.rx + o.rxcycle AS Stand," +
                                "o.Slope AS [Percent Slope]," +
                                "o.gis_yard_dist_ft AS [One-way Yarding Distance]," +
                                "Null AS [YearCostCalc]," +
                                "o.elev AS [Project Elevation]," +
                                "o.[Harvesting system] AS [Harvesting System]," +
                                "o.[CHIPS TPA] AS [Chip tree per acre]," +
                                "o.[CHIPS Chip Fraction] AS [Residue fraction for chip trees]," +
                                "o.[CHIPS Average Vol (ft3)] AS [Chip trees average volume(ft3)]," +
                                "o.[CHIPS Average Density (lbs/ft3)]," +
                                "o.[CHIPS Hwd Proportion]," +
                                "o.[SMLOGS TPA] AS [Small log trees per acre]," +
                                "o.[SMLOGS Chip Fraction] AS [Small log trees residue fraction]," +
                                "o.[SMLOGS Average Vol (ft3)] AS [Small log trees average volume(ft3)]," +
                                "o.[SMLOGS Average Density (lbs/ft3)] AS [Small log trees average density(lbs/ft3)]," +
                                "o.[SMLOGS Hwd Proportion] AS [Small log trees hardwood proportion]," +
                                "o.[LGLOGS TPA] AS [Large log trees per acre]," +
                                "o.[LGLOGS Chip Fraction] AS [Large log trees residue fraction]," +
                                "o.[LGLOGS Average Vol (ft3)] AS [Large log trees average vol(ft3)]," +
                                "o.[LGLOGS Average Density (lbs/ft3)] AS [Large log trees average density(lbs/ft3)]," +
                                "o.[LGLOGS Hwd Proportion] AS [Large log trees hardwood proportion]," +
                                "o.[BRUSH CUT not utilized TPA] AS BrushCutTPA," +
                                "o.[BRUSH CUT not utilized Average Vol (ft3)] AS BrushCutAvgVol," +
                                "Null AS BLANK3," +
                                "Null AS BLANK4," +
                                "Null AS BLANK5," +
                                "Null AS [Costs per green ton]," +
                                "Null AS [Costs per cubic foot]," +
                                "Null AS [Costs per cubic foot1]," +
                                "o.rxpackage + o.rx + o.rxcycle AS RxPackage_Rx_RxCycle " +
                        "INTO " + p_strDestTableName + " " +
                        "FROM " + p_strSourceTableName + " o ";
            }
            public static string AppendToOPCOSTInputTable(string p_strDestTableName, string p_strSourceTableName)
            {
                return "INSERT INTO " + p_strDestTableName + " " +
                                "(Stand," +
                                "[Percent Slope]," +
                                "[One-way Yarding Distance]," +
                                "[YearCostCalc]," +
                                "[Project Elevation]," +
                                "[Harvesting System]," +
                                "[Chip tree per acre]," +
                                "[Residue fraction for chip trees]," +
                                "[Chip trees average volume(ft3)]," +
                                "[CHIPS Average Density (lbs/ft3)]," +
                                "[CHIPS Hwd Proportion]," +
                                "[Small log trees per acre]," +
                                "[Small log trees residue fraction]," +
                                "[Small log trees average volume(ft3)]," +
                                "[Small log trees average density(lbs/ft3)]," +
                                "[Small log trees hardwood proportion]," +
                                "[Large log trees per acre]," +
                                "[Large log trees residue fraction]," +
                                "[Large log trees average vol(ft3)]," +
                                "[Large log trees average density(lbs/ft3)]," +
                                "[Large log trees hardwood proportion]," +
                                "BrushCutTPA," +
                                "BrushCutAvgVol," +
                                "BLANK3," +
                                "BLANK4," +
                                "BLANK5," +
                                "[Costs per green ton]," +
                                "[Costs per cubic foot]," +
                                "[Costs per cubic foot1]," +
                                "RxPackage_Rx_RxCycle) " +
                        "SELECT Stand," +
                                "[Percent Slope]," +
                                "[One-way Yarding Distance]," +
                                "[YearCostCalc]," +
                                "[Project Elevation]," +
                                "[Harvesting System]," +
                                "[Chip tree per acre]," +
                                "[Residue fraction for chip trees]," +
                                "[Chip trees average volume(ft3)]," +
                                "[CHIPS Average Density (lbs/ft3)]," +
                                "[CHIPS Hwd Proportion]," +
                                "[Small log trees per acre]," +
                                "[Small log trees residue fraction]," +
                                "[Small log trees average volume(ft3)]," +
                                "[Small log trees average density(lbs/ft3)]," +
                                "[Small log trees hardwood proportion]," +
                                "[Large log trees per acre]," +
                                "[Large log trees residue fraction]," +
                                "[Large log trees average vol(ft3)]," +
                                "[Large log trees average density(lbs/ft3)]," +
                                "[Large log trees hardwood proportion]," +
                                "BrushCutTPA," +
                                "BrushCutAvgVol," +
                                "BLANK3," +
                                "BLANK4," +
                                "BLANK5," +
                                "[Costs per green ton]," +
                                "[Costs per cubic foot]," +
                                "[Costs per cubic foot1]," +
                                "RxPackage_Rx_RxCycle " +
                        "FROM " + p_strSourceTableName;

            }
            public static string PopulateFrcsVariableValuesTable(string p_strDestTableName,
                                                         string p_strSourceTableName,
                                                         string p_strSlopeExpr)
            {
                string sql = "";
                //removals per acre
                sql = "SELECT IIF(i.stand IS NOT NULL AND LEN(TRIM(i.stand)) " + 
                               ">= 25,MID(i.stand,1,25),i.stand) AS biosum_cond_id," + 
                            "MID(i.RxPackage_Rx_RxCycle,1,3) AS rxpackage," + 
                            "MID(i.RxPackage_Rx_RxCycle,4,3) AS rx," + 
                            "MID(i.RxPackage_Rx_RxCycle,7,1) AS rxcycle," + 
                            "i.[Harvesting System]," + 
                 			"IIF(i.[Chip tree per acre] IS NOT NULL,i.[chip tree per acre],0) " + 
                     "AS RemovalsCT,";
                     
                sql = sql + 
                                "IIF(i.[small log trees per acre] IS NOT NULL," + 
                                "i.[small log trees per acre],0) " + 
                                "AS RemovalsSLT,";
                     
                sql = sql + 
                                "IIF(i.[Large log trees per acre] IS NOT NULL," + 
                                "i.[Large log trees per acre],0) " + 
                                "AS RemovalsLLT,";
                     
                sql = sql + 
                                "RemovalsSLT + RemovalsLLT " + 
                                "AS RemovalsALT,";
                    
                sql = sql + 
                                "RemovalsCT + RemovalsSLT " + 
                                "AS RemovalsST,";
                     
                sql = sql + 
                                "RemovalsCT + RemovalsSLT + RemovalsLLT " + 
                                "AS Removals,";
                    
                //tree volume per cubic feet removed
                sql = sql + 
                                "IIF(i.[Chip trees average volume(ft3)] IS NOT NULL," + 
                                "i.[Chip trees average volume(ft3)],0) " + 
                                "AS TreeVolCT,";
                   
                sql = sql + 
    				            "IIF(i.[Large log trees average vol(ft3)] IS NOT NULL," + 
                                "i.[Large log trees average vol(ft3)],0) " + 
                                "AS TreeVolLLT,";
                   
                sql = sql + 
                                "IIF(i.[Small log trees average volume(ft3)] IS NOT NULL," + 
                                "i.[Small log trees average volume(ft3)],0) " + 
                                "AS TreeVolSLT,";
                   
                  
                    
                //volume per acre removed
                sql = sql + 
                                "RemovalsCT * TreeVolCT " + 
                                "AS VolPerAcreCT,";
                   
                sql = sql + 
                                "RemovalsSLT * TreeVolSLT " + 
                                "AS VolPerAcreSLT,";
                   
                sql = sql + 
    				            "RemovalsLLT * TreeVolLLT " + 
                                "AS VolPerAcreLLT,";
                   
                sql = sql + 
                                "VolPerAcreSLT + VolPerAcreLLT " + 
                                "AS VolPerAcreALT,";
                   
                sql = sql + 
                                "VolPerAcreCT + VolPerAcreSLT " + 
                               "AS VolPerAcreST,";
                   
                sql = sql + 
                               "(VolPerAcreCT + VolPerAcreSLT + VolPerAcreLLT) " + 
                               "AS VolPerAcre,";
                   
                   
                sql = sql + "IIF(RemovalsST > 0,VolPerAcreST/RemovalsST,0) AS TreeVolST,";
                sql = sql + "IIF(RemovalsALT > 0,VolPerAcreALT/RemovalsALT,0) AS TreeVolALT,";
                sql = sql + "IIF(Removals > 0,VolPerAcre/Removals,0) AS TreeVol,";
                sql = sql + "i.[Percent Slope] AS Slope ";
    	
                sql = sql + 
                                 "INTO " + p_strDestTableName + " " + 
                                 "FROM " + p_strSourceTableName + " i " + 
                                 "WHERE i.[Percent Slope] " + p_strSlopeExpr ;

                return sql;
            }
 
            public static string AppendToOPCOSTHarvestCostsTable(string p_strOPCOSTOutputTableName,
                                                          string p_strOPCOSTInputTableName,
                                                          string p_strHarvestCostsTableName,
                                                          string p_strDateTimeCreated)
            {
                return "INSERT INTO " + p_strHarvestCostsTableName + " " +
                    "(biosum_cond_id, RXPackage, RX, RXCycle, " +
                    "harvest_cpa, chip_cpa, assumed_movein_cpa, " +
                    "ideal_complete_cpa, ideal_harvest_cpa, ideal_chip_cpa, ideal_assumed_movein_cpa, " +
                    "override_YN, DateTimeCreated )" +
                    "SELECT o.biosum_cond_id, o.RxPackage,o.RX,o.RXCycle, " +
                    "IIF (RIGHT(CSTR(o.harvest_cpa), 6) = '1.#INF', 0,o.harvest_cpa ), " +
                    "o.chip_cpa, o.assumed_movein_cpa, " +
                    "0, 0, 0, 0, " +
                    "IIF(n.[Unadjusted One-way Yarding distance] > 0 OR n.[Unadjusted Small log trees per acre] > 0 " +
                    "OR n.[Unadjusted Small log trees average volume (ft3)] > 0 OR n.[Unadjusted Large log trees per acre] > 0 " +
                    "OR n.[Unadjusted Large log trees average vol(ft3)] >0, 'Y','N') , " +
                    "'" + p_strDateTimeCreated + "' AS DateTimeCreated " +
                    "from (" + p_strOPCOSTOutputTableName + " o " +
                    "INNER JOIN " + p_strOPCOSTInputTableName + " n ON (o.biosum_cond_id = n.biosum_cond_id) AND " +
                    "(o.rxPackage = n.rxPackage) AND (o.RX = n.RX) AND (o.RXCycle = n.rxCycle)) ";
            }
            public static string AppendToHarvestCostsTable(string p_strFRCSOutputTableName,
                                                           string p_strHarvestCostsTableName,
                                                           string p_strDateTimeCreated)
            {
                return "INSERT INTO " + p_strHarvestCostsTableName + " " +
                              "(biosum_cond_id,RXPackage,RX,RXCycle,harvest_cpa,DateTimeCreated ) " +
                               "SELECT MID(STAND,1,25) AS biosum_cond_id," +
                                       "MID(rxpackage_rx_rxcycle,1,3) AS RXPackage," +
                                       "MID(rxpackage_rx_rxcycle,4,3) AS RX," +
                                       "MID(rxpackage_rx_rxcycle,7,1) AS RXCycle," +
                                       "IIF([$/Acre] IS NOT NULL," +
                                       "IIF(LEN(TRIM([$/Acre])) > 0 AND " +
                                           "ASC([$/Acre]) > 64 AND " +
                                           "ASC([$/Acre]) < 122 OR " +
                                           "ASC([$/Acre]) = 32 OR " +
                                           "ASC([$/Acre]) = 36,null,CDBL([$/Acre])),null) AS harvest_cpa, " +
                                       "'" + p_strDateTimeCreated + "' AS DateTimeCreated " +
                                "FROM " + p_strFRCSOutputTableName;
            }
            public static string UpdateHarvestCostsTableWithCompleteCostsPerAcre(
                string p_strScenarioCostRevenueEscalatorValuesTableName,
                string p_strTotalAdditionalCostsTableName,
                string p_strHarvestCostsTableName,
                string p_strScenarioId,
                bool p_bIncludeZeroHarvestCpa, bool p_bUsingSqlite)
            {
                string strSql = "";
                if (!p_bUsingSqlite)
                {
                    strSql = "UPDATE " + p_strHarvestCostsTableName + " h " +
                            "INNER JOIN " + p_strTotalAdditionalCostsTableName + " a " +
                            "ON h.biosum_cond_id = a.biosum_cond_id AND h.rx=a.rx," +
                            "scenario_cost_revenue_escalators e " +
                            "SET h.complete_cpa = " +
                            "IIF(h.RXCycle='1',(h.harvest_cpa+a.complete_additional_cpa)," +
                            "IIF(h.RXCycle='2',(h.harvest_cpa+a.complete_additional_cpa) * e.EscalatorOperatingCosts_Cycle2," +
                            "IIF(h.RXCycle='3',(h.harvest_cpa+a.complete_additional_cpa) * e.EscalatorOperatingCosts_Cycle3," +
                            "IIF(h.RXCycle='4',(h.harvest_cpa+a.complete_additional_cpa) * e.EscalatorOperatingCosts_Cycle4,0)))) ";
                    if (p_bIncludeZeroHarvestCpa == false)
                    {
                        strSql += "WHERE h.harvest_cpa IS NOT NULL AND h.harvest_cpa > 0  AND ";
                    }
                    else
                    {
                        strSql += "WHERE h.harvest_cpa IS NOT NULL AND ";
                    }
                    strSql += "TRIM(UCASE(e.scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'";
                }
                else
                {
                    strSql = "UPDATE " + p_strHarvestCostsTableName + " AS h " +
                        "SET complete_cpa = " +
                        "(SELECT CASE WHEN h.RXCycle = '1' THEN h.harvest_cpa + a.complete_additional_cpa " +
                        "WHEN h.RXCycle = '2' THEN h.harvest_cpa + a.complete_additional_cpa * " +
                        "(SELECT EscalatorOperatingCosts_Cycle2 from definitions.scenario_cost_revenue_escalators " +
                        "WHERE TRIM(scenario_id) = '" + p_strScenarioId.Trim() + "') " +
                        "WHEN h.RXCycle = '3' THEN h.harvest_cpa + a.complete_additional_cpa * " +
                        "(SELECT EscalatorOperatingCosts_Cycle3 from definitions.scenario_cost_revenue_escalators " +
                        "WHERE TRIM(scenario_id) = '" + p_strScenarioId.Trim() + "') " +
                        "WHEN h.RXCycle = '4' THEN h.harvest_cpa + a.complete_additional_cpa * " +
                        "(SELECT EscalatorOperatingCosts_Cycle4 from definitions.scenario_cost_revenue_escalators " +
                        "WHERE TRIM(scenario_id) = '" + p_strScenarioId.Trim() + "') " +
                        "ELSE 0 END " +
                        "FROM " + p_strTotalAdditionalCostsTableName + " AS a " +
                        "WHERE h.biosum_cond_id = a.biosum_cond_id AND h.RX = a.RX ";
                    if (p_bIncludeZeroHarvestCpa == false)
                    {
                        strSql += "AND h.harvest_cpa IS NOT NULL AND h.harvest_cpa > 0) ";
                    }
                    else
                    {
                        strSql += "AND h.harvest_cpa IS NOT NULL) ";
                    }
                    strSql += "WHERE EXISTS( SELECT* FROM " + p_strTotalAdditionalCostsTableName + " AS a " +
                              "WHERE h.biosum_cond_id = a.biosum_cond_id AND h.RX = a.RX ";
                    if (p_bIncludeZeroHarvestCpa == false)
                    {
                        strSql += "AND h.harvest_cpa IS NOT NULL AND h.harvest_cpa > 0) ";
                    }
                    else
                    {
                        strSql += "AND h.harvest_cpa IS NOT NULL) ";
                    }
                }
                    return strSql;
            }
            //2018-30-JUL: No longer calculating ideal costs
            //See uc_processor_scenario_run.RunScenario_UpdateHarvestCostsTableWithAdditionalCosts() to re-implement
            //Commented out for now
            public static string UpdateHarvestCostsTableWithIdealCompleteCostsPerAcre(
                string p_strScenarioCostRevenueEscalatorValuesTableName,
                string p_strTotalAdditionalCostsTableName,
                string p_strHarvestCostsTableName,
                string p_strScenarioId)
            {
                return "UPDATE " + p_strHarvestCostsTableName + " h " +
                       "INNER JOIN " + p_strTotalAdditionalCostsTableName + " a " +
                       "ON h.biosum_cond_id = a.biosum_cond_id AND h.rx=a.rx," +
                       "scenario_cost_revenue_escalators e " +
                       "SET h.ideal_complete_cpa = " +
                           "IIF(h.RXCycle='1',(h.ideal_harvest_cpa+a.complete_additional_cpa)," +
                           "IIF(h.RXCycle='2',(h.ideal_harvest_cpa+a.complete_additional_cpa) * e.EscalatorOperatingCosts_Cycle2," +
                           "IIF(h.RXCycle='3',(h.ideal_harvest_cpa+a.complete_additional_cpa) * e.EscalatorOperatingCosts_Cycle3," +
                           "IIF(h.RXCycle='4',(h.ideal_harvest_cpa+a.complete_additional_cpa) * e.EscalatorOperatingCosts_Cycle4,0)))) " +
                       "WHERE h.ideal_harvest_cpa IS NOT NULL AND h.ideal_harvest_cpa > 0  AND " +
                       "TRIM(UCASE(e.scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'";
            }
                  

           
           
        }
        
		public class Reference
		{
			
			public string m_strRefHarvestMethodTable="";

			private Queries _oQueries=null;	
			private bool _bLoadDataSources=true;
			string m_strSql="";

			public Reference()
			{
			}

			public Queries ReferenceQueries
			{
				get {return _oQueries;}
				set {_oQueries=value;}
			}
			public bool LoadDatasource
			{
				get {return _bLoadDataSources;}
				set {_bLoadDataSources=value;}
			}


			public void LoadDatasources()
			{
				m_strRefHarvestMethodTable = ReferenceQueries.m_oDataSource.getValidDataSourceTableName("HARVEST METHODS");
				
			
				if (m_strRefHarvestMethodTable.Trim().Length == 0 && this._oQueries._strScenarioType!="optimizer") 
				{
					
					MessageBox.Show("!!Could Not Locate Harvest Methods Reference Table!!","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					ReferenceQueries.m_intError=-1;
					return;
				}					
			
			}
			/// <summary>
			/// 
			/// </summary>
			/// <param name="p_strTable">Harvest Method reference table</param>
			/// <param name="p_strFields">comma-delimited list of field names</param>
			/// <param name="p_strSteepSlopeYN">comma-delimited list of 'Y','N', or 'Y','N'</param>
			/// <returns></returns>
			public static string HarvestMethodTable_Select(string p_strTable,string p_strFields,string p_strWhereExpression)
			{
				return "SELECT " + p_strFields + " FROM " + p_strTable + " WHERE " + p_strWhereExpression;
			}

		}
        public class Utilities
        {
            public Utilities()
			{
			}
            /// <summary>
            /// Assign a sequential unique row number to each row.
            /// </summary>
            /// <param name="p_strIntoTableName"></param>
            /// <param name="p_strTableName"></param>
            /// <param name="p_strColumnName">Each data item for the column name must be unique</param>
            /// <param name="p_strRowNumberColumnName"></param>
            /// <returns></returns>
            public static string AssignRowNumberToEachRow(string p_strIntoTableName,string p_strTableName,string p_strColumnName,string p_strRowNumberColumnName)
            {
                string strSQL = "";

               
                if (p_strIntoTableName.Trim().Length > 0)
                {
                    strSQL = "SELECT (SELECT COUNT(a." + p_strColumnName + ") " +
                                     "FROM " + p_strTableName + " a " +
                                     "WHERE a." + p_strColumnName + " <= b." + p_strColumnName + ") AS " + p_strRowNumberColumnName + "," +
                                     "b." + p_strColumnName + " " +
                             "INTO " + p_strIntoTableName + " " + 
                             "FROM " + p_strTableName + " b " + 
                             "ORDER BY b." + p_strColumnName;
                }
                else
                {
                    strSQL = "SELECT (SELECT COUNT(a." + p_strColumnName + ") " +
                                     "FROM " + p_strTableName + " a " +
                                     "WHERE a." + p_strColumnName + " <= b." + p_strColumnName + ") AS " + p_strRowNumberColumnName + "," +
                                     "b." + p_strColumnName + " " +
                             "FROM " + p_strTableName + " b " + 
                             "ORDER BY b." + p_strColumnName;
                }

                return strSQL;
            }

        }
	}
}

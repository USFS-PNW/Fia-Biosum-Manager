using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using SQLite.ADO;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Get Project and Core Analysis Datasource information
	/// and create links to the datasource tables.
	/// </summary>
	public class Datasource
	{

		public int m_intError;
		public string m_strError;
		public string m_strTable;
		public const int TABLETYPE = 0;
		public const int PATH = 1;
		public const int DBFILE = 2;
		public const int FILESTATUS = 3;
		public const int TABLE = 4;
		public const int TABLESTATUS = 5;
		public const int RECORDCOUNT = 6;
		public const int COLUMN_LIST = 7;
		public const int DATATYPE_LIST=8;
		public const int MACROVAR=9;
		public string m_strRandomPathAndFile = "";
		public int m_intNumberOfValidTables=0;  //MDB file is FOUND and table is FOUND
        /// <summary>
        /// 1st Array Item: sequence of datasource;
        /// 2nd Array Item:
        /// 0=TABLETYPE,1=DIRECTORY PATH,2=MDBFILE,3=FILESTATUS,4=TABLE,5=TABLESTATUS
        /// ,6=RECORDCOUNT,7=COLUMN_LIST (COMMA-DELIMITED),8=DATATYPE_LIST (COMMA-DELIMITED)
        /// 9=MACRO VARIABLE NAME
        /// </summary>
		public string[,] m_strDataSource;
		public string m_strScenarioId;
		public string m_strDataSourceDBFile;
		public string m_strDataSourceTableName;
		public int m_intNumberOfTables;
		bool _bLoadFieldNamesAndDatatypes=false;
		bool _bLoadTableRecordCount=true;
	    public static string[] g_strProjectDatasourceTableTypesArray =
	    {
			"Plot",
			"Condition",
			"Tree",
			"Owner Groups",
			"Treatment Prescriptions",
			"Treatment Prescriptions Harvest Cost Columns",
			"Treatment Packages",
			Datasource.TableTypes.SeqNumDefinitions,
            Datasource.TableTypes.SeqNumRxPackageAssign, 
			Datasource.TableTypes.FvsTreeSpecies,
			"Travel Times",
			"Processing Sites",
			"FVS Tree List For Processor",
			"FIA Tree Macro Plot Breakpoint Diameter",
			Datasource.TableTypes.HarvestMethods,
			Datasource.TableTypes.PopStratumAdjFactors,
			"Site Tree",
			Datasource.TableTypes.FiaTreeSpeciesReference,
	    };

		public static FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem g_oCurrentSQLMacroSubstitutionVariableItem=
			            new SQLMacroSubstitutionVariableItem();

		


		public Datasource()
		{

		}
		///<summary>
		///Access project datasource information and functionality.
		///Project Datasource information is loaded
		///into strDataSource[,] array 
		///from the constructor.
	    ///</summary>
	    ///<param name="strProjDir">Project Root Directory</param>
		public Datasource(string strProjDir)
		{
			this.m_strDataSourceDBFile = strProjDir + "\\db\\project.db";
			this.m_strDataSourceTableName = "datasource";
			this.m_strScenarioId="";
			this.populate_datasource_array();
		}
		~Datasource()
		{
		}
		
		///<summary>
		///Load a 2 dimensional array with 
		///this datasource information:
		///Table Type, MDB paths and files, table names, file/table
		///found, table record count.
		///OPTIONAL: table columns and datatypes will also load into the array
		///          if the variable LoadColumnNamesAndDataTypes is set to true
 	    ///</summary>
		public void populate_datasource_array_access()
		{
            macrosubst oMacroSub = new macrosubst();
            oMacroSub.ReferenceGeneralMacroSubstitutionVariableCollection = frmMain.g_oGeneralMacroSubstitutionVariable_Collection;
            
            int intRecCnt=0;    		
			string strPathAndFile="";
			string strSQL="";
			string strConn="";
			this.m_intNumberOfTables=0;

            FIA_Biosum_Manager.ado_data_access  p_ado = new ado_data_access();
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();


            System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
			strConn = p_ado.getMDBConnString(this.m_strDataSourceDBFile,"admin","");
			oConn.ConnectionString = strConn;
			this.m_intError = 0;
			this.m_strError = "";
			try
			{
				oConn.Open();
			}
			catch (System.Data.OleDb.OleDbException oleException)
			{
				this.m_strError = "Failed to make an oleDb connection with " + strConn;
				MessageBox.Show (this.m_strError + " OledbError=" + oleException.Message);
				this.m_intError = -1;
				return;
			}
			intRecCnt = Convert.ToInt32(p_ado.getRecordCount(oConn,"select count(*) from " + this.m_strDataSourceTableName,this.m_strDataSourceTableName));
			this.m_strDataSource = new String[intRecCnt,10];
			System.Data.OleDb.OleDbCommand oCommand = oConn.CreateCommand();
		    if (this.m_strScenarioId.Trim().Length > 0)
		    {
		        oCommand.CommandText = "select table_type,path,file,table_name from scenario_datasource" +
		                               " where scenario_id = '" + this.m_strScenarioId.Trim() + "';";
		    }
		    else
		    {
		        oCommand.CommandText =
		            "select table_type,path,file,table_name from " + this.m_strDataSourceTableName + ";";
			}
				                           
			try
			{
				System.Data.OleDb.OleDbDataReader oDataReader = oCommand.ExecuteReader();
				int x = 0;
				
				ado_data_access oExistsAdo = new ado_data_access();
                using (var oExistsConn = new System.Data.OleDb.OleDbConnection())
                {
                    while (oDataReader.Read())
                    {
                        this.m_intNumberOfTables++;
                        // Add a ListItem object to the ListView.
                        this.m_strDataSource[x, TABLETYPE] = oDataReader["table_type"].ToString().Trim();
                        this.m_strDataSource[x, PATH] = oDataReader["path"].ToString().Trim();
                        this.m_strDataSource[x, DBFILE] = oDataReader["file"].ToString().Trim();
                        strPathAndFile = oMacroSub.GeneralTranslateVariableSubstitution(oDataReader["path"].ToString().Trim()) +
                            "\\" + oDataReader["file"].ToString().Trim();
                        if (System.IO.File.Exists(strPathAndFile) == true)
                        {
                            this.m_strDataSource[x, FILESTATUS] = "F";
                            this.m_strDataSource[x, TABLE] = oDataReader["table_name"].ToString().Trim();
                            string strExistsConn = oExistsAdo.getMDBConnString(strPathAndFile, "", "");
                            bool bSQLite = false;
                            if (System.IO.Path.GetExtension(this.m_strDataSource[x, DBFILE]).Trim().ToUpper().Equals(".DB"))
                            {
                                // This is an SQLite data source
                                bSQLite = true;
                            }

                            // this is the first time the connection is used -> not open yet
                            if (!bSQLite)
                            {
                                if (String.IsNullOrEmpty(oExistsConn.ConnectionString))
                                {
                                    oExistsConn.ConnectionString = strExistsConn;
                                    oExistsConn.Open();
                                }
                                else
                                {
                                    // close and reopen the connection if the target database has changed
                                    // the connectionString returned by the connection doesn't include the "Password" key that is included
                                    // in strExistsConn
                                    if (oExistsConn.ConnectionString + "Password=;" != strExistsConn)
                                    {
                                        if (oExistsConn.State != ConnectionState.Closed)
                                        {
                                            oExistsConn.Close();
                                            oExistsConn.ConnectionString = strExistsConn;
                                            oExistsConn.Open();
                                        }
                                    }
                                }
                            }
                            
                            //see if the table exists in the mdb database container
                            if (!bSQLite && oExistsAdo.TableExist(oExistsConn, oDataReader["table_name"].ToString().Trim()) == true)
                            {
                                this.m_strDataSource[x, TABLESTATUS] = "F";
                                this.m_strDataSource[x, RECORDCOUNT] = "0";
                                this.m_strDataSource[x, COLUMN_LIST] = "";
                                this.m_strDataSource[x, DATATYPE_LIST] = "";

                                if (this.LoadTableRecordCount || this.LoadTableColumnNamesAndDataTypes)
                                {
                                    strConn = p_ado.getMDBConnString(strPathAndFile, "admin", "");
                                    p_ado.OpenConnection(strConn);
                                    if (p_ado.m_intError == 0)
                                    {
                                        strSQL = "select count(*) from " + oDataReader["table_name"].ToString();
                                        if (this.LoadTableRecordCount) this.m_strDataSource[x, RECORDCOUNT] = Convert.ToString(p_ado.getRecordCount(strConn, strSQL, oDataReader["table_name"].ToString()));
                                        if (this.LoadTableColumnNamesAndDataTypes) p_ado.getFieldNamesAndDataTypes(strConn, "select * from " + oDataReader["table_name"].ToString(), ref this.m_strDataSource[x, COLUMN_LIST], ref this.m_strDataSource[x, DATATYPE_LIST]);
                                        p_ado.CloseConnection(p_ado.m_OleDbConnection);
                                        while (p_ado.m_OleDbConnection.State != ConnectionState.Closed)
                                            System.Threading.Thread.Sleep(5000);
                                        p_ado.m_OleDbConnection.Dispose();
                                    }
                                }
                            }
                            else if (bSQLite)
                            {
                                string strDbConn = oDataMgr.GetConnectionString(strPathAndFile);
                                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strDbConn))
                                {
                                    conn.Open();
                                    if (oDataMgr.TableExist(conn, oDataReader["table_name"].ToString().Trim()) == true)
                                    {
                                        this.m_strDataSource[x, TABLESTATUS] = "F";
                                        this.m_strDataSource[x, RECORDCOUNT] = "0";
                                        this.m_strDataSource[x, COLUMN_LIST] = "";
                                        this.m_strDataSource[x, DATATYPE_LIST] = "";

                                        if (this.LoadTableRecordCount || this.LoadTableColumnNamesAndDataTypes)
                                        {
                                            oDataMgr.CloseConnection(conn);
                                            string strDsConn = oDataMgr.GetConnectionString(strPathAndFile);
                                            using (System.Data.SQLite.SQLiteConnection deetsConn = new System.Data.SQLite.SQLiteConnection(strDsConn))
                                            {
                                                deetsConn.Open();
                                                if (oDataMgr.m_intError == 0)
                                                {
                                                    strSQL = "select count(*) from " + oDataReader["table_name"].ToString();
                                                    if (this.LoadTableRecordCount) this.m_strDataSource[x, RECORDCOUNT] = Convert.ToString(oDataMgr.getRecordCount(deetsConn, strSQL, oDataReader["table_name"].ToString()));
                                                    if (this.LoadTableColumnNamesAndDataTypes) oDataMgr.getFieldNamesAndDataTypes(deetsConn, "select * from " + oDataReader["table_name"].ToString(), ref this.m_strDataSource[x, COLUMN_LIST], ref this.m_strDataSource[x, DATATYPE_LIST]);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
										this.m_strDataSource[x, TABLESTATUS] = "NF";
									}
								}
                            }
                            else
                            {
                                this.m_strDataSource[x, TABLESTATUS] = "NF";
                                this.m_strDataSource[x, RECORDCOUNT] = "0";
                            }
                        }
                        else
                        {
                            this.m_strDataSource[x, FILESTATUS] = "NF";
                            this.m_strDataSource[x, TABLE] = oDataReader["table_name"].ToString().Trim();
                            this.m_strDataSource[x, TABLESTATUS] = "NF";
                            this.m_strDataSource[x, RECORDCOUNT] = "0";
                        }
                        UpdateTableMacroVariable(this.m_strDataSource[x, TABLETYPE], this.m_strDataSource[x, TABLE]);

                        x++;
                    }
				}
                oExistsAdo = null;
				oDataReader.Close();
                oDataReader.Dispose();
			}
			catch
			{
				this.m_intError = -1;
				this.m_strError = "The Query Command " + oCommand.CommandText.ToString() + " Failed";
				MessageBox.Show(this.m_strError);
				if (oConn != null)
				{
					if (oConn.State==System.Data.ConnectionState.Open)
					{
						oConn.Close();
						while (oConn.State != System.Data.ConnectionState.Closed)
							System.Threading.Thread.Sleep(1000);
                        oConn.Dispose();
						oConn = null;
					}
				}
				p_ado= null;
				return;
			}
			finally
			{
				if (oConn != null)
				{
					if (oConn.State==System.Data.ConnectionState.Open)
					{
						oConn.Close();
						while (oConn.State != System.Data.ConnectionState.Closed)
							System.Threading.Thread.Sleep(1000);
                        oConn.Dispose();
						oConn = null;
					}
				}
			}
			
			if (oConn != null)
			{
				if (oConn.State==System.Data.ConnectionState.Open)
				{
					oConn.Close();
					while (oConn.State != System.Data.ConnectionState.Closed)
						System.Threading.Thread.Sleep(1000);

					oConn = null;
				}
			}
			p_ado = null;
		}

        ///<summary>
        ///Load a 2 dimensional array with 
        ///this datasource information:
        ///Table Type, DB paths and files, table names, file/table
        ///found, table record count.
        ///OPTIONAL: table columns and datatypes will also load into the array
        ///          if the variable LoadColumnNamesAndDataTypes is set to true
        ///</summary>

		public void populate_datasource_array()
        {
			macrosubst oMacroSub = new macrosubst();
			oMacroSub.ReferenceGeneralMacroSubstitutionVariableCollection = frmMain.g_oGeneralMacroSubstitutionVariable_Collection;

			int intRecCnt = 0;
			string strPathAndFile = "";
			string strSQL = "";
			string strConn = "";
			this.m_intNumberOfTables = 0;

			DataMgr dataMgr = new DataMgr();

			this.m_intError = 0;
			this.m_strError = "";

			strConn = dataMgr.GetConnectionString(this.m_strDataSourceDBFile);
			using (System.Data.SQLite.SQLiteConnection sqlConn = new System.Data.SQLite.SQLiteConnection(strConn))
			{
				sqlConn.Open();
				intRecCnt = Convert.ToInt32(dataMgr.getRecordCount(sqlConn, "select count(*) from " + this.m_strDataSourceTableName, this.m_strDataSourceTableName));
				this.m_strDataSource = new String[intRecCnt, 10];
				if (this.m_strScenarioId.Trim().Length > 0)
				{
					strSQL = "select table_type,path,file,table_name from scenario_datasource" +
							 " where TRIM(UPPER(scenario_id)) = '" + this.m_strScenarioId.Trim().ToUpper() + "';";
				}
				else
				{
					strSQL = "select table_type,path,file,table_name from " + this.m_strDataSourceTableName + ";";
				}
				try
				{
					dataMgr.SqlQueryReader(sqlConn, strSQL);
					int x = 0;
					while (dataMgr.m_DataReader.Read())
					{
						this.m_intNumberOfTables++;
						// Add a ListItem object to the array
						this.m_strDataSource[x, TABLETYPE] = dataMgr.m_DataReader["table_type"].ToString().Trim();
						this.m_strDataSource[x, PATH] = dataMgr.m_DataReader["path"].ToString().Trim();
						this.m_strDataSource[x, DBFILE] = dataMgr.m_DataReader["file"].ToString().Trim();
						strPathAndFile = oMacroSub.GeneralTranslateVariableSubstitution(dataMgr.m_DataReader["path"].ToString().Trim()) +
							"\\" + dataMgr.m_DataReader["file"].ToString().Trim();
						if (System.IO.File.Exists(strPathAndFile) == true)
						{
							this.m_strDataSource[x, FILESTATUS] = "F";
							this.m_strDataSource[x, TABLE] = dataMgr.m_DataReader["table_name"].ToString().Trim();DataMgr oExistsDataMgr = new DataMgr();
							string strExistsConn = oExistsDataMgr.GetConnectionString(strPathAndFile);
							using (System.Data.SQLite.SQLiteConnection existsConn = new System.Data.SQLite.SQLiteConnection(strExistsConn))
							{
								existsConn.Open();
								//see if the table exists in the db database container
								if (oExistsDataMgr.TableExist(existsConn, dataMgr.m_DataReader["table_name"].ToString().Trim()))
								{
									this.m_strDataSource[x, TABLESTATUS] = "F";
									this.m_strDataSource[x, RECORDCOUNT] = "0";
									this.m_strDataSource[x, COLUMN_LIST] = "";
									this.m_strDataSource[x, DATATYPE_LIST] = "";

									if (this.LoadTableRecordCount || this.LoadTableColumnNamesAndDataTypes)
									{
										DataMgr sourceDataMgr = new DataMgr();
										strConn = sourceDataMgr.GetConnectionString(strPathAndFile);
										sourceDataMgr.OpenConnection(strConn);
										if (sourceDataMgr.m_intError == 0)
										{
											strSQL = "select count(*) from " + dataMgr.m_DataReader["table_name"].ToString();
											if (this.LoadTableRecordCount)
											{
												this.m_strDataSource[x, RECORDCOUNT] = Convert.ToString(sourceDataMgr.getRecordCount(strConn, strSQL, dataMgr.m_DataReader["table_name"].ToString()));
											}
											if (this.LoadTableColumnNamesAndDataTypes)
											{
												sourceDataMgr.getFieldNamesAndDataTypes(sourceDataMgr.m_Connection, "select * from " + dataMgr.m_DataReader["table_name"].ToString(), ref this.m_strDataSource[x, COLUMN_LIST], ref this.m_strDataSource[x, DATATYPE_LIST]);
											}
										}
									}
								}
								else
								{
									this.m_strDataSource[x, TABLESTATUS] = "NF";
									this.m_strDataSource[x, RECORDCOUNT] = "0";
								}
							}
						}
						else
						{
							this.m_strDataSource[x, FILESTATUS] = "NF";
							this.m_strDataSource[x, TABLE] = dataMgr.m_DataReader["table_name"].ToString().Trim();
							this.m_strDataSource[x, TABLESTATUS] = "NF";
							this.m_strDataSource[x, RECORDCOUNT] = "0";
						}
						UpdateTableMacroVariable(this.m_strDataSource[x, TABLETYPE], this.m_strDataSource[x, TABLE]);

						x++;
					}
					dataMgr.m_DataReader.Close();
				}
				catch (Exception e)
				{
					this.m_intError = -1;
					this.m_strError = "The Query Command " + strSQL + " Failed";
					MessageBox.Show(this.m_strError);
					return;
				}
			}
		}

		///<summary>
		///Validate the existance of the table in the datasource
		///</summary>
		/// <param name="strTableName">Table name to validate.</param>
		public bool DataSourceTableExist(string strTableName)
		{
			int x;
			for (x=0; x <= this.m_intNumberOfTables-1; x++)
			{
				if (this.m_strDataSource[x,TABLE].Trim().ToUpper()==strTableName.Trim().ToUpper())
				{
					if (this.m_strDataSource[x,TABLESTATUS].Trim().ToUpper()=="NF")
					{
						return false;
				}
					if (this.m_strDataSource[x,FILESTATUS].Trim().ToUpper()=="NF")
					{
						return false;
					}
					return true;
				}   
			}
			return false;
		}

        public bool DataSourceTableExist(int intTableType)
        {
            return intTableType > -1 &&
                   this.m_strDataSource[intTableType, Datasource.FILESTATUS] == "F" &&
                   this.m_strDataSource[intTableType, Datasource.TABLESTATUS] == "F";
        }

		
		public string CreateDB()
        {
			macrosubst oMacroSub = new macrosubst();
			oMacroSub.ReferenceGeneralMacroSubstitutionVariableCollection = frmMain.g_oGeneralMacroSubstitutionVariable_Collection;
			int x;

			FIA_Biosum_Manager.env p_env = new env();
			this.m_intNumberOfValidTables = 0;

			// used to get the temporary random file name
			utils p_utils = new utils();

			// used to create a link to the table
			DataMgr p_dataMgr = new DataMgr();

			string strTempDB = p_utils.getRandomFile(p_env.strTempDir, "db");
			p_dataMgr.CreateDbFile(strTempDB);
			
			if (strTempDB.Trim().Length == 0)
            {
				MessageBox.Show("!!None of the data source tables are found!!");
			}
			return strTempDB;
		}

		
	///<summary>
	///Return the location of the specified table within the m_strDataSource array.
	///-1 is returned if the strTableType is not found or the MDB file is not
	///found or the table is not found
	///</summary>
	/// <param name="strTableType">The unique id for the datasource table</param>
		public int getValidTableNameRow(string strTableType)
		{
			int x;
			for (x=0; x<= this.m_intNumberOfTables-1;x++)
			{
				if (strTableType.Trim().ToUpper() == 
					this.m_strDataSource[x,TABLETYPE].Trim().ToUpper()
					&&
					this.m_strDataSource[x,FILESTATUS].Trim().ToUpper() =="F" 
					&&
					this.m_strDataSource[x,TABLESTATUS].Trim().ToUpper() == "F")
				{
					  break;
				}
			}
			if (x > this.m_intNumberOfTables-1)
			{
				x=-1;
			}
			return x;
			
		}
		///<summary>
		///Return the location of the specified table within the m_strDataSource array.
		///</summary>
		/// <param name="strTableType">The unique id for the datasource table</param>
		public int getTableNameRow(string strTableType)
		{
			int x;
			for (x=0; x<= this.m_intNumberOfTables-1;x++)
			{
				if (strTableType.Trim().ToUpper() == 
					this.m_strDataSource[x,TABLETYPE].Trim().ToUpper())
				{
					break;
				}
			}
			if (x > this.m_intNumberOfTables-1)
			{
				x=-1;
			}
			return x;
			
		}

    ///<summary>
    ///number of tables that are found
    ///</summary>
		public void getNumberOfValidTables()
		{
			int x;
			this.m_intNumberOfValidTables=0;
			for (x=0; x <= this.m_intNumberOfTables - 1; x++)
			{
				if (this.m_strDataSource[x,TABLESTATUS].Trim().ToUpper()=="F" &&
					this.m_strDataSource[x,FILESTATUS].Trim().ToUpper()=="F")
				{
					this.m_intNumberOfValidTables++;
				}
			}

		}

    ///<summary>
    ///Add the datasource tables to a list box
    ///</summary>
    ///<param name="listbox1">datasource table names will be added to the specified listbox object</param>
		public void LoadDataSourceTablesIntoListBox(System.Windows.Forms.ListBox listbox1)
		{
			int x;
			for (x=0; x <= this.m_intNumberOfTables - 1; x++)
			{
				if (this.m_strDataSource[x,TABLESTATUS].Trim().ToUpper()=="F" &&
					this.m_strDataSource[x,FILESTATUS].Trim().ToUpper()=="F")
				{
					listbox1.Items.Add(this.m_strDataSource[x,TABLE].Trim());
				}
			}
		}
		
  		///<summary>
		///Return the location of the specified table within the strDataSource array.
		///-1 is returned if the strTableType is not found
		 ///</summary>
		/// <param name="strTableType">The unique id for the datasource table</param>
		public int getDataSourceTableNameRow(string pcTableType)
		{
			int x;
			for (x=0; x<= this.m_intNumberOfTables-1;x++)
			{
				if (pcTableType.Trim().ToUpper() == 
					this.m_strDataSource[x,TABLETYPE].Trim().ToUpper())
				{
					return x;
				}
			}
			return -1;
		}
		
		///<summary>
		//return a valid table name associated with the table type
		///</summary>
    /// <param name="strTableType">The unique id for the datasource table</param>		 
		public string getValidDataSourceTableName(string strTableType)
		{
			int x;
			for (x=0; x<= this.m_intNumberOfTables-1;x++)
			{
                string strFileStatus = this.m_strDataSource[x, FILESTATUS];
                if (strFileStatus != null)
                    strFileStatus = strFileStatus.Trim().ToUpper();
                string strTableStatus = this.m_strDataSource[x, TABLESTATUS];
                if (strTableStatus != null)
                    strTableStatus = strTableStatus.Trim().ToUpper();
                if (strTableType.Trim().ToUpper() == 
					this.m_strDataSource[x,TABLETYPE].Trim().ToUpper()
					&&
					strFileStatus =="F" 
					&&
					strTableStatus == "F")
				{
					return this.m_strDataSource[x,TABLE].Trim();
				}
			}
			return "";
		}
		/// <summary>
		/// get the full path to mdb file
		/// </summary>
		/// <param name="p_strTableType"></param>
		/// <returns></returns>
		public string getFullPathAndFile(string p_strTableType)
		{
			int x;
            macrosubst oMacroSub = new macrosubst();
            oMacroSub.ReferenceGeneralMacroSubstitutionVariableCollection = frmMain.g_oGeneralMacroSubstitutionVariable_Collection;
            for (x=0; x<= this.m_intNumberOfTables-1;x++)
			{
                if (p_strTableType.Trim().ToUpper() == 
					this.m_strDataSource[x,TABLETYPE].Trim().ToUpper())
				{
                    string strPathAndFile = oMacroSub.GeneralTranslateVariableSubstitution(this.m_strDataSource[x, PATH].ToString().Trim())
                        + "\\" + this.m_strDataSource[x, DBFILE].ToString().Trim();
                    return strPathAndFile;
				}
			}
			return "";

		}

	    public static void UpdateTableMacroVariable(string p_strTableType,string p_strTableName)
		{
			
			int x;
			
			for (x=0;x<=frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count-1;x++)
			{
				if (frmMain.g_oSQLMacroSubstitutionVariable_Collection.Item(x).Description.Trim().ToUpper()
					== p_strTableType.Trim().ToUpper() + " TABLE")
				{
					frmMain.g_oSQLMacroSubstitutionVariable_Collection.Item(x).SQLVariableSubstitutionString=p_strTableName.Trim();
					Datasource.g_oCurrentSQLMacroSubstitutionVariableItem.CopyProperties(frmMain.g_oSQLMacroSubstitutionVariable_Collection.Item(x),ref Datasource.g_oCurrentSQLMacroSubstitutionVariableItem);
					return;
				}
			}
			
			FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem oItem = new SQLMacroSubstitutionVariableItem();
			oItem.Description= p_strTableType.Trim() + " Table";
			oItem.Index=frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count;
			x=oItem.Index;
			

			switch (p_strTableType.Trim().ToUpper())
			{
				case "PLOT":
					oItem.VariableName="PlotTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName;
					break;
				case "CONDITION":
					oItem.VariableName="CondTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName;
					break;
				case "TREE":
					oItem.VariableName="TreeTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName;
					break;
				case "HARVEST COSTS":
					oItem.VariableName="HarvestCostsTable";
                    if (p_strTableName.Trim().Length == 0) p_strTableName = Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName;
					break;
                case "HARVEST METHODS":
					oItem.VariableName="HarvestMethodsTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=Tables.Reference.DefaultFiadbFVSVariantTableName;
					break;
				case "TREATMENT PRESCRIPTIONS":
					if (p_strTableName.Trim().Length == 0) p_strTableName= Tables.FVS.DefaultRxTableName;
					oItem.VariableName="RxTable";
					break;
                case "TREATMENT PACKAGES":
					if (p_strTableName.Trim().Length == 0) p_strTableName=Tables.FVS.DefaultRxPackageTableName;
					oItem.VariableName="RxPackageTable";
					break;
				case "TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS":
                    if (p_strTableName.Trim().Length == 0) p_strTableName = Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName;
					oItem.VariableName="TreeVolValBySpcGrpDiaGrpTable";
					break;
				case "FVS TREE LIST FOR PROCESSOR":
					oItem.VariableName="FvsTreeListForProcessorTable";
                    if (p_strTableName.Trim().Length == 0) p_strTableName = Tables.FVS.DefaultFVSTreeTableName;
					break;
				case "TRAVEL TIMES":
                    if (p_strTableName.Trim().Length == 0) p_strTableName = Tables.TravelTime.DefaultTravelTimeTableName;
					oItem.VariableName="TravelTimesTable";
					break;
				case "PROCESSING SITES":
					oItem.VariableName="PSitesTable";
                    if (p_strTableName.Trim().Length == 0) p_strTableName = Tables.TravelTime.DefaultProcessingSiteTableName;
					break;
                case "BIOSUM POP STRATUM ADJUSTMENT FACTORS":
                    oItem.VariableName="BiosumPopStratumAdjustmentFactorTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName;
					break;
                case "FIA TREE MACRO PLOT BREAKPOINT DIAMETER":
                    oItem.VariableName = "TreeMacroPlotBreakPointDiameterTable";
                    if (p_strTableName.Trim().Length == 0) p_strTableName=Tables.Reference.DefaultTreeMacroPlotBreakPointDiaTableName;
					break;
				case "SITE TREE":
					oItem.VariableName="SiteTreeTable";
					if (p_strTableName.Trim().Length == 0) p_strTableName=frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName;
					break;
			}
			if (oItem.VariableName.Trim().Length > 0)
			{
				oItem.SQLVariableSubstitutionString=p_strTableName.Trim();
				Datasource.g_oCurrentSQLMacroSubstitutionVariableItem.CopyProperties(oItem,ref Datasource.g_oCurrentSQLMacroSubstitutionVariableItem);
				frmMain.g_oSQLMacroSubstitutionVariable_Collection.Add(oItem);
			}
			else
			{
				Datasource.g_oCurrentSQLMacroSubstitutionVariableItem = new SQLMacroSubstitutionVariableItem();
			}

		}
		public static void InititializeMacroVariables()
		{
			for (int x=0;x<=Datasource.g_strProjectDatasourceTableTypesArray.Length - 1;x++)
			{
				Datasource.UpdateTableMacroVariable(Datasource.g_strProjectDatasourceTableTypesArray[x],"");
			}
			//
			//plot/condition connection
			//
			FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem oItem = new SQLMacroSubstitutionVariableItem();
			oItem.Description= "Key relationship between Plot and Condition Table";
			oItem.Index=frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count;
			oItem.VariableName="PlotCondConnection";
			oItem.SQLVariableSubstitutionString="@@CondTable@@.biosum_plot_id=@@PlotTable@@.biosum_plot_id";
			frmMain.g_oSQLMacroSubstitutionVariable_Collection.Add(oItem);
			//
			//condition/tree connection
			//
			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.Description= "Key relationship between Condition and Tree Table";
			oItem.Index=frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count;
			oItem.VariableName="CondTreeConnection";
			oItem.SQLVariableSubstitutionString="@@CondTable@@.biosum_cond_id=@@TreeTable@@.biosum_cond_id";
			frmMain.g_oSQLMacroSubstitutionVariable_Collection.Add(oItem);
			//
			//condition/tree connection
			//
			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.Description= "Key relationship between Condition and Tree Table";
			oItem.Index=frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count;
			oItem.VariableName="CondTreeConnection";
			oItem.SQLVariableSubstitutionString="@@CondTable@@.biosum_cond_id=@@TreeTable@@.biosum_cond_id";
			frmMain.g_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

            



		}

        /// <summary>
        /// get distinct FVS variants
        /// </summary>
        /// <param name="p_strVariantArray"></param>
        public string[] getVariants()
        {        
            //
            //get all the variants
            //
            int x;
            string[] strVariantArray = null;
			DataMgr dataMgr = new DataMgr();

			x = getDataSourceTableNameRow("PLOT");
			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dataMgr.GetConnectionString(getFullPathAndFile(Datasource.TableTypes.Plot))))
			{
				conn.Open();
				System.Collections.Generic.List<string> lstVariantArray = dataMgr.getStringList(conn, "SELECT DISTINCT FVS_VARIANT FROM " + m_strDataSource[x, Datasource.TABLE].Trim());
				strVariantArray = lstVariantArray.ToArray();
			}
			return strVariantArray;       
        }

		public void ValidateDataSources(
		   ref System.Collections.Generic.IDictionary<string, string[]> dictSources)
		{
			string strPathAndFile = "";
			string strConn = "";

			macrosubst oMacroSub = new macrosubst();
			oMacroSub.ReferenceGeneralMacroSubstitutionVariableCollection = frmMain.g_oGeneralMacroSubstitutionVariable_Collection;
			DataMgr dataMgr = new DataMgr();

			foreach (var tableType in dictSources.Keys)
            {
				string[] arrSource = dictSources[tableType];
				strPathAndFile = oMacroSub.GeneralTranslateVariableSubstitution(arrSource[PATH]) +
					"\\" + arrSource[DBFILE];

				if (System.IO.File.Exists(strPathAndFile) == true && System.IO.Path.GetExtension(strPathAndFile).Equals(".db"))
                {
					arrSource[FILESTATUS] = "F";

					strConn = dataMgr.GetConnectionString(strPathAndFile);
					using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                    {
						conn.Open();

						if (dataMgr.TableExist(conn, arrSource[TABLE]))
                        {
							arrSource[TABLESTATUS] = "F";
							dataMgr.m_strSQL = "SELECT COUNT(*) FROM " + arrSource[TABLE];
							long recordCount = dataMgr.getRecordCount(conn, dataMgr.m_strSQL, arrSource[TABLE]);
							arrSource[RECORDCOUNT] = Convert.ToString(recordCount);
						}
					}
				}
                else
                {
					arrSource[FILESTATUS] = "NF";
					arrSource[TABLESTATUS] = "NF";
					arrSource[RECORDCOUNT] = "0";
				}
			}
		}

        public void UpdateDataSourcePath(string strTableType, string strPath, string strFile, string strTableName)
        {
            ado_data_access oAdo = new ado_data_access();
            string strConn = oAdo.getMDBConnString(this.m_strDataSourceDBFile, "", "");
            if (this.m_strScenarioId.Trim().Length > 0)
            {
                oAdo.m_strSQL = "UPDATE " + this.m_strDataSourceTableName + " SET path = '" + strPath + "', " +
                    "file = '" + strFile + "', " +
                    "table_name = '" + strTableName + "'" +
                    " WHERE scenario_id = '" + this.m_strScenarioId + "' AND " +
                    "table_type = '" + strTableType + "';";
            }
            else
            {
                oAdo.m_strSQL = "UPDATE " + this.m_strDataSourceTableName + " SET path = '" + strPath + "', " +
                    "file = '" + strFile + "', " +
                    "table_name = '" + strTableName + "'" +
                    " WHERE table_type = '" + strTableType + "';";
            }
            using (var oConn = new System.Data.OleDb.OleDbConnection(strConn))
            {
                oConn.Open();
                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
            }
        }
		public bool LoadTableColumnNamesAndDataTypes
		{
			get {return this._bLoadFieldNamesAndDatatypes;}
			set {this._bLoadFieldNamesAndDatatypes=value;}
		}
		public bool LoadTableRecordCount
		{
			get {return this._bLoadTableRecordCount;}
			set {this._bLoadTableRecordCount=value;}

		}

		public System.Collections.Generic.List<string> getDataSourceDbsList()
        {
			int x;
			macrosubst oMacroSub = new macrosubst();
			oMacroSub.ReferenceGeneralMacroSubstitutionVariableCollection = frmMain.g_oGeneralMacroSubstitutionVariable_Collection;

			System.Collections.Generic.List<string> lstDataSourcePaths = new System.Collections.Generic.List<string>();

			for (x = 0; x <= m_intNumberOfTables - 1; x++)
            {
				string strFileStatus = this.m_strDataSource[x, FILESTATUS];
				if (strFileStatus != null)
                {
					strFileStatus = strFileStatus.Trim().ToUpper();
				}
				string strTableStatus = this.m_strDataSource[x, TABLESTATUS];
				if (strTableStatus != null)
                {
					strTableStatus = strTableStatus.Trim().ToUpper();
				}

				string strPath = oMacroSub.GeneralTranslateVariableSubstitution(this.m_strDataSource[x, PATH].Trim()) + "\\" +
								 this.m_strDataSource[x, DBFILE].Trim();
				if (!lstDataSourcePaths.Contains(strPath) && strFileStatus == "F" && strTableStatus == "F")
                {
					lstDataSourcePaths.Add(strPath);
                }
			}

			return lstDataSourcePaths;
		}

        public class TableTypes
        {
            static public string FvsTreeSpecies = "FVS Tree Species";
            static public string FiaTreeSpeciesReference = "FIA Tree Species Reference";
            static public string HarvestMethods = "Harvest Methods";
            static public string FVSVariant = "FIADB FVS Variant";
            static public string TravelTimes = "Travel Times";
            static public string ProcessingSites = "Processing Sites";
            static public string Plot = "Plot";
            static public string Condition = "Condition";
            static public string Tree = "Tree";
			static public string SiteTree = "Site Tree";
            static public string SeqNumDefinitions = "FVS PRE-POST SeqNum Definitions";
            static public string SeqNumRxPackageAssign = "FVS PRE-POST SeqNum Treatment Package Assign";
            static public string RxPackage = "Treatment Packages";
            static public string Rx = "Treatment Prescriptions";
            static public string RxHarvestCostColumns = "Treatment Prescriptions Harvest Cost Columns";
			static public string TreeSpecies = "Tree Species";
			static public string FiaTreeMacroPlotBreakpointDia = "FIA Tree Macro Plot Breakpoint Diameter";
			static public string OwnerGroups = "Owner Groups";
			static public string PopStratumAdjFactors = "BIOSUM Pop Stratum Adjustment Factors";

		}
		
	}

}

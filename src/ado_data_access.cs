using System;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.IO;
namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for ado_data_access.
	/// </summary>
	public class ado_data_access
	{
		public System.Data.DataSet m_DataSet;
		public System.Data.OleDb.OleDbDataAdapter m_OleDbDataAdapter;
		public System.Data.OleDb.OleDbCommand m_OleDbCommand;
		public System.Data.OleDb.OleDbConnection m_OleDbConnection;
		public System.Data.OleDb.OleDbDataReader m_OleDbDataReader;
		public System.Data.DataTable m_DataTable;
		public System.Data.OleDb.OleDbTransaction m_OleDbTransaction;
		public string m_strError;
		public int m_intError;
		public string m_strSQL;
		public string m_strTable;
        private bool _bDisplayErrors = true;
        private int m_intAttempts = 1;




		public ado_data_access()
		{
			//
			// TODO: Add constructor logic here
			//
		}
		~ado_data_access()
		{
            if (m_OleDbConnection != null)
            {
                if (m_OleDbConnection.State == System.Data.ConnectionState.Open) CloseConnection(m_OleDbConnection);
            }
		}
        /// <summary>
        /// Open a connection to an MS Access database file
        /// </summary>
        /// <param name="strConn"></param>
		public void OpenConnection(string strConn)
		{
			System.Data.OleDb.OleDbConnection p_OleDbConnection = new System.Data.OleDb.OleDbConnection();

			p_OleDbConnection.ConnectionString = strConn;
			
			try
			{
				p_OleDbConnection.Open();
				
			}
			catch (Exception caught)
			{
				this.m_strError = caught.Message;
				this.m_intError = -1;
                if (_bDisplayErrors)
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:OpenConnection  \n" + 
					"Err Msg - " + this.m_strError,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}
			this.m_OleDbConnection = p_OleDbConnection;
		}
        /// <summary>
        /// Open a connection to an MS Access database file
        /// </summary>
        /// <param name="strConn"></param>
        public void OpenConnection(string strConn, int p_intAttempts)
        {
            
            int x;
            bool bDisplayErrors=DisplayErrors;

            System.Data.OleDb.OleDbConnection p_OleDbConnection = new System.Data.OleDb.OleDbConnection();

           

            DisplayErrors = false;

            for (x = 1; x <= p_intAttempts; x++)
            {
                OpenConnection(strConn, ref p_OleDbConnection);
                if (m_intError == 0)     break;
                System.Threading.Thread.Sleep(5000);
                

            }

            
            if (m_intError == 0) m_OleDbConnection = p_OleDbConnection;
            else if (bDisplayErrors)
                MessageBox.Show("!!Error!! \n" +
                    "Module - ado_data_access:OpenConnection  \n" +
                    "Err Msg - " + this.m_strError,
                    "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
              
           
        }
        /// <summary>
        /// Open a connection to an MS Access database file
        /// </summary>
        /// <param name="strConn"></param>
        /// <param name="p_OleDbConnection"></param>
		public void OpenConnection(string strConn, ref System.Data.OleDb.OleDbConnection p_OleDbConnection)
		{
            this.m_intError=0;
			this.m_strError="";
			try
			{
			    p_OleDbConnection.ConnectionString = strConn;
			
			
				p_OleDbConnection.Open();
				
			}
			catch (Exception caught)
			{
				this.m_strError = caught.Message;
				this.m_intError = -1;
                if (_bDisplayErrors)
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:OpenConnection  \n" + 
					"Err Msg - " + this.m_strError,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
			}
			
		}
        public void CloseAndDisposeConnection(System.Data.OleDb.OleDbConnection p_Connection, bool p_bClearPool)
        {
            try
            {

                if (p_Connection.State != ConnectionState.Closed)
                {
                    if (m_OleDbDataReader != null) m_OleDbDataReader.Dispose();

                    if (m_OleDbCommand != null) m_OleDbCommand.Dispose();

                    if (m_OleDbDataAdapter != null)
                    {
                        if (m_OleDbDataAdapter.SelectCommand != null)
                        {
                            m_OleDbDataAdapter.SelectCommand.Dispose();
                        }
                        if (m_OleDbDataAdapter.UpdateCommand != null)
                        {
                            m_OleDbDataAdapter.UpdateCommand.Dispose();
                        }
                        if (m_OleDbDataAdapter.DeleteCommand != null)
                        {
                            m_OleDbDataAdapter.DeleteCommand.Dispose();
                        }
                        if (m_OleDbDataAdapter.InsertCommand != null)
                        {
                            m_OleDbDataAdapter.InsertCommand.Dispose();
                        }
                    }

                }
                while (p_Connection.State != System.Data.ConnectionState.Closed)
                {

                    p_Connection.Close();
                    System.Threading.Thread.Sleep(1000);



                }
                if (p_Connection.State == ConnectionState.Closed)
                {
                    if (p_bClearPool) System.Data.OleDb.OleDbConnection.ReleaseObjectPool();
                    p_Connection.Dispose();
                    p_Connection = null;

                }
            }
            catch (Exception e)
            {
                m_strError = e.Message;
                m_intError = -1;
            }
        }

        /// <summary>
        /// Close an MS Access Database file connection
        /// </summary>
        /// <param name="p_oConn"></param>
		public void CloseConnection(System.Data.OleDb.OleDbConnection p_oConn)
		{
			try
			{
				while (p_oConn.State != System.Data.ConnectionState.Closed)
				{
					p_oConn.Close();
					System.Threading.Thread.Sleep(1000);
				}
			}
			catch (Exception e)
			{
			}
		}
        /// <summary>
        /// Execute a SQL Query (Update, Insert, Delete, Into)
        /// </summary>
        /// <param name="strConn"></param>
        /// <param name="strSQL"></param>
        public void SqlNonQuery(string strConn, string strSQL)
        {
            using (var p_OleDbConnection = new OleDbConnection(strConn))
            using (var p_OleDbCommand = new OleDbCommand(strSQL, p_OleDbConnection))
            {
                try
                {
                    p_OleDbCommand.Connection.Open();
                    p_OleDbCommand.ExecuteNonQuery();
                }
                catch (Exception caught)
                {
                    this.m_strError = caught.Message + " The SQL command " + strSQL + " Failed";
                    this.m_intError = -1;
                    if (_bDisplayErrors)
                        MessageBox.Show("!!Error!! \n" +
                                        "Module - ado_data_access:SqlNonQuery  \n" +
                                        "Err Msg - " + this.m_strError,
                            "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                }
            }
        }


        /// <summary>
        /// Execute a SQL Query (Update, Insert, Delete, Into)
        /// </summary>
        /// <param name="p_OleDbConnection"></param>
        /// <param name="strSQL"></param>
        public void SqlNonQuery(System.Data.OleDb.OleDbConnection p_OleDbConnection, string strSQL)
        {
            using (var p_OleDbCommand = new OleDbCommand(strSQL, p_OleDbConnection))
            {
                try
                {
                    p_OleDbCommand.ExecuteNonQuery();
                }
                catch (Exception caught)
                {
                    this.m_strError = caught.Message + " The SQL command " + strSQL + " Failed";
                    this.m_intError = -1;
                    if (_bDisplayErrors)
                        MessageBox.Show("!!Error!! \n" +
                                        "Module - ado_data_access:SqlNonQuery  \n" +
                                        "Err Msg - " + this.m_strError,
                            "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                }
            }
        }


        /// <summary>
        /// Execute a SQL Query (Update, Insert, Delete, Into) to a connection with an active transaction command
        /// </summary>
        /// <param name="p_OleDbConnection"></param>
        /// <param name="p_OleDbTransaction"></param>
        /// <param name="strSQL"></param>
		public void SqlNonQuery(System.Data.OleDb.OleDbConnection p_OleDbConnection,System.Data.OleDb.OleDbTransaction p_OleDbTransaction, string strSQL)
		{
		    using (var p_OleDbCommand = new OleDbCommand(strSQL, p_OleDbConnection, p_OleDbTransaction))
		    {
		        try
		        {
		            p_OleDbCommand.ExecuteNonQuery();
		        }
		        catch (Exception caught)
		        {
		            this.m_strError = caught.Message + " The SQL command " + strSQL + " Failed";
		            this.m_intError = -1;
		            if (_bDisplayErrors)
		                MessageBox.Show("!!Error!! \n" +
		                                "Module - ado_data_access:SqlNonQuery  \n" +
		                                "Err Msg - " + this.m_strError,
		                    "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
		                    System.Windows.Forms.MessageBoxIcon.Exclamation);
		        }
		    }
		}
        /// <summary>
        /// Load oledb data reader with records returned from SQL 
        /// </summary>
        /// <param name="strConn"></param>
        /// <param name="strSql"></param>
		public void SqlQueryReader(string strConn,string strSql)
		{
			this.m_intError=0;
			this.m_strError="";
            this.OpenConnection(strConn);
			if (this.m_intError==0)
			{
			    this.m_OleDbCommand = this.m_OleDbConnection.CreateCommand();
				this.m_OleDbCommand.CommandText = strSql;
				try
				{
					this.m_OleDbDataReader = this.m_OleDbCommand.ExecuteReader();
				}
				catch (Exception caught)
				{
					this.m_intError = -1;
					this.m_strError = caught.Message + " The Query Command " + this.m_OleDbCommand.CommandText.ToString() + " Failed";

                    if (_bDisplayErrors)
					MessageBox.Show("!!Error!! \n" + 
						"Module - ado_data_access:SqlQueryReader  \n" + 
						"Err Msg - " + this.m_strError,
						"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);
					this.m_OleDbDataReader = null;
                    this.m_OleDbCommand = null;
					this.m_OleDbConnection.Close();
					this.m_OleDbConnection = null;
					return;
				}
			}
		}
        /// <summary>
        /// Load oledb data reader with records returned from SQL 
        /// </summary>
        /// <param name="p_OleDbConnection"></param>
        /// <param name="strSql"></param>
		public void SqlQueryReader(System.Data.OleDb.OleDbConnection p_OleDbConnection,string strSql)
		{
			this.m_intError=0;
			this.m_strError="";
			
				this.m_OleDbCommand = p_OleDbConnection.CreateCommand();
				this.m_OleDbCommand.CommandText = strSql;
				try
				{
					this.m_OleDbDataReader = this.m_OleDbCommand.ExecuteReader();
				}
				catch (Exception caught)
				{
					this.m_intError = -1;
					this.m_strError = caught.Message + " The Query Command " + this.m_OleDbCommand.CommandText.ToString() + " Failed";
                    if (_bDisplayErrors)
					MessageBox.Show("!!Error!! \n" + 
						"Module - ado_data_access:SqlQueryReader  \n" + 
						"Err Msg - " + this.m_strError,
						"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);
					this.m_OleDbDataReader = null;
					this.m_OleDbCommand = null;
					return;
				}
			    
			
		}
		/// <summary>
        /// Load oledb data reader with records returned from SQL that have a db connection with an active transaction 
		/// </summary>
		/// <param name="p_OleDbConnection"></param>
		/// <param name="p_trans"></param>
		/// <param name="strSql"></param>
		public void SqlQueryReader(System.Data.OleDb.OleDbConnection p_OleDbConnection,System.Data.OleDb.OleDbTransaction p_trans,string strSql)
		{
			this.m_intError=0;
			this.m_strError="";
			
			this.m_OleDbCommand = p_OleDbConnection.CreateCommand();
			this.m_OleDbCommand.CommandText = strSql;
            this.m_OleDbCommand.Transaction = p_trans;
			try
			{
				this.m_OleDbDataReader = this.m_OleDbCommand.ExecuteReader();
			}
			catch (Exception caught)
			{
				this.m_intError = -1;
				this.m_strError = caught.Message + " The Query Command " + this.m_OleDbCommand.CommandText.ToString() + " Failed";
                if (_bDisplayErrors)
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:SqlQueryReader  \n" + 
					"Err Msg - " + this.m_strError,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_OleDbDataReader = null;
				this.m_OleDbCommand = null;
				return;
			}
			    
			
		}
        /// <summary>
        /// Returns a ADO.NET datatable with the table schemas contained in the MS Access Database file
        /// </summary>
        /// <param name="p_OleDbConnection"></param>
        /// <param name="strSql"></param>
        /// <returns></returns>
		public System.Data.DataTable getTableSchema(System.Data.OleDb.OleDbConnection p_OleDbConnection,string strSql)
		{
			System.Data.DataTable p_dt;
			this.m_intError=0;
			this.m_strError="";
			
			this.m_OleDbCommand = p_OleDbConnection.CreateCommand();
			this.m_OleDbCommand.CommandText = strSql;
			try
			{
				this.m_OleDbDataReader = this.m_OleDbCommand.ExecuteReader(CommandBehavior.KeyInfo);
				p_dt = this.m_OleDbDataReader.GetSchemaTable();
			}
			catch (Exception caught)
			{
				this.m_intError = -1;
				this.m_strError = caught.Message + " The Query Command " + this.m_OleDbCommand.CommandText.ToString() + " Failed";
                if (_bDisplayErrors)
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:getTableSchema  \n" + 
					"Err Msg - " + this.m_strError,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_OleDbDataReader = null;
				this.m_OleDbCommand = null;
				return null;
			}
			this.m_OleDbDataReader.Close();
			return p_dt;
			
		}
        /// <summary>
        /// Returns a ADO.NET datatable with the table schemas contained in the MS Access Database file
        /// </summary>
        /// <param name="p_OleDbConnection"></param>
        /// <param name="p_trans"></param>
        /// <param name="strSql"></param>
        /// <returns></returns>
		public System.Data.DataTable getTableSchema(System.Data.OleDb.OleDbConnection p_OleDbConnection,
			                                        System.Data.OleDb.OleDbTransaction p_trans,
			                                        string strSql)
		{
			System.Data.DataTable p_dt;
			this.m_intError=0;
			this.m_strError="";
			
			this.m_OleDbCommand = p_OleDbConnection.CreateCommand();
			this.m_OleDbCommand.CommandText = strSql;
			this.m_OleDbCommand.Transaction = p_trans;
			try
			{
				this.m_OleDbDataReader = this.m_OleDbCommand.ExecuteReader(CommandBehavior.KeyInfo);
				p_dt = this.m_OleDbDataReader.GetSchemaTable();
			}
			catch (Exception caught)
			{
				this.m_intError = -1;
				this.m_strError = caught.Message + " The Query Command " + this.m_OleDbCommand.CommandText.ToString() + " Failed";
                if (_bDisplayErrors)
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:getTableSchema  \n" + 
					"Err Msg - " + this.m_strError,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_OleDbDataReader = null;
				this.m_OleDbCommand = null;
				return null;
			}
			this.m_OleDbDataReader.Close();
			return p_dt;
			
		}
        /// <summary>
        /// Returns a ADO.NET datatable with the table schemas contained in the MS Access Database file
        /// </summary>
        /// <param name="strConn"></param>
        /// <param name="strSql"></param>
        /// <returns></returns>
		public System.Data.DataTable getTableSchema(string strConn,string strSql)
		{
			System.Data.DataTable p_dt;
			this.m_intError=0;
			this.m_strError="";
			
			this.OpenConnection(strConn);
			if (this.m_intError==0)
			{
				this.m_OleDbCommand = this.m_OleDbConnection.CreateCommand();
				this.m_OleDbCommand.CommandText = strSql;
				try
				{
					this.m_OleDbDataReader = this.m_OleDbCommand.ExecuteReader(CommandBehavior.KeyInfo);
					p_dt = this.m_OleDbDataReader.GetSchemaTable();
				}
				catch (Exception caught)
				{
					this.m_intError = -1;
					this.m_strError = caught.Message + " The Query Command " + this.m_OleDbCommand.CommandText.ToString() + " Failed";
                    if (_bDisplayErrors)
					MessageBox.Show("!!Error!! \n" + 
						"Module - ado_data_access:getTableSchema  \n" + 
						"Err Msg - " + this.m_strError,
						"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);
					this.m_OleDbDataReader = null;
					this.m_OleDbCommand = null;
					return null;
				}
				this.m_OleDbDataReader.Close();
				return p_dt;
			}
			return null;
			
			
		}
        /// <summary>
        /// Get the table names contained in the MS Access database file connection
        /// </summary>
        /// <param name="p_oConn"></param>
        /// <returns></returns>
        public string[] getTableNames(System.Data.OleDb.OleDbConnection p_oConn)
        {
            string strDelimiter = ",";
            string strTables = "";
            DataTable tables = p_oConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables_Info,
                new object[] {null, null, null, null});
            int x = 0;
            foreach (DataRow row in tables.Rows)
            {
                if (row["table_name"] != System.DBNull.Value)
                {
                    if (Convert.ToString(row["table_name"]).ToUpper().IndexOf("MSYS") < 0)
                    {
                        if (Convert.ToString(row["table_type"]).Trim().ToUpper() == "TABLE" ||
                            Convert.ToString(row["table_type"]).Trim().ToUpper() == "LINK" ||
                            Convert.ToString(row["table_type"]).Trim().ToUpper() == "PASS-THROUGH")
                        {
                            strTables = strTables + Convert.ToString(row["table_name"]) + ",";
                            x++;
                        }
                    }
                }
            }
            if (strTables.Trim().Length > 0) strTables = strTables.Substring(0, strTables.Length - 1);

            string[] strTablesArray = strTables.Split(strDelimiter.ToCharArray());

            return strTablesArray;
        }


        /// <summary>
        /// Get the table names (excluding linked tables) contained in the MS Access database file connection
        /// </summary>
        /// <param name="p_oConn"></param>
        /// <returns></returns>
        public string[] getTableNamesOfSpecificTypes(System.Data.OleDb.OleDbConnection p_oConn, bool includeRegularTables = true, bool includeLinkedTables = false, bool includePassThroughTables = false)
        {
            string strDelimiter = ",";
            string strTables = "";
            DataTable tables = p_oConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables_Info,
                new object[] {null, null, null, null});
            int x = 0;
            foreach (DataRow row in tables.Rows)
            {
                if (row["table_name"] != System.DBNull.Value)
                {
                    if (Convert.ToString(row["table_name"]).ToUpper().IndexOf("MSYS") < 0)
                    {
                        if (includeRegularTables && Convert.ToString(row["table_type"]).Trim().ToUpper() ==
                            "TABLE")
                        {
                            strTables = strTables + Convert.ToString(row["table_name"]) + ",";
                            x++;
                        }
                        else if (includeLinkedTables && Convert.ToString(row["table_type"]).Trim().ToUpper() ==
                            "LINK")
                        {
                            strTables = strTables + Convert.ToString(row["table_name"]) + ",";
                            x++;
                        }
                        else if (includePassThroughTables && Convert.ToString(row["table_type"]).Trim().ToUpper() ==
                            "PASS-THROUGH")
                        {
                            strTables = strTables + Convert.ToString(row["table_name"]) + ",";
                            x++;
                        }
                    }
                }
            }
            if (strTables.Trim().Length > 0) strTables = strTables.Substring(0, strTables.Length - 1);
            string[] strTablesArray = strTables.Split(strDelimiter.ToCharArray());
            return strTablesArray;
        }


        /// <summary>
        /// Check if the table exists in the MS Access database file
        /// </summary>
        /// <param name="p_oConn"></param>
        /// <param name="p_strTableName"></param>
        /// <returns></returns>
        public bool TableExist(System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName)
        {
            DataTable tables = p_oConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables_Info,
                new object[] {null, null, null, null});
            bool bFound = false;
            foreach (DataRow row in tables.Rows)
            {
                if (row["table_name"] != System.DBNull.Value)
                {
                    if (Convert.ToString(row["table_name"]).ToUpper().IndexOf("MSYS") < 0)
                    {
                        if (Convert.ToString(row["table_type"]).Trim().ToUpper() == "TABLE" ||
                            Convert.ToString(row["table_type"]).Trim().ToUpper() == "LINK" ||
                            Convert.ToString(row["table_type"]).Trim().ToUpper() == "PASS-THROUGH")
                        {
                            if (Convert.ToString(row["table_name"]).Trim().ToUpper() ==
                                p_strTableName.Trim().ToUpper())
                            {
                                bFound = true;
                                break;
                            }
                        }
                    }
                }
            }
            tables.Dispose();
            return bFound;
        }
        /// <summary>
        /// Check if the table exists in the MS Access database file
        /// </summary>
        /// <param name="p_strConn"></param>
        /// <param name="p_strTableName"></param>
        /// <returns></returns>
		public bool TableExist(string p_strConn,string p_strTableName)
		{
			bool bFound=false;
			this.OpenConnection(p_strConn);
			if (m_intError==0)
			{
				bFound = TableExist(this.m_OleDbConnection,p_strTableName);
				this.CloseConnection(this.m_OleDbConnection);
			}
			return bFound;
		}
        /// <summary>
        /// Check if the column exists in the designated table
        /// </summary>
        /// <param name="p_oConn"></param>
        /// <param name="p_strTableName"></param>
        /// <param name="p_strColumnName"></param>
        /// <returns></returns>
		public bool ColumnExist(System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName,string p_strColumnName)
		{
			bool bFound=false;
			if (TableExist(p_oConn,p_strTableName))
			{
				string[] strArray = this.getFieldNamesArray(p_oConn,"SELECT * FROM " + p_strTableName);
				if (strArray != null)
				{
					for (int x=0;x<=strArray.Length - 1;x++)
					{
						if (p_strColumnName.Trim().ToUpper() == strArray[x].Trim().ToUpper())
						{
							bFound=true;
							break;
						}
					}
				}
			}
			return bFound;
		}
        /// <summary>
        /// Return a comma-delimited string of all the columns returned from executing the SELECT sql command
        /// </summary>
        /// <param name="p_oConn"></param>
        /// <param name="p_strSql"></param>
        /// <returns></returns>
		public string getFieldNames(System.Data.OleDb.OleDbConnection p_oConn,string p_strSql)
		{
			this.m_intError=0;
			System.Data.DataTable oTableSchema = this.getTableSchema(p_oConn,p_strSql);
			if (this.m_intError !=0) return "";
			string strFields="";
			
			for (int x=0; x<=oTableSchema.Rows.Count-1;x++)
			{
				strFields = strFields + oTableSchema.Rows[x]["columnname"].ToString().Trim() + ",";
			}
			if (strFields.Trim().Length > 0) strFields=strFields.Substring(0,strFields.Trim().Length -1);

			return strFields;
			
		}
        /// <summary>
        /// Return a comma-delimited string of all the columns returned from executing the SELECT sql command
        /// </summary>
        /// <param name="p_strConn"></param>
        /// <param name="p_strSql"></param>
        /// <returns></returns>
		public string getFieldNames(string p_strConn,string p_strSql)
		{
			string strFields="";
			this.m_intError=0;
			this.OpenConnection(p_strConn);
			if (this.m_intError==0)
			{
				System.Data.DataTable oTableSchema = this.getTableSchema(this.m_OleDbConnection,p_strSql);
				if (this.m_intError !=0) 
				{
					this.CloseConnection(this.m_OleDbConnection);
					return "";
				}
				
			
				for (int x=0; x<=oTableSchema.Rows.Count-1;x++)
				{
					strFields = strFields + oTableSchema.Rows[x]["columnname"].ToString().Trim() + ",";
				}
				if (strFields.Trim().Length > 0) strFields=strFields.Substring(0,strFields.Trim().Length -1);
				
				
			}
			this.CloseConnection(this.m_OleDbConnection);
			return strFields;
			
		}
        /// <summary>
        /// obtain comma-delimited strings of field names and their data types from executing the SELECT SQL
        /// </summary>
        /// <param name="p_strConn"></param>
        /// <param name="p_strSql"></param>
        /// <param name="p_strFieldNamesList"></param>
        /// <param name="p_strDataTypesList"></param>
		public void getFieldNamesAndDataTypes(string p_strConn,string p_strSql,ref string p_strFieldNamesList,ref string p_strDataTypesList)
		{
			string strFields="";
			string strDataTypes="";
			this.m_intError=0;
			this.OpenConnection(p_strConn);
			if (this.m_intError==0)
			{
				System.Data.DataTable oTableSchema = this.getTableSchema(this.m_OleDbConnection,p_strSql);
				if (this.m_intError !=0) 
				{
					this.CloseConnection(this.m_OleDbConnection);
					return ;
				}
				
			
				for (int x=0; x<=oTableSchema.Rows.Count-1;x++)
				{
					strFields = strFields + oTableSchema.Rows[x]["columnname"].ToString().Trim() + ",";
					strDataTypes = strDataTypes + oTableSchema.Rows[x]["datatype"].ToString().Trim() + ",";
				}
				if (strFields.Trim().Length > 0) strFields=strFields.Substring(0,strFields.Trim().Length -1);
				if (strDataTypes.Trim().Length > 0) strDataTypes=strDataTypes.Substring(0,strDataTypes.Trim().Length -1);
			}
			p_strFieldNamesList = p_strFieldNamesList + strFields;
			p_strDataTypesList = p_strDataTypesList + strDataTypes;
			this.CloseConnection(this.m_OleDbConnection);


		}
        /// <summary>
        /// obtain comma-delimited strings of field names and their data types from executing the SELECT SQL
        /// </summary>
        /// <param name="p_oConn"></param>
        /// <param name="p_strSql"></param>
        /// <param name="p_strFieldNamesList"></param>
        /// <param name="p_strDataTypesList"></param>
		public void getFieldNamesAndDataTypes(System.Data.OleDb.OleDbConnection p_oConn,string p_strSql,ref string p_strFieldNamesList,ref string p_strDataTypesList)
		{
			string strFields="";
			string strDataTypes="";
			this.m_intError=0;
			
			if (this.m_intError==0)
			{
				System.Data.DataTable oTableSchema = this.getTableSchema(p_oConn,p_strSql);
				if (this.m_intError !=0) return ;
				
			
				for (int x=0; x<=oTableSchema.Rows.Count-1;x++)
				{
					strFields = strFields + oTableSchema.Rows[x]["columnname"].ToString().Trim() + ",";
					strDataTypes = strDataTypes + oTableSchema.Rows[x]["datatype"].ToString().Trim() + ",";
				}
				if (strFields.Trim().Length > 0) strFields=strFields.Substring(0,strFields.Trim().Length -1);
				if (strDataTypes.Trim().Length > 0) strDataTypes=strDataTypes.Substring(0,strDataTypes.Trim().Length -1);
			}
			p_strFieldNamesList = p_strFieldNamesList + strFields;
			p_strDataTypesList = p_strDataTypesList + strDataTypes;


		}

		
        /// <summary>
        /// Return an array of field names after executing the SELECT SQL 
        /// </summary>
        /// <param name="p_oConn"></param>
        /// <param name="p_strSql"></param>
        /// <returns></returns>
		public string[] getFieldNamesArray(System.Data.OleDb.OleDbConnection p_oConn,string p_strSql)
		{
			this.m_intError=0;
			string strList = getFieldNames(p_oConn,p_strSql);
			if (strList.Trim().Length==0) return null;

			string strDelimiter=",";
			string[] strArray = strList.Split(strDelimiter.ToCharArray());
			return strArray;

			
		}




		/****
		 **** format strings to be used in an sql statement
		 ****/
		public string FixString(string SourceString , string StringToReplace, string StringReplacement)
		{
			SourceString = SourceString.Replace(StringToReplace, StringReplacement);
			return(SourceString);
		}
        /// <summary>
        /// returns Y or N for whether the field is a string or not. 
        /// </summary>
        /// <param name="strFieldType"></param>
        /// <returns></returns>
		public string getIsTheFieldAStringDataType(string strFieldType)
		{
			switch (strFieldType.Trim())
			{
				case "System.Single":
					return "N";
				case "System.Double":
					return "N";
				case "System.Decimal":
					return "N";
				case "System.String":
					return "Y";
				case "System.Int16":
					return "N";
				case "System.Char":
					return "Y";
				case "System.Int32":
					return "N";
				case "System.DateTime":
					return "Y";
				case "System.DayOfWeek":
					return "Y";
				case "System.Int64":
					return "N";
				case "System.Byte":
					return "N";
				case "System.Boolean":
					return "N";
				default:
					//return "N";
					MessageBox.Show(strFieldType + " is undefined");
					return "N";
			}


		}
		/// <summary>
		/// Return a formatted string used to compile a CREATE TABLE sql command
		/// </summary>
		/// <param name="p_oRow">ADO.NET datarow</param>
		/// <returns></returns>
		public string FormatCreateTableSqlFieldItem(System.Data.DataRow p_oRow)
		{
			string strLine;
			string strColumn = p_oRow["ColumnName"].ToString().Trim();
			string strDataType = p_oRow["DataType"].ToString().Trim();
			string strPrecision = "";
			string strScale = "";
			string strSize = "";

			if (p_oRow["ColumnSize"] != null)
				strSize = Convert.ToString(p_oRow["ColumnSize"]);

			if (p_oRow["NumericPrecision"] != null)
				strPrecision = Convert.ToString(p_oRow["NumericPrecision"]);

			if (p_oRow["NumericScale"] != null)
				strScale = Convert.ToString(p_oRow["NumericScale"]);


			
            strColumn = "[" + strColumn + "]";

            
			switch (strDataType)
			{

				case "System.Single":
					strDataType = "single";
					break;
				case "System.Double":
					strDataType = "double";
					break;
				case "System.Decimal":
					strDataType = "decimal";
					break;
				case "System.String":
					strDataType = "text";
					break;
				case "System.Int16":
					strDataType = "short";
					break;
				case "System.Char":
					strDataType = "text";
					break;
				case "System.Int32":
					strDataType = "integer";
					break;
				case "System.DateTime":
					strDataType = "datetime";
					break;
				case "System.DayOfWeek":
					break;
				case "System.Int64":
					break;
				case "System.Byte":
					strDataType = "byte";
					break;
				case "System.Boolean":
					break;



			}

			strLine = strColumn + " " + strDataType;

			if (strSize.Trim().Length > 0 && strDataType == "text")
				if (Convert.ToInt32(strSize) < 256)
					strLine = strLine + " (" + strSize + ")";
				else
				{
					strLine = strColumn + " memo";
				}
			else
			{
				if (strDataType == "decimal")
				{
					if (strPrecision.Trim() == "0")
						strLine = strColumn + " double";
					else 
						strLine = strLine + " (" + strPrecision + "," + strScale + ")";
				}
                
                    
			}
			return strLine;

		}
        /// <summary>
        /// Format SQL command string for these issues:
        /// 1. numeric values to handle MS ACCESS decimal length maximums
        /// 2. MS ACCESS reserved words
        /// </summary>
        /// <param name="p_strColumnName"></param>
        /// <param name="p_strDataType"></param>
        /// <param name="p_strColumnSize"></param>
        /// <param name="p_strNumericPrecision"></param>
        /// <param name="p_strNumericScale"></param>
        /// <param name="p_bReservedWordFormatting"></param>
        /// <returns></returns>
		public string FormatSelectSqlFieldItem(string p_strColumnName,string p_strDataType,
			string p_strColumnSize, string p_strNumericPrecision,
			string p_strNumericScale,bool p_bReservedWordFormatting)
		{
			string strLine="";
			string strPrecision = "";
			string strScale = "";
			string strSize = "";

			if (p_strColumnSize != null && p_strColumnSize.Trim().Length > 0)
				strSize = p_strColumnSize;

			if (p_strNumericPrecision != null && p_strNumericPrecision.Trim().Length > 0)
				strPrecision = p_strNumericPrecision;

			if (p_strNumericScale != null && p_strNumericScale.Trim().Length > 0)
				strScale = p_strNumericScale;

			if (p_bReservedWordFormatting)
			{
				p_strColumnName=FormatReservedWordColumnName(p_strColumnName);
			}




			switch (p_strDataType)
			{

				case "System.Single":
					p_strDataType = "single";
					break;
				case "System.Double":
					p_strDataType = "double";
					break;
				case "System.Decimal":
					p_strDataType = "decimal";
					break;
				case "System.String":
					p_strDataType = "text";
					break;
				case "System.Int16":
					p_strDataType = "short";
					break;
				case "System.Char":
					p_strDataType = "text";
					break;
				case "System.Int32":
					p_strDataType = "integer";
					break;
				case "System.DateTime":
					p_strDataType = "datetime";
					break;
				case "System.DayOfWeek":
					break;
				case "System.Int64":
					break;
				case "System.Byte":
					break;
				case "System.Boolean":
					break;



			}

			strLine = p_strColumnName;

			if (p_strDataType == "decimal")
			{
				if (strPrecision.Trim() == "0")
					strLine = "ROUND(" + p_strColumnName + ",14) AS " + p_strColumnName;
			}
			else if (p_strDataType == "double")
			{
				strLine = "ROUND(" + p_strColumnName + ",14) AS " + p_strColumnName;
			}
          

			return strLine;
		}
		public string FormatSelectSqlFieldItem(System.Data.DataRow p_oRow,bool p_bReservedWordFormatting)
		{
			string strLine="";
			string strColumn = p_oRow["ColumnName"].ToString().Trim();
			string strDataType = p_oRow["DataType"].ToString().Trim();
			string strPrecision = "";
			string strScale = "";
			string strSize = "";

			if (p_oRow["ColumnSize"] != null)
				strSize = Convert.ToString(p_oRow["ColumnSize"]);

			if (p_oRow["NumericPrecision"] != null)
				strPrecision = Convert.ToString(p_oRow["NumericPrecision"]);

			if (p_oRow["NumericScale"] != null)
				strScale = Convert.ToString(p_oRow["NumericScale"]);

			return strLine = FormatSelectSqlFieldItem(strColumn,strDataType,strSize,strPrecision,strScale,p_bReservedWordFormatting);
		}
        /// <summary>
        /// Format reserved words used as column names in SQL expressions.
        /// </summary>
        /// <param name="p_strColumnName"></param>
        /// <returns></returns>
		public string FormatReservedWordColumnName(string p_strColumnName)
		{
            if (p_strColumnName.Trim().ToUpper() == "VALUE" ||
                p_strColumnName.Trim().ToUpper() == "USE" ||
                p_strColumnName.Trim().ToUpper() == "YEAR" ||
                p_strColumnName.Trim().ToUpper() == "DESC" ||
                p_strColumnName.Trim().ToUpper() == "TO" ||
                p_strColumnName.Trim().ToUpper() == "AS")
            {
                p_strColumnName = "`" + p_strColumnName.Trim() + "`";
            }
            else if (p_strColumnName.Substring(0, 1) == "_")
            {
                p_strColumnName = "[" + p_strColumnName.Trim() + "]";
            }

			return p_strColumnName;

		}
		public string FormatReservedWordsInColumnNameList(string p_strList,string p_strDelimiter)
		{
		    string strList="";
			if (p_strList.Trim().Length == 0) return "";
			string[] strArray = p_strList.Split(p_strDelimiter.ToCharArray());
			for (int x=0;x<=strArray.Length -1 ;x++)
			{
				strList = strList + this.FormatReservedWordColumnName(strArray[x]) + ",";
			}
			strList = strList.Substring(0,strList.Length - 1);
			return strList;
		}
        /// <summary>
        /// Create ADO.NET data set and fill it with data from SQL execution
        /// </summary>
        /// <param name="strConn"></param>
        /// <param name="strSQL"></param>
        /// <param name="strTableName"></param>
		public void CreateDataSet(string strConn,
			string strSQL,string strTableName)
		{
			this.m_intError=0;
			this.m_strError="";
			try
			{
				this.OpenConnection(strConn);
				if (this.m_intError == 0)
				{
					this.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter(strSQL, this.m_OleDbConnection);
					this.m_DataSet = new DataSet();
					this.m_OleDbDataAdapter.Fill(this.m_DataSet,strTableName);
					this.m_OleDbConnection.Close();
				}

			}
			catch (Exception caught)
			{
				this.m_intError = -1;
				this.m_strError = caught.Message + " : SQL query command " + strSQL + " failed" ;
                if (_bDisplayErrors)
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:CreateDataSet  \n" + 
					"Err Msg - " + this.m_strError,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_OleDbConnection.Close();
				this.m_OleDbDataAdapter = null;
				this.m_DataSet  = null;
				return;
			}
			
		}
        /// <summary>
        /// Create ADO.NET data set and fill it with data from SQL execution 
        /// </summary>
        /// <param name="p_conn"></param>
        /// <param name="strSQL"></param>
        /// <param name="strTableName"></param>
		public void CreateDataSet(System.Data.OleDb.OleDbConnection p_conn,
			string strSQL,string strTableName)
		{
			this.m_intError=0;
			this.m_strError="";
			try
			{
					this.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter(strSQL, p_conn);
					this.m_DataSet = new DataSet();
					this.m_OleDbDataAdapter.Fill(this.m_DataSet,strTableName);
			}
			catch (Exception caught)
			{
				this.m_intError = -1;
				this.m_strError = caught.Message + " : SQL query command " + strSQL + " failed" ;
                if (_bDisplayErrors)
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:CreateDataSet  \n" + 
					"Err Msg - " + this.m_strError,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_OleDbDataAdapter = null;
				this.m_DataSet  = null;
				return;
			}
			
		}

        /// <summary>
        /// Create ADO.NET data table and fill it with data from SQL execution 
        /// </summary>
        /// <param name="p_conn"></param>
        /// <param name="strSQL"></param>
        /// <param name="strTableName"></param>
        /// <param name="bAcceptChanges">Set to false if you want to write data to another connection</param>
        public void CreateDataTable(System.Data.OleDb.OleDbConnection p_conn,
            string strSQL, string strTableName, bool bAcceptChanges)
        {
            this.m_intError = 0;
            this.m_strError = "";
            try
            {
                this.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter(strSQL, p_conn);
                this.m_OleDbDataAdapter.AcceptChangesDuringFill = bAcceptChanges;
                this.m_DataTable = new DataTable();
                this.m_OleDbDataAdapter.Fill(this.m_DataTable);
            }
            catch (Exception caught)
            {
                this.m_intError = -1;
                this.m_strError = caught.Message + " : SQL query command " + strSQL + " failed";
                if (_bDisplayErrors)
                    MessageBox.Show("!!Error!! \n" +
                        "Module - ado_data_access:CreateDataTable  \n" +
                        "Err Msg - " + this.m_strError,
                        "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_OleDbDataAdapter = null;
                this.m_DataSet = null;
                return;
            }

        }

		public void AddSQLQueryToDataSet(System.Data.OleDb.OleDbConnection p_conn,
			ref System.Data.OleDb.OleDbDataAdapter p_da,
			ref System.Data.DataSet p_ds,
			string strSQL, 
			string strTableName)
		{
			this.m_intError=0;
			this.m_strError="";
			System.Data.OleDb.OleDbCommand p_OleDbCommand;
			try
			{
				p_OleDbCommand = p_conn.CreateCommand();
				p_OleDbCommand.CommandText = strSQL;
				p_da.SelectCommand = p_OleDbCommand;
				p_da.Fill(p_ds,strTableName);
			}
			catch (Exception caught)
			{
				this.m_intError = -1;
				this.m_strError = caught.Message + " : SQL query command " + strSQL + " failed" ;
                if (_bDisplayErrors)
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:AddSQLQueryToDataSet  \n" + 
					"Err Msg - " + this.m_strError,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				MessageBox.Show(this.m_strError);
			}

		}
		public void AddSQLQueryToDataSet(System.Data.OleDb.OleDbConnection p_conn,
			ref System.Data.OleDb.OleDbDataAdapter p_da,
			ref System.Data.DataSet p_ds,
			ref System.Data.OleDb.OleDbTransaction p_trans,
			string strSQL, 
			string strTableName)
		{
			this.m_intError=0;
			this.m_strError="";
			System.Data.OleDb.OleDbCommand p_OleDbCommand;
			try
			{
				p_OleDbCommand = p_conn.CreateCommand();
				p_OleDbCommand.CommandText = strSQL;
				p_OleDbCommand.Transaction = p_trans;
				p_da.SelectCommand = p_OleDbCommand;
				p_da.Fill(p_ds,strTableName);
			}
			catch (Exception caught)
			{
				this.m_intError = -1;
				this.m_strError = caught.Message + " : SQL query command " + strSQL + " failed" ;
                if (_bDisplayErrors)
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:AddSQLQueryToDataSet  \n" + 
					"Err Msg - " + this.m_strError,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				MessageBox.Show(this.m_strError);
			}

		}
	
        /// <summary>
        /// Return number of rows
        /// </summary>
        /// <param name="strConn"></param>
        /// <param name="strSQL"></param>
        /// <param name="strTableName"></param>
        /// <returns></returns>
		public long getRecordCount(string strConn, string strSQL, string strTableName)
		{
            long intRecTtl = 0;
            this.m_intError = 0;
            this.m_strError = "";
		    using (var p_OleDbConnection = new OleDbConnection(strConn))
            using (var p_OleDbCommand = new OleDbCommand(strSQL, p_OleDbConnection))
            {
                try
                {
                    p_OleDbCommand.Connection.Open();
                    intRecTtl = Convert.ToInt32(p_OleDbCommand.ExecuteScalar());
                }
                catch (Exception caught)
                {
                    this.m_intError = -1;
                    this.m_strError = caught.Message + "  SQL query command: " + strSQL + " failed";
                    if (_bDisplayErrors)
                        MessageBox.Show("!!Error!! \n" +
                                        "Module - ado_data_access:getRecordCount  \n" +
                                        "Err Msg - " + this.m_strError,
                            "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                    MessageBox.Show(this.m_strError);
                }
            }
            return intRecTtl;
		}
        /// <summary>
        /// return number of rows
        /// </summary>
        /// <param name="p_conn"></param>
        /// <param name="strSQL"></param>
        /// <param name="strTableName"></param>
        /// <returns></returns>
        public int getRecordCount(System.Data.OleDb.OleDbConnection p_conn, string strSQL, string strTableName)
        {
            int intRecTtl = 0;
            this.m_intError = 0;
            this.m_strError = "";
            using (var p_OleDbCommand = new OleDbCommand(strSQL, p_conn))
            {
                try
                {
                    intRecTtl = Convert.ToInt32(p_OleDbCommand.ExecuteScalar());
                }
                catch (Exception caught)
                {
                    this.m_intError = -1;
                    this.m_strError = caught.Message + "  SQL query command: " + strSQL + " failed";
                    if (_bDisplayErrors)
                        MessageBox.Show("!!Error!! \n" +
                                        "Module - ado_data_access:getRecordCount  \n" +
                                        "Err Msg - " + this.m_strError,
                            "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                }
            }
            return intRecTtl;
        }

        /// <summary>
        /// return number of rows
        /// </summary>
        /// <param name="p_conn"></param>
        /// <param name="p_trans"></param>
        /// <param name="strSQL"></param>
        /// <param name="strTableName"></param>
        /// <returns></returns>
		public long getRecordCount(System.Data.OleDb.OleDbConnection p_conn, System.Data.OleDb.OleDbTransaction p_trans, string strSQL, string strTableName)
        {
            int intRecTtl = 0;
			this.m_intError=0;
			this.m_strError="";
            using (var p_OleDbCommand = new OleDbCommand(strSQL, p_conn, p_trans))
            {
                try
                {
                    intRecTtl = Convert.ToInt32(p_OleDbCommand.ExecuteScalar());
                }
                catch (Exception caught)
                {
                    this.m_intError = -1;
                    this.m_strError = caught.Message + "  SQL query command: " + strSQL + " failed";
                    if (_bDisplayErrors)
                        MessageBox.Show("!!Error!! \n" +
                                        "Module - ado_data_access:getRecordCount  \n" +
                                        "Err Msg - " + this.m_strError,
                            "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                }
            }
	        return intRecTtl;
		}
        /// <summary>
        /// return number of rows
        /// </summary>
        /// <param name="p_conn"></param>
        /// <param name="strSQL"></param>
        /// <param name="strTableName"></param>
        /// <returns></returns>
        public uint getRecordCountUInt(System.Data.OleDb.OleDbConnection p_conn, string strSQL, string strTableName)
        {
            uint intRecTtl = 0;
            this.m_intError = 0;
            this.m_strError = "";
            using (var p_OleDbCommand = new OleDbCommand(strSQL, p_conn))
            {
                try
                {
                    intRecTtl = Convert.ToUInt32(p_OleDbCommand.ExecuteScalar());
                }
                catch (Exception caught)
                {
                    this.m_intError = -1;
                    this.m_strError = caught.Message + "  SQL query command: " + strSQL + " failed";
                    if (_bDisplayErrors)
                        MessageBox.Show("!!Error!! \n" +
                                        "Module - ado_data_access:getRecordCount  \n" +
                                        "Err Msg - " + this.m_strError,
                            "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                }
            }
            return intRecTtl;
        }
        /// <summary>
        /// Return a ADO.NET datatable after converting from a dataview
        /// </summary>
        /// <param name="p_dv"></param>
        /// <param name="p_frmTherm"></param>
        /// <returns></returns>
		public System.Data.DataTable ConvertDataViewToDataTable(
			System.Data.DataView p_dv, ref FIA_Biosum_Manager.frmTherm p_frmTherm)
		{
			int x=0;
			int p_intCounter=p_frmTherm.progressBar1.Value;
			System.Data.DataTable p_dtNew;
            
			if (p_frmTherm.Text.Trim().Length == 0)
			{
				p_frmTherm.Text = "Converting DataView To DataTable";
			}
			//copy exact structure from the view to the new table
			p_dtNew = p_dv.Table.Clone();
			int idx = 0;
			p_frmTherm.progressBar1.Maximum += p_dtNew.Columns.Count;
			//create an array containing all the column names in the new data table
			string[] strColNames = new string[p_dtNew.Columns.Count];
			for (x=0;x<=p_dtNew.Columns.Count-1;x++)
			{
				System.Windows.Forms.Application.DoEvents();
				if (p_frmTherm.AbortProcess == true) 
				{
					return p_dtNew;
				}
				p_intCounter++;
				p_frmTherm.Increment(p_intCounter);
				strColNames[idx++] = p_dtNew.Columns[x].ColumnName;
			}
            p_frmTherm.progressBar1.Maximum += p_dv.Table.Rows.Count;
			//append each row in the dataview to the new table
			System.Collections.IEnumerator viewEnumerator = p_dv.GetEnumerator();
			
			while (viewEnumerator.MoveNext())
			{
				System.Windows.Forms.Application.DoEvents();
				if (p_frmTherm.AbortProcess == true) 
				{
					return p_dtNew;
				}

				p_intCounter++;
				p_frmTherm.Increment(p_intCounter);

				DataRowView drv = (DataRowView)viewEnumerator.Current;
				DataRow dr = p_dtNew.NewRow();
				try
				{
					foreach (string strName in strColNames)
					{
						//value in data table row and column equal to value in 
						//dataview row and column value
						dr[strName] = drv[strName];
						
					}
				}
				catch (Exception ex)
				{
                    if (_bDisplayErrors)
					MessageBox.Show("!!Error!! \n" + 
						"Module - ado_data_access:ConvertDataViewToDataTable  \n" + 
						"Err Msg - " + ex.Message,
						"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);
				
				}
				//append the new row to the data table
				p_dtNew.Rows.Add(dr);
			}
			return p_dtNew;
		}
        /// <summary>
        /// return an MS Access connection string
        /// </summary>
        /// <param name="strMDBPathAndFile"></param>
        /// <param name="strUserId"></param>
        /// <param name="strPW"></param>
        /// <returns></returns>
		public string getMDBConnString(string strMDBPathAndFile,string strUserId,string strPW)
		{
            string strProvider="";
			strProvider="Microsoft.Ace.OLEDB.12.0";

			if (strPW.Trim().Length > 0) 
			{
				//password protected mdb file
				return "Provider=" + strProvider + ";Data Source=" + strMDBPathAndFile + ";Jet OLEDB:Database Password=" + strPW.Trim() + ";";
			}
			else
				//standard security
                return "Provider=" + strProvider + ";Data Source=" + strMDBPathAndFile + ";User Id=" + strUserId.Trim() + ";Password=" + strPW.Trim() + ";";
		}
		public void getScenarioDataSourceConnStringAndTable(ref string strMDBFile,
			ref string strTable, 
			ref string strConn,
			string strTableType,  
			string strScenarioId,
			System.Data.OleDb.OleDbConnection p_OleDbScenarioConn)
		{
			string strSQL="select path,file, table_name from scenario_datasource " + 
				          "where trim(ucase(table_type)) = '" + strTableType.Trim().ToUpper() + "' " + 
				          "and trim(ucase(scenario_id)) = '"  + strScenarioId.Trim().ToUpper() + "';";
             
			this.SqlQueryReader(p_OleDbScenarioConn,strSQL);
            
			if (this.m_intError == 0)
			{

				while (this.m_OleDbDataReader.Read())
				{
					strTable = this.m_OleDbDataReader["table_name"].ToString();
					strMDBFile = this.m_OleDbDataReader["path"].ToString().Trim() + "\\" + this.m_OleDbDataReader["file"].ToString().Trim();
					strConn = this.getMDBConnString(strMDBFile,"admin",""); // "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBFile + ";User Id=admin;Password=;";
					break;
				}
				this.m_OleDbDataReader.Close();
			}
			this.m_OleDbDataReader = null;
			this.m_OleDbCommand = null;

		}
		public int getNumberOfCoreTables(string strScenarioMDB, string strScenarioId)
		{

			int intCount=0;
			System.Data.OleDb.OleDbConnection p_conn;
                        
			ado_data_access p_ado = new ado_data_access();
			p_conn = new System.Data.OleDb.OleDbConnection();
			string strConn = this.getMDBConnString(strScenarioMDB,"admin","");
			p_ado.OpenConnection(strConn, ref p_conn);	
			if (p_ado.m_intError != 0)
			{
				p_ado = null;
				return intCount;
			}
			string strSQL = "SELECT table_name FROM scenario_datasource WHERE " + 
				" scenario_id = '" + strScenarioId + "';";

			p_ado.SqlQueryReader(p_conn, strSQL);
			if (p_ado.m_intError == 0)
			{
				while (p_ado.m_OleDbDataReader.Read())
				{
					if (p_ado.m_OleDbDataReader["table_name"] != System.DBNull.Value)
					{
						if (p_ado.m_OleDbDataReader["table_name"].ToString().Trim().Length > 0)
						{
							intCount++; 
						}
					}
				}
			}
			p_ado.m_OleDbDataReader.Close();
			p_ado.m_OleDbDataReader = null;
			p_conn.Close();
			p_conn=null;
			p_ado = null;
			return intCount;
		}
	
		public string CreateSQLNOTINString(System.Data.OleDb.OleDbConnection p_conn,string strSQL)
		{
			string str = "";
			this.SqlQueryReader(p_conn, strSQL);
			if (this.m_intError == 0)
			{
				if (this.m_OleDbDataReader.HasRows)
				{
					while (this.m_OleDbDataReader.Read())
					{
						if (str.Trim().Length == 0)
						{
							str = this.m_OleDbDataReader[0].ToString().Trim();
						}
						else
						{
						    str += "," + this.m_OleDbDataReader[0].ToString().Trim();
						}
					}
				}
				this.m_OleDbDataReader.Close();
			}
			return str;
			
			
		}
		public string CreateCommaDelimitedList(System.Data.OleDb.OleDbConnection p_conn,string strSQL, string p_strEnclosingCharacter)
		{
			string str = "";
			this.SqlQueryReader(p_conn, strSQL);
			if (this.m_intError == 0)
			{
				if (this.m_OleDbDataReader.HasRows)
				{
					
					while (this.m_OleDbDataReader.Read())
					{
						if (str.Trim().Length == 0)
						{
							
							str = p_strEnclosingCharacter + this.m_OleDbDataReader[0].ToString().Trim() + p_strEnclosingCharacter;
						}
						else
						{
							str += "," + p_strEnclosingCharacter +  this.m_OleDbDataReader[0].ToString().Trim() + p_strEnclosingCharacter;
						}
					}
				}
				this.m_OleDbDataReader.Close();
			}
			return str;
			
			
		}
        
	    
	    /// <summary>
        /// return a single string value resulting from SQL command 
        /// </summary>
        /// <param name="p_conn"></param>
        /// <param name="strSQL"></param>
        /// <param name="strTableName"></param>
        /// <returns></returns>
		public string getSingleStringValueFromSQLQuery(System.Data.OleDb.OleDbConnection p_conn, string strSQL, string strTableName)
		{
			string strValue="";
			this.m_intError=0;
			this.m_strError="";
            using (var p_OleDbCommand = new OleDbCommand(strSQL, p_conn))
            {
                try
                {
                    strValue = Convert.ToString(p_OleDbCommand.ExecuteScalar());
                }
                catch (Exception caught)
                {
                    this.m_intError = -1;
                    this.m_strError = caught.Message + "  SQL query command: " + strSQL + " failed";
                    if (_bDisplayErrors)
                        MessageBox.Show("!!Error!! \n" +
                                        "Module - ado_data_access:getSingleStringValueFromSQLQuery  \n" +
                                        "Err Msg - " + this.m_strError,
                            "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                }
            }
			return strValue;
		}


	    
	    /// <summary>
        /// return a single string value resulting from SQL command 
        /// </summary>
        /// <param name="p_conn"></param>
        /// <param name="p_trans"></param>
        /// <param name="strSQL"></param>
        /// <param name="strTableName"></param>
        /// <returns></returns>
		public string getSingleStringValueFromSQLQuery(System.Data.OleDb.OleDbConnection p_conn,
			System.Data.OleDb.OleDbTransaction p_trans, string strSQL,string strTableName)
		{
			string strValue="";
			this.m_intError=0;
			this.m_strError="";
            using (var p_OleDbCommand = new OleDbCommand(strSQL, p_conn, p_trans))
            {
                try
                {
                    strValue = Convert.ToString(p_OleDbCommand.ExecuteScalar());
                }
                catch (Exception caught)
                {
                    this.m_intError = -1;
                    this.m_strError = caught.Message + "  SQL query command: " + strSQL + " failed";
                    if (_bDisplayErrors)
                        MessageBox.Show("!!Error!! \n" +
                                        "Module - ado_data_access:getSingleStringValueFromSQLQuery  \n" +
                                        "Err Msg - " + this.m_strError,
                            "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                }
            }
			return strValue;

		}
        /// <summary>
        /// return a single numeric double value resulting from SQL command 
        /// </summary>
        /// <param name="p_conn"></param>
        /// <param name="strSQL"></param>
        /// <param name="strTableName"></param>
        /// <returns></returns>
		public double getSingleDoubleValueFromSQLQuery(System.Data.OleDb.OleDbConnection p_conn, string strSQL, string strTableName)
		{
			double dblValue=-1;
			this.m_intError=0;
			this.m_strError="";
            using (var p_OleDbCommand = new OleDbCommand(strSQL, p_conn))
            {
                try
                {
                    dblValue = Convert.ToDouble(p_OleDbCommand.ExecuteScalar());
                }
                catch (Exception caught)
                {
                    this.m_intError = -1;
                    this.m_strError = caught.Message + "  SQL query command: " + strSQL + " failed";
                    if (_bDisplayErrors)
                        MessageBox.Show("!!Error!! \n" +
                                        "Module - ado_data_access:getSingleStringValueFromSQLQuery  \n" +
                                        "Err Msg - " + this.m_strError,
                            "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                }
            }
			return dblValue;

		}
        /// <summary>
        /// return a single numeric double value resulting from SQL command 
        /// </summary>
        /// <param name="p_conn"></param>
        /// <param name="p_trans"></param>
        /// <param name="strSQL"></param>
        /// <param name="strTableName"></param>
        /// <returns></returns>
        public double getSingleDoubleValueFromSQLQuery(System.Data.OleDb.OleDbConnection p_conn,
            System.Data.OleDb.OleDbTransaction p_trans,
            string strSQL, string strTableName)
        {
            System.Data.OleDb.OleDbCommand p_OleDbCommand;
            double dblValue = -1;
            this.m_intError = 0;
            this.m_strError = "";
            try
            {

                p_OleDbCommand = p_conn.CreateCommand();
                p_OleDbCommand.CommandText = strSQL;
                p_OleDbCommand.Transaction = p_trans;
                dblValue = Convert.ToInt32(p_OleDbCommand.ExecuteScalar());


            }
            catch (Exception caught)
            {
                this.m_intError = -1;
                this.m_strError = caught.Message + "  SQL query command: " + strSQL + " failed";
                if (_bDisplayErrors)
                MessageBox.Show("!!Error!! \n" +
                    "Module - ado_data_access:getSingleDoubleValueFromSQLQuery  \n" +
                    "Err Msg - " + this.m_strError,
                    "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);

            }
            p_OleDbCommand = null;
            return dblValue;

        }

		/// <summary>
		/// Converts a given delimited file into a dataset. 
		/// Assumes that the first line    
		/// of the text file contains the column names.
		/// </summary>
		/// <param name="File">The name of the file to open</param>    
		/// <param name="TableName">The name of the 
		/// Table to be made within the DataSet returned</param>
		/// <param name="delimiter">The string to delimit by</param>
		/// <returns></returns>  
		public void ConvertDelimitedTextToDataTable(System.Data.DataSet p_ds, 
			                                            string p_strFile, 
                                             			string p_strTableName, string p_strDelimiter)
		{   
			//The DataSet to Return
			//DataSet result = new DataSet();
            this.m_intError=0;
			int i = 0;
			int x;
			int z;
			int intPos;
			try
			{
				//Open the file in a stream reader.
				StreamReader s = new StreamReader(p_strFile);
        
				//Split the first line into the columns       
				string[] columns = s.ReadLine().Split(p_strDelimiter.ToCharArray());
  
				//Add the new DataTable to the RecordSet
				p_ds.Tables.Add(p_strTableName);
    
				//Cycle the colums, adding those that don't exist yet 
				//and sequencing the one that do.
				foreach(string col in columns)
				{
					bool added = false;
					string next = "";
					i=0;
					while(!added)        
					{
						//Build the column name and remove any unwanted characters.
						string columnname = col + next;
						columnname = columnname.Replace("#","");
						columnname = columnname.Replace("'","");
						columnname = columnname.Replace("&","");
						columnname = columnname.Replace("\"","");
        
						//See if the column already exists
						if(!p_ds.Tables[p_strTableName].Columns.Contains(columnname))
						{
							//if it doesn't then we add it here and mark it as added
							p_ds.Tables[p_strTableName].Columns.Add(columnname);
							added = true;
						}
						else
						{
							//if it did exist then we increment the sequencer and try again.
							i++;  
							next = "_" + i.ToString();
						}         
					}
				}
    
				//Read the rest of the data in the file.        
				string AllData = s.ReadToEnd();
    
				//Split off each row at the Carriage Return/Line Feed
				//Default line ending in most <A class=iAs style="FONT-WEIGHT: normal; FONT-SIZE: 100%; PADDING-BOTTOM: 1px; COLOR: darkgreen; BORDER-BOTTOM: darkgreen 0.07em solid; BACKGROUND-COLOR: transparent; TEXT-DECORATION: underline" href="#" target=_blank itxtdid="2592535">windows</A> exports.  
				//You may have to edit this to match your particular file.
				//This will work for Excel, Access, etc. default exports.
				string[] rows = AllData.Split("\r\n".ToCharArray());

 
				//Now add each row to the DataSet        
				foreach(string r in rows)
				{
					if (r.Trim().Length > 0)
					{
						string[] items = r.Split(p_strDelimiter.ToCharArray());
						//There could be occurances of the delimiter in a text string.
						//If there are then the length of the items array will be 
						//greater than the number of table columns.
						if (items.Length != p_ds.Tables[p_strTableName].Columns.Count)
						{
							intPos=0;
							//load each column value in the row individually
							//reformat the row and put it in strRow
							string strRow="";
							string strDelimiter="#";
							for (i=0;i<= p_ds.Tables[p_strTableName].Columns.Count - 1;i++)
							{
								string strBuildColumnString="";
								for (x=intPos;x<=r.Length-1;x++)
								{
									//check if we are at the end of the row
									if (x == r.Length - 1)
									{
										if (strBuildColumnString.Trim().Length > 0)
										{
											strRow = strRow + strBuildColumnString + strDelimiter;
										}
										else
										{
											strRow = strRow + " " + strDelimiter;
										}
									}
										
									else 
										
									{
										//check for starting double quote
										if (r.Substring(x,1)=="\"")
										{
											if (strBuildColumnString.Trim().Length == 0)
											{
												//check for null value
												if (r.Substring(x,2)=="\"\"")
												{
													strRow = strRow + " " + strDelimiter;
													intPos = x+1;
													//find the delimiter
													for (z=intPos;z<=r.Length-1;z++)
													{
														if (r.Substring(z,p_strDelimiter.Length)==p_strDelimiter)
														{
															intPos=z;
															break;
														}
													}
												}
												else if (r.Substring(x,2)=="\"" + ",")
												{
													strRow = strRow + " " + strDelimiter;
													intPos = x+1;
												}
												else
												{
											
											
													x=x+1;
													//another check for null value
													//if (r.Substring(x+1,1) != "\"")
													//{
													//MessageBox.Show(r.Substring(x+1,1));
													//check for the ending double quote
													strBuildColumnString=r.Substring(x,r.Length - x);
													z = strBuildColumnString.IndexOf("\"",2);
													if (strBuildColumnString.Trim() == "\"\"")
													{
														//no value
														strBuildColumnString=" ";
														z = r.IndexOf("\"",x);

													}
													else
													{
														strBuildColumnString=strBuildColumnString.Substring(0,z);
														z = r.IndexOf("\"",x+1);
													}
													//MessageBox.Show(strBuildColumnString + " " + strBuildColumnString.Length.ToString());
											
													//finished with the column
													strRow=strRow + strBuildColumnString + strDelimiter;
													//find where the next delimiter is
													intPos = r.IndexOf(p_strDelimiter,z);
													
												}
												break;

											}
										}
										else if (r.Substring(x,p_strDelimiter.Length) == p_strDelimiter)
										{
											if (strBuildColumnString.Trim().Length ==0)
											{
												//check for a null value
												if (r.Substring(x+1,p_strDelimiter.Length)==p_strDelimiter)
												{
													strRow = strRow + " " + strDelimiter;
													intPos=x+1;
													break;
												}
											}
											else
											{
												strRow=strRow + strBuildColumnString + strDelimiter;
												intPos=x;
												break;
											}
										}
										else
										{
									
											strBuildColumnString=strBuildColumnString + r.Substring(x,1);
										}
									}
									
								}

							}
							if (strRow.Trim().Length > 0) strRow=strRow.Substring(0,strRow.Length - strDelimiter.Length);
							
							//Split the row at the delimiter.
							string[] items2 = strRow.Split(strDelimiter.ToCharArray());
      
							//Add the item
							p_ds.Tables[p_strTableName].Rows.Add(items2);  
							
							}
						else
						{
							//remove unwanted characters
							for (x=0;x<=items.Length-1;x++)
							{
								//double and single quotations
								items[x] = items[x].Replace("\"","");
								items[x] = items[x].Replace("'","");
							}
							//Add the item
							p_ds.Tables[p_strTableName].Rows.Add(items);  

						}
					}
				}
			}
			catch (Exception caught)
			{

				this.m_intError=-1;
                if (_bDisplayErrors)
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:ConvertDelimitedTextToDataTable  \n" + 
					"Err Msg - " + caught.Message.ToString().Trim(),
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
    
			//Return the imported data.        

		}

		/// <summary>
		/// Create an oledb data adapter insert command. The select sql statement
		/// is used to get the data types of the fields used in the insert.
		/// </summary>
		/// <param name="p_conn">the oledb database connection</param>
		/// <param name="p_da">the data adapter</param>
		/// <param name="p_trans">oledb transaction object</param>
		/// <param name="p_strSQL">select sql statement containing fields in the insert command</param>
		/// <param name="p_strTable">table name that records are inserted</param>
		public void ConfigureDataAdapterInsertCommand(System.Data.OleDb.OleDbConnection p_conn, 
			                                          System.Data.OleDb.OleDbDataAdapter p_da,
													   System.Data.OleDb.OleDbTransaction p_trans,
			                                           string p_strSQL,string p_strTable)
		{

			this.m_intError=0;
			System.Data.DataTable p_dtTableSchema = this.getTableSchema(p_conn,p_trans, p_strSQL);
			if (this.m_intError !=0) return;
			string strFields = "";
			string strValues = "";
			int x;
			try
			{
				//Build the plot insert sql
				for (x=0; x<=p_dtTableSchema.Rows.Count-1;x++)
				{
					if (strFields.Trim().Length == 0)
					{
						strFields = "(";
					}
					else
					{	
						strFields = strFields + "," ;
					}
					strFields = strFields + p_dtTableSchema.Rows[x]["columnname"].ToString().Trim();
					if (strValues.Trim().Length == 0)
					{
						strValues = "(";
					}
					else
					{	
						strValues = strValues + ",";
					}
					strValues = strValues + "?";

				}
				strFields = strFields + ")";
				strValues = strValues + ");";
				//create an insert command 
				p_da.InsertCommand = p_conn.CreateCommand();
				//bind the transaction object to the insert command
				p_da.InsertCommand.Transaction = p_trans;
				p_da.InsertCommand.CommandText = 
					"INSERT INTO " + p_strTable + " "  + strFields + " VALUES " + strValues;
				//define field datatypes for the data adapter
				for (x=0; x<=p_dtTableSchema.Rows.Count-1;x++)
				{
					strFields=p_dtTableSchema.Rows[x]["columnname"].ToString().Trim();
					switch (p_dtTableSchema.Rows[x]["datatype"].ToString().Trim())
					{
						case "System.String" :
							p_da.InsertCommand.Parameters.Add
								(strFields, 
								System.Data.OleDb.OleDbType.VarWChar,
								0,
								strFields);
							break;
						case "System.Double":
							p_da.InsertCommand.Parameters.Add
								(strFields, 
								System.Data.OleDb.OleDbType.Double,
								0,
								strFields);
							break;
						case "System.Boolean":
							p_da.InsertCommand.Parameters.Add
								(strFields, 
								System.Data.OleDb.OleDbType.Boolean,
								0,
								strFields);
							break;
						case "System.DateTime":
							p_da.InsertCommand.Parameters.Add
								(strFields, 
								System.Data.OleDb.OleDbType.DBTimeStamp,
								0,
								strFields);
							break;
						case "System.Decimal":
							p_da.InsertCommand.Parameters.Add
								(strFields, 
								System.Data.OleDb.OleDbType.Decimal,
								0,
								strFields);
							break;
						case "System.Int16":
							p_da.InsertCommand.Parameters.Add
								(strFields, 
								System.Data.OleDb.OleDbType.SmallInt,
								0,
								strFields);
							break;
						case "System.Int32":
							p_da.InsertCommand.Parameters.Add
								(strFields, 
								System.Data.OleDb.OleDbType.Integer,
								0,
								strFields);
							break;
						case "System.Int64":
							p_da.InsertCommand.Parameters.Add
								(strFields, 
								System.Data.OleDb.OleDbType.BigInt,
								0,
								strFields);
							break;
						case "System.SByte":
							p_da.InsertCommand.Parameters.Add
								(strFields, 
								System.Data.OleDb.OleDbType.SmallInt,
								0,
								strFields);
							break;
						case "System.Byte":
							p_da.InsertCommand.Parameters.Add
								(strFields, 
								System.Data.OleDb.OleDbType.TinyInt,
								0,
								strFields);
							break;
						case "System.Single":
							p_da.InsertCommand.Parameters.Add
								(strFields, 
								System.Data.OleDb.OleDbType.Single,
								0,
								strFields);
							break;
						default:
							MessageBox.Show("Could Not Set Data Adapter Parameter For DataType " + p_dtTableSchema.Rows[x]["datatype"].ToString().Trim());
							break;
					}
									
				}
			}
			catch (Exception e)
			{
                this.m_strError = e.Message;
                if (_bDisplayErrors)
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:ConfigureDataAdapterInsertCommand  \n" + 
					"Err Msg - " + e.Message.ToString().Trim(),
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
               
			}

		}
/// <summary>
/// create the update command for the data adapter. 
/// </summary>
/// <param name="p_conn">oledb connection object</param>
/// <param name="p_da">oledb dataadapter object</param>
/// <param name="p_trans">oledb transaction object</param>
/// <param name="p_strSQL">select sql statement to get the update field data types</param>
/// <param name="p_strSQLUniqueRecordFields">select SQL statement listing fields used for a records unique id are queried and their field types obtained and added to the dataadapter updates parameters list</param>
/// <param name="p_strTable">table name to be updated</param>
		public void ConfigureDataAdapterUpdateCommand(System.Data.OleDb.OleDbConnection p_conn, 
			System.Data.OleDb.OleDbDataAdapter p_da,
			System.Data.OleDb.OleDbTransaction p_trans,
			string p_strSQL,string p_strSQLUniqueRecordFields,string p_strTable)
		{

			this.m_intError=0;
			System.Data.DataTable p_dtTableSchema = this.getTableSchema(p_conn,p_trans, p_strSQL);
			System.Data.DataTable p_dtTableSchema2 = new DataTable();
			if (this.m_intError !=0) return;
			string strField = "";
			string strValue = "";
			string strSQL="";
			//string strWhere="";
			int x;
			try
			{
				//Build the plot update sql
				for (x=0; x<=p_dtTableSchema.Rows.Count-1;x++)
				{
					strField = p_dtTableSchema.Rows[x]["columnname"].ToString().Trim();
					if (strValue.Trim().Length == 0)
					{
						strValue = strField + "=?";
					}
					else
					{	
						strValue += "," + strField + "=?";
					}
				}
				
				strSQL = 
					"UPDATE " + p_strTable + " SET "  +  strValue;

				//get the unique record id
				if (p_strSQLUniqueRecordFields.Trim().Length > 0)
				{
					strValue="";
					p_dtTableSchema2 = this.getTableSchema(p_conn,p_trans,p_strSQLUniqueRecordFields);
					if (this.m_intError !=0) return;
					//build the where condition
					for (x=0; x<=p_dtTableSchema2.Rows.Count-1;x++)
					{
						strField = p_dtTableSchema2.Rows[x]["columnname"].ToString().Trim();
						if (strValue.Trim().Length == 0)
						{
							strValue = strField + "=?";
						}
						else
						{	
							strValue += " AND " + strField + "=?";
						}
					}
					strSQL += " WHERE " + strValue;
				}


				//create an insert command 
				p_da.UpdateCommand = p_conn.CreateCommand();
				//bind the transaction object to the insert command
				p_da.UpdateCommand.Transaction = p_trans;
				p_da.UpdateCommand.CommandText = strSQL;

				//copy the table schema records containing update fields info to a new table
                System.Data.DataTable p_dt = p_dtTableSchema.Copy();

				//define field datatypes for the data adapter
				for (;;)
				{
					for (x=0; x<=p_dt.Rows.Count-1;x++)
					{
						strField=p_dt.Rows[x]["columnname"].ToString().Trim();
						switch (p_dt.Rows[x]["datatype"].ToString().Trim())
						{
							case "System.String" :
								p_da.UpdateCommand.Parameters.Add
									(strField, 
									System.Data.OleDb.OleDbType.VarWChar,
									0,
									strField);
								break;
							case "System.Double":
								p_da.UpdateCommand.Parameters.Add
									(strField, 
									System.Data.OleDb.OleDbType.Double,
									0,
									strField);
								break;
							case "System.Boolean":
								p_da.UpdateCommand.Parameters.Add
									(strField, 
									System.Data.OleDb.OleDbType.Boolean,
									0,
									strField);
								break;
							case "System.DateTime":
								p_da.UpdateCommand.Parameters.Add
									(strField, 
									System.Data.OleDb.OleDbType.DBTimeStamp,
									0,
									strField);
								break;
							case "System.Decimal":
								p_da.UpdateCommand.Parameters.Add
									(strField, 
									System.Data.OleDb.OleDbType.Decimal,
									0,
									strField);
								break;
							case "System.Int16":
								p_da.UpdateCommand.Parameters.Add
									(strField, 
									System.Data.OleDb.OleDbType.SmallInt,
									0,
									strField);
								break;
							case "System.Int32":
								p_da.UpdateCommand.Parameters.Add
									(strField, 
									System.Data.OleDb.OleDbType.Integer,
									0,
									strField);
								break;
							case "System.Int64":
								p_da.UpdateCommand.Parameters.Add
									(strField, 
									System.Data.OleDb.OleDbType.BigInt,
									0,
									strField);
								break;
							case "System.SByte":
								p_da.UpdateCommand.Parameters.Add
									(strField, 
									System.Data.OleDb.OleDbType.SmallInt,
									0,
									strField);
								break;
							case "System.Byte":
								p_da.UpdateCommand.Parameters.Add
									(strField, 
									System.Data.OleDb.OleDbType.TinyInt,
									0,
									strField);
								break;
							case "System.Single":
								p_da.UpdateCommand.Parameters.Add
									(strField, 
									System.Data.OleDb.OleDbType.Single,
									0,
									strField);
								break;
							default:
								MessageBox.Show("Could Not Set Data Adapter Parameter For DataType " + p_dt.Rows[x]["datatype"].ToString().Trim());
								break;
						}
									
					}
					if (p_strSQLUniqueRecordFields.Trim().Length > 0)
					{
						//clear the data table of all its records
						p_dt.Clear();
						//copy the table schema records containing where clause fields info to a new table
						p_dt = p_dtTableSchema2.Copy();
						p_strSQLUniqueRecordFields = "";
					}
					else
					{
						break;
					}
				}
			}
			catch (Exception e)
			{
                this.m_strError = e.Message;
                if (_bDisplayErrors)
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:ConfigureDataAdapterUpdateCommand  \n" + 
					"Err Msg - " + e.Message.ToString().Trim(),
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
		}

		/// <summary>
		/// create the delete command for the data adapter. 
		/// </summary>
		/// <param name="p_conn">oledb connection object</param>
		/// <param name="p_da">oledb dataadapter object</param>
		/// <param name="p_trans">oledb transaction object</param>
		/// <param name="p_strSQLUniqueRecordFields">select SQL statement listing fields used for a records unique id are queried and their field types obtained and added to the dataadapter delete command parameters list</param>
		/// <param name="p_strTable">table name to be updated</param>
		public void ConfigureDataAdapterDeleteCommand(System.Data.OleDb.OleDbConnection p_conn, 
			System.Data.OleDb.OleDbDataAdapter p_da,
			System.Data.OleDb.OleDbTransaction p_trans,
			string p_strSQLUniqueRecordFields,string p_strTable)
		{

			this.m_intError=0;
			System.Data.DataTable p_dt = this.getTableSchema(p_conn,p_trans, p_strSQLUniqueRecordFields);
			if (this.m_intError !=0) return;
			string strField = "";
			string strValue = "";
			string strSQL="";
			//string strWhere="";
			int x;
			try
			{
				strSQL = "DELETE FROM " + p_strTable + " ";
				//build the where condition
				for (x=0; x<=p_dt.Rows.Count-1;x++)
				{
					strField = p_dt.Rows[x]["columnname"].ToString().Trim();
					if (strValue.Trim().Length == 0)
					{
						strValue = strField + "=?";
					}
					else
					{	
						strValue += " AND " + strField + "=?";
					}
				}
				strSQL += " WHERE " + strValue;
				


				//create an insert command 
				p_da.DeleteCommand = p_conn.CreateCommand();
				//bind the transaction object to the insert command
				p_da.DeleteCommand.Transaction = p_trans;
				p_da.DeleteCommand.CommandText = strSQL;

				

				//define field datatypes for the data adapter
				
				
				for (x=0; x<=p_dt.Rows.Count-1;x++)
				{
					strField=p_dt.Rows[x]["columnname"].ToString().Trim();
					switch (p_dt.Rows[x]["datatype"].ToString().Trim())
					{
						case "System.String" :
							p_da.DeleteCommand.Parameters.Add
								(strField, 
								System.Data.OleDb.OleDbType.VarWChar,
								0,
								strField);
							break;
						case "System.Double":
							p_da.DeleteCommand.Parameters.Add
								(strField, 
								System.Data.OleDb.OleDbType.Double,
								0,
								strField);
							break;
						case "System.Boolean":
							p_da.DeleteCommand.Parameters.Add
								(strField, 
								System.Data.OleDb.OleDbType.Boolean,
								0,
								strField);
							break;
						case "System.DateTime":
							p_da.DeleteCommand.Parameters.Add
								(strField, 
								System.Data.OleDb.OleDbType.DBTimeStamp,
								0,
								strField);
							break;
						case "System.Decimal":
							p_da.DeleteCommand.Parameters.Add
								(strField, 
								System.Data.OleDb.OleDbType.Decimal,
								0,
								strField);
							break;
						case "System.Int16":
							p_da.DeleteCommand.Parameters.Add
								(strField, 
								System.Data.OleDb.OleDbType.SmallInt,
								0,
								strField);
							break;
						case "System.Int32":
							p_da.DeleteCommand.Parameters.Add
								(strField, 
								System.Data.OleDb.OleDbType.Integer,
								0,
								strField);
							break;
						case "System.Int64":
							p_da.DeleteCommand.Parameters.Add
								(strField, 
								System.Data.OleDb.OleDbType.BigInt,
								0,
								strField);
							break;
						case "System.SByte":
							p_da.DeleteCommand.Parameters.Add
								(strField, 
								System.Data.OleDb.OleDbType.SmallInt,
								0,
								strField);
							break;
						case "System.Byte":
							p_da.DeleteCommand.Parameters.Add
								(strField, 
								System.Data.OleDb.OleDbType.TinyInt,
								0,
								strField);
							break;
						case "System.Single":
							p_da.DeleteCommand.Parameters.Add
								(strField, 
								System.Data.OleDb.OleDbType.Single,
								0,
								strField);
							break;
						default:
							MessageBox.Show("Could Not Set Data Adapter Parameter For DataType " + p_dt.Rows[x]["datatype"].ToString().Trim());
							break;
					}
									
				}
			
					
				
			}
			catch (Exception e)
			{
                m_strError = e.Message;
                if (_bDisplayErrors)
				MessageBox.Show("!!Error!! \n" + 
					"Module - ado_data_access:ConfigureDataAdapterUpdateCommand  \n" + 
					"Err Msg - " + e.Message.ToString().Trim(),
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
		}
		/// <summary>
		/// See if there are other values not equal to the specified value
		/// Error return values: -100 = Table Does Not Exist; -200 Column does Not Exist
		/// </summary>
		/// <param name="p_oConn"></param>
		/// <param name="p_strTableName">Table to search</param>
		/// <param name="p_strColumnName">Column rows to search</param>
		/// <param name="p_strValue">search for values not equal to this specified value</param>
		/// <param name="p_bNumericDataType"></param>
		/// <returns></returns>
		public bool ValuesExistNotEqualToTargetValue(System.Data.OleDb.OleDbConnection p_oConn, 
			                                              string p_strTableName, 
			                                              string p_strColumnName, 
			                                              string p_strValue,
														  bool p_bNumericDataType)
		{
			m_intError=0;
			m_strError="";
			int z=0;
			bool bFound=false;
			string strSql="";

			if (TableExist(p_oConn,p_strTableName))
			{
				if (ColumnExist(p_oConn,p_strTableName,p_strColumnName))
				{
					strSql ="SELECT COUNT(*) FROM (SELECT TOP 1 * FROM " + p_strTableName.Trim() + " WHERE "; 
					if (p_bNumericDataType)
					{
						strSql = strSql + p_strColumnName.Trim() + " <> " + p_strValue.Trim() + ")";
						
					}
					else
					{
						strSql = strSql + "TRIM(" + p_strColumnName.Trim() + ") <> '" + p_strValue.Trim() + "')";
					}
					if (getRecordCount(p_oConn,strSql,p_strTableName.Trim()) > 0)
					{
						bFound=true;
					}
				}
				else
				{
					m_intError=this.ErrorCodeColumnNotFound;
					m_strError=p_strTableName + "." + p_strColumnName + " does not exist";
				}
			}
			else
			{
				m_intError=this.ErrorCodeTableNotFound;
				m_strError=p_strTableName + " does not exist";
			}
			
			return bFound;
		}
		/// <summary>
		/// See if the target value exists 
 		/// Error return values: -100 = Table Does Not Exist; -200 Column does Not Exist
		/// </summary>
		/// <param name="p_oAdo"></param>
		/// <param name="p_oConn"></param>
		/// <param name="p_strTableName"></param>
		/// <param name="p_intRxYear"></param>
		/// <returns></returns>
		public bool ValuesExistEqualToTargetValue(System.Data.OleDb.OleDbConnection p_oConn,
												  string p_strTableName, 
												  string p_strColumnName, 
												  string p_strValue,
												  bool p_bNumericDataType)
		{
			int z=0;
			

			m_intError=0;
			m_strError="";
			
			bool bFound=false;
			string strSql="";

			if (TableExist(p_oConn,p_strTableName))
			{
				if (ColumnExist(p_oConn,p_strTableName,p_strColumnName))
				{
					strSql ="SELECT COUNT(*) FROM (SELECT TOP 1 * FROM " + p_strTableName.Trim() + " WHERE "; 
					if (p_bNumericDataType)
					{
						strSql = strSql + p_strColumnName.Trim() + " = " + p_strValue.Trim() + ")";
						
					}
					else
					{
						strSql = strSql + "TRIM(" + p_strColumnName.Trim() + ") = '" + p_strValue.Trim() + "')";
					}
					if (getRecordCount(p_oConn,strSql,p_strTableName.Trim()) > 0)
					{
						bFound=true;
					}
				}
				else
				{
					m_intError=this.ErrorCodeColumnNotFound;
					m_strError=p_strTableName + "." + p_strColumnName + " does not exist";
				}
			}
			else
			{
				m_intError=this.ErrorCodeTableNotFound;
				m_strError=p_strTableName + " does not exist";
			}
			
			return bFound;

			
		}


		/// <summary>
		/// Alter an MS Access table to add a primary key index
		/// </summary>
		/// <param name="p_oConn"></param>
		/// <param name="p_strTableName"></param>
		/// <param name="p_strIndexName"></param>
		/// <param name="p_strColumnList"></param>
		public void AddPrimaryKey(System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName, string p_strIndexName,string p_strColumnList)
		{
			this.m_strSQL = "ALTER TABLE " + p_strTableName + " " + 
					        "ADD CONSTRAINT " + p_strIndexName + " " + 
					        "PRIMARY KEY (" + p_strColumnList + ")";
			this.SqlNonQuery(p_oConn,this.m_strSQL);
			
		}
        /// <summary>
        /// Alter an MS Access table to add an autonumber data type to a column
        /// </summary>
        /// <param name="p_oConn"></param>
        /// <param name="p_strTableName"></param>
        /// <param name="p_strColumnName"></param>
		public void AddAutoNumber(System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName,string p_strColumnName)
		{
			this.m_strSQL = "ALTER TABLE " + p_strTableName + " ALTER COLUMN " + p_strColumnName + " AUTOINCREMENT";
			SqlNonQuery(p_oConn,m_strSQL);
		}
        /// <summary>
        /// Create an index
        /// </summary>
        /// <param name="p_oConn"></param>
        /// <param name="p_strTableName"></param>
        /// <param name="p_strIndexName"></param>
        /// <param name="p_strColumnList"></param>
		public void AddIndex(System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName, string p_strIndexName, string p_strColumnList)
		{
			m_strSQL = "CREATE INDEX " + p_strIndexName + " " + 
				       "ON " + p_strTableName + " " + 
				       "(" + p_strColumnList + ")";
			SqlNonQuery(p_oConn,m_strSQL);

		}
        /// <summary>
        /// Create a unique index
        /// </summary>
        /// <param name="p_oConn"></param>
        /// <param name="p_strTableName"></param>
        /// <param name="p_strIndexName"></param>
        /// <param name="p_strColumnList"></param>
        public void AddUniqueIndex(System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName, string p_strIndexName, string p_strColumnList)
        {
            m_strSQL = "CREATE UNIQUE INDEX " + p_strIndexName + " " +
                       "ON " + p_strTableName + " " +
                       "(" + p_strColumnList + ")";
            SqlNonQuery(p_oConn, m_strSQL);

        }
        public void AddColumn(System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName, string p_strColumnName,string p_strDataType, string p_strSize)
        {
            if (p_strSize.Trim().Length > 0)
                SqlNonQuery(p_oConn, "ALTER TABLE " + p_strTableName + " " +
                            "ADD COLUMN " + p_strColumnName + " " + p_strDataType + " (" + p_strSize + ")");
            else
                SqlNonQuery(p_oConn, "ALTER TABLE " + p_strTableName + " " +
                            "ADD COLUMN " + p_strColumnName + " " + p_strDataType);
        }

        public void AddColumn(System.Data.OleDb.OleDbConnection p_oConn, string p_strTableName, string p_strColumnName, string p_strDataType, string p_strSize, string p_strDefault)
        {
            if (p_strDefault.Trim().Length < 1) 
                AddColumn(p_oConn, p_strTableName, p_strColumnName, p_strDataType, p_strSize);
            else if (p_strDefault.Trim().Length > 0 && p_strSize.Trim().Length > 0)
                SqlNonQuery(p_oConn, "ALTER TABLE " + p_strTableName + " " +
                            "ADD COLUMN " + p_strColumnName + " " + p_strDataType + " (" + p_strSize + ") DEFAULT '" + p_strDefault + "'");
            else if (p_strDefault.Trim().Length > 0 && p_strSize.Trim().Length < 1)
                SqlNonQuery(p_oConn, "ALTER TABLE " + p_strTableName + " " +
                            "ADD COLUMN " + p_strColumnName + " " + p_strDataType + " DEFAULT '" + p_strDefault + "'");
        }

		public bool ReconcileTableColumns(System.Data.OleDb.OleDbConnection p_oDestConn,
			string p_strDestTable,
			System.Data.OleDb.OleDbConnection p_oSourceConn,
			string p_strSourceTable)
		{
			bool bFound=false;
			bool bModified=false;
			int z,zz;
			System.Data.DataTable oDtSourceSchema = getTableSchema(p_oSourceConn,"SELECT * FROM " + p_strSourceTable);
			System.Data.DataTable oDtDestSchema = getTableSchema(p_oDestConn,"SELECT * FROM " + p_strDestTable);
			for (z=0;z<=oDtSourceSchema.Rows.Count-1;z++)
			{
				string strSourceColumnFormat="";
				string strDestColumnFormat="";			
				if (oDtSourceSchema.Rows[z]["ColumnName"] != System.DBNull.Value)
				{
					bFound=false;
					for (zz=0;zz<=oDtDestSchema.Rows.Count-1;zz++)
					{
						if (oDtDestSchema.Rows[zz]["ColumnName"] != System.DBNull.Value)
						{
							if (oDtSourceSchema.Rows[z]["ColumnName"].ToString().Trim().ToUpper() == 
								oDtDestSchema.Rows[zz]["ColumnName"].ToString().Trim().ToUpper())
							{
								strSourceColumnFormat = this.FormatCreateTableSqlFieldItem(oDtSourceSchema.Rows[z]);
								strDestColumnFormat = this.FormatCreateTableSqlFieldItem(oDtDestSchema.Rows[zz]);
								if (strSourceColumnFormat.Trim().ToUpper() != strDestColumnFormat.Trim().ToUpper())
								{
									//alter the column to the new specs
									this.m_strSQL = "ALTER TABLE " + p_strDestTable + " ALTER COLUMN " + strSourceColumnFormat;
									this.SqlNonQuery(p_oDestConn,m_strSQL);
									bModified=true;
								}
								bFound=true;
								break;
							}
						}
					}
					if (!bFound)
					{
						//column not found so let's add it
					    strSourceColumnFormat = this.FormatCreateTableSqlFieldItem(oDtSourceSchema.Rows[z]);
						SqlNonQuery(p_oDestConn,"ALTER TABLE " + p_strDestTable + " " + 
							"ADD COLUMN " + strSourceColumnFormat);
						bModified=true;
					}
				}
											
			}
			return bModified;

		}
        public bool DisplayErrors
        {
            get { return _bDisplayErrors; }
            set { _bDisplayErrors = value; }
        }
		public int ErrorCodeTableNotFound
		{
			get {return -100;}
		}
		public int ErrorCodeColumnNotFound
		{
			get {return -200;}
		}
		public int ErrorCodeNoErrors
		{
			get {return 0;}
		}
		public int ErrorCodeTableEmpty
		{
			get {return -2;}
		}
		


	

		


	}



}

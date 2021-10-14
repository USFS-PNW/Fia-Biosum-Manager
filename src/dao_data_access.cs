using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Access.Dao;
namespace FIA_Biosum_Manager
{
    /// <summary>
    /// Summary description for dao_access.
    /// </summary>
    public class dao_data_access
    {
        public Microsoft.Office.Interop.Access.Dao.DBEngineClass m_DaoDbEngine;
        public Microsoft.Office.Interop.Access.Dao.Workspace m_DaoWorkspace;
        public Microsoft.Office.Interop.Access.Dao.Database m_DaoDatabase;
        public Microsoft.Office.Interop.Access.Dao.TableDef m_DaoTableDef;
        public Microsoft.Office.Interop.Access.Dao.Databases m_DaoDatabases;
        public Microsoft.Office.Interop.Access.Dao.Field m_DaoField;
        public Microsoft.Office.Interop.Access.Dao.TableDefClass m_DaoTableDefClass;
        public Microsoft.Office.Interop.Access.Dao.Index m_DaoIndex;
        public Microsoft.Office.Interop.Access.Dao.Recordset m_DaoRecordSet;
        public int m_intErrorCount = 0;
        public int m_intError;
        public string m_strError;
        public string m_strTable;
        private bool _bDisplayErrors = true;
        private string m_strDebugFile = Environment.GetEnvironmentVariable("tmp").Trim() + "\\biosum_dao_debug.txt";
        private bool m_bDebug=false;



        public dao_data_access()
        {
            //
            // TODO: Add constructor logic here
            //
            if (m_bDebug)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//dao_data_access.constructor\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            if (m_bDebug) frmMain.g_oUtils.WriteText(m_strDebugFile, "this.m_DaoDbEngine = new Microsoft.Office.Interop.Access.Dao.DBEngineClass()\r\n");
            this.m_DaoDbEngine = new Microsoft.Office.Interop.Access.Dao.DBEngineClass();
            if (m_bDebug) frmMain.g_oUtils.WriteText(m_strDebugFile, "m_DaoDbEngine.DefaultType = Convert.ToInt32(Microsoft.Office.Interop.Access.Dao.WorkspaceTypeEnum.dbUseJet)\r\n");
            this.m_DaoDbEngine.DefaultType = Convert.ToInt32(Microsoft.Office.Interop.Access.Dao.WorkspaceTypeEnum.dbUseJet);
            if (m_bDebug) frmMain.g_oUtils.WriteText(m_strDebugFile, "m_DaoWorkspace = this.m_DaoDbEngine.CreateWorkspace('Replication', 'admin', '', Microsoft.Office.Interop.Access.Dao.WorkspaceTypeEnum.dbUseJet);\r\n");
            this.m_DaoWorkspace = this.m_DaoDbEngine.CreateWorkspace("Replication", "admin", "", Microsoft.Office.Interop.Access.Dao.WorkspaceTypeEnum.dbUseJet);
            if (m_bDebug)
            {
                this.WriteText(m_strDebugFile, "DAO version:" + m_DaoDbEngine.Version + "\r\n");
                this.WriteText(m_strDebugFile, "DAO Default Type:" + m_DaoDbEngine.DefaultType.ToString() + "\r\n");
                Type oType = m_DaoDbEngine.GetType();
                this.WriteText(m_strDebugFile, "GetType.AssemblyQualifiedName=" + oType.AssemblyQualifiedName.ToString() + "\r\n");
                this.WriteText(m_strDebugFile, "GetType.Name=" + oType.Name.ToString() + "\r\n");
                this.WriteText(m_strDebugFile, "GetType.GUID=" + oType.GUID.ToString() + "\r\n");
                this.WriteText(m_strDebugFile, "GetType.Assembly.FullName=" + oType.Assembly.FullName.ToString() + "\r\n");
                this.WriteText(m_strDebugFile, "GetType.Assembly.Location=" + oType.Assembly.Location.ToString() + "\r\n");
                this.WriteText(m_strDebugFile, "GetType.Assembly.CodeBase=" + oType.Assembly.CodeBase.ToString() + "\r\n");
            }
        }
        ~dao_data_access()
        {
            if (m_bDebug)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//dao_data_access.destructor\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            try
            {
                if (this.m_DaoWorkspace != null) 
                    this.m_DaoWorkspace.Close();
            }
            catch
            {
                if (this.m_DaoWorkspace != null)
                    this.m_DaoWorkspace = null;
            }

            this.m_DaoWorkspace = null;

        }
        /// <summary>
        /// open a mdb file and load the result into database variable m_DaoDatabase
        /// </summary>
        /// <param name="strFullPath">full path and mdb file name</param>
        public void OpenDb(string strFullPath)
        {
            Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabase;
            this.m_strError = "";
            this.m_intError = 0;
            try
            {
                p_DaoDatabase = this.m_DaoDbEngine.Workspaces[0].OpenDatabase(strFullPath, false, false, string.Empty);
                
            }
            catch (Exception ex)
            {
                this.m_strError = "DAO Error Opening Database " + strFullPath.ToString();
                this.m_intError = -1;
                MessageBox.Show(this.m_strError);
                return;
            }
            this.m_DaoDatabase = p_DaoDatabase;
        }
        /// <summary>
        /// see if a table exists in an mdb file
        /// </summary>
        /// <param name="p_DaoDatabase">currently open MDB database file</param>
        /// <param name="strTable">name of the table</param>
        /// <returns>return true if the table exists</returns>
        public bool TableExists(Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabase, string strTable)
        {
            bool p_bResult = false;

            foreach (Microsoft.Office.Interop.Access.Dao.TableDef table in p_DaoDatabase.TableDefs)
            {
                if (table.Name.ToLower() == strTable.ToLower())
                {
                    p_bResult = true;
                    break;
                }

            }
            return p_bResult;
        }
        /// <summary>
        /// see if a table exists in an mdb file
        /// </summary>
        /// <param name="strMDBPathAndFile">full path and MDB file name</param>
        /// <param name="strTable">name of the table</param>
        /// <returns></returns>
        public bool TableExists(string strMDBPathAndFile, string strTable)
        {
            bool p_bResult = false;
            //int z = 0;
            this.m_intError = 0;
            this.m_strError = "";
            this.OpenDb(strMDBPathAndFile);
            if (this.m_intError == 0)
            {
                foreach (Microsoft.Office.Interop.Access.Dao.TableDef table in this.m_DaoDatabase.TableDefs)
                {
                    if (table.Name.ToLower() == strTable.ToLower())
                    {
                        p_bResult = true;
                        break;
                    }
                }
                this.m_DaoDatabase.Close();
            }

            return p_bResult;
        }
        public bool RenameTable(Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabase, string p_strTable, 
                                string p_strNewTable, bool p_bDelete, bool p_bRefreshLink)
        {

            this.m_intError = 0;
            this.m_strError = "";


            try
            {
                //see if the p_strNewTable already exists
                if (TableExists(m_DaoDatabase, p_strNewTable))
                {
                    if (p_bDelete)
                    {
                        if (!DeleteTableFromMDB(m_DaoDatabase, p_strNewTable)) return false;
                    }
                    else return false;


                }
                for (int x = 0; x <= m_DaoDatabase.TableDefs.Count - 1; x++)
                {
                    if (m_DaoDatabase.TableDefs[x].Name.ToLower().Trim() == p_strTable.ToLower().Trim())
                    {
                        m_DaoDatabase.TableDefs[x].Name = p_strNewTable.Trim();
                        if (p_bRefreshLink)
                            m_DaoDatabase.TableDefs[x].RefreshLink();
                        break;
                    }
                }


            }
            catch (Exception err)
            {
                m_intError = -1;
                m_strError = err.Message;
                MessageBox.Show(this.m_strError);
                return false;
            }

            return true;


        }
        public bool RenameTable(string strMDBPathAndFile, string p_strTable, string p_strNewTable, 
                                bool p_bDelete, bool p_bRefreshLink)
        {

            this.m_intError = 0;
            this.m_strError = "";
            bool bReturn = false;

            this.OpenDb(strMDBPathAndFile);

            if (m_intError != 0) return false;

            bReturn = RenameTable(m_DaoDatabase, p_strTable, p_strNewTable, p_bDelete, p_bRefreshLink);

            this.m_DaoDatabase.Close();
            this.m_DaoDatabase = null;

            return bReturn;






        }
        /// <summary>
        /// create a link to a table
        /// </summary>
        /// <param name="strMDBFileDestination">mdb file that will contain the table link</param>
        /// <param name="strTableDestination">table link name</param>
        /// <param name="strMDBFileSource">mdb file that contains the table to be linked</param>
        /// <param name="strTableSource">name of the table</param>
        public void CreateTableLink(string strMDBFileDestination, string strTableDestination, string strMDBFileSource, string strTableSource)
        {
            string strConnect = ";DATABASE=" + strMDBFileSource;
            this.m_intError = 0;
            this.m_strError = "";
            /********************************************************
             **open the destination file that will contain the link
             ********************************************************/
            this.OpenDb(strMDBFileDestination);
            if (this.m_intError == 0)
            {
                try
                {
                    /****************************************************
                     **get the table definition of the source table
                     ****************************************************/
                    this.m_DaoTableDef = this.m_DaoDatabase.CreateTableDef(strTableDestination, 0, strTableSource, strConnect);

                    /******************************************************************
                     **append the link to the source table in the destination file
                     ******************************************************************/
                    this.m_DaoDatabase.TableDefs.Append(this.m_DaoTableDef);

                }
                catch (Exception caught)
                {
                    this.m_intError = -1;
                    this.m_strError = caught.Message;
                    MessageBox.Show(this.m_strError);
                    return;
                }
                this.m_DaoDatabase.Close();
                this.m_DaoTableDef = null;

            }
            this.m_DaoDatabase = null;
        }
        /// <summary>
        ///create a link to a table 
        /// </summary>
        /// <param name="strMDBFileDestination">mdb file that will contain the table link</param>
        /// <param name="strTableDestination">table link name</param>
        /// <param name="strMDBFileSource   ">mdb file that contains the table to be linked</param>
        /// <param name="strTableSource">name of the table</param>
        /// <param name="bDeleteExistingLink">if true then delete an existing link that has the same name as specified in strTableDestinaion</param>
        public void CreateTableLink(string strMDBFileDestination, string strTableDestination, string strMDBFileSource, string strTableSource, bool bDeleteExistingLink)
        {
            string strConnect = ";DATABASE=" + strMDBFileSource;
            this.m_intError = 0;
            this.m_strError = "";
            /********************************************************
             **open the destination file that will contain the link
             ********************************************************/
            this.OpenDb(strMDBFileDestination);
            if (this.m_intError == 0)
            {
                try
                {
                    if (this.TableExists(this.m_DaoDatabase, strTableDestination) == true)
                    {
                        if (bDeleteExistingLink == true)
                        {
                            this.m_DaoDatabase.TableDefs.Delete(strTableDestination);
                        }

                    }

                    /****************************************************
                     **get the table definition of the source table
                     ****************************************************/
                    this.m_DaoTableDef = this.m_DaoDatabase.CreateTableDef(strTableDestination, 0, strTableSource, strConnect);

                    /******************************************************************
                     **append the link to the source table in the destination file
                     ******************************************************************/
                    this.m_DaoDatabase.TableDefs.Append(this.m_DaoTableDef);

                }
                catch (Exception caught)
                {
                    this.m_intError = -1;
                    this.m_strError = caught.Message;
                    MessageBox.Show(this.m_strError);
                    return;
                }
                this.m_DaoDatabase.Close();
                this.m_DaoTableDef = null;

            }
            this.m_DaoDatabase = null;
        }
        /// <summary>
        ///create a link to a table 
        /// </summary>
        /// <param name="strMDBFileDestination">mdb file that will contain the table link</param>
        /// <param name="strTableDestination">table link name</param>
        /// <param name="strMDBFileSource   ">mdb file that contains the table to be linked</param>
        /// <param name="strTableSource">name of the table</param>
        /// <param name="bDeleteExistingLink">if true then delete an existing link that has the same name as specified in strTableDestinaion</param>
        public void CreateTableLink(Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabaseDestination, string strTableDestination, string strMDBFileSource, string strTableSource, bool bDeleteExistingLink)
        {
            string strConnect = ";DATABASE=" + strMDBFileSource;
            this.m_intError = 0;
            this.m_strError = "";
            /********************************************************
             **open the destination file that will contain the link
             ********************************************************/

            try
            {
                if (this.TableExists(p_DaoDatabaseDestination, strTableDestination) == true)
                {
                    if (bDeleteExistingLink == true)
                    {
                        p_DaoDatabaseDestination.TableDefs.Delete(strTableDestination);
                    }

                }

                /****************************************************
                 **get the table definition of the source table
                 ****************************************************/
                this.m_DaoTableDef = p_DaoDatabaseDestination.CreateTableDef(strTableDestination, 0, strTableSource, strConnect);

                /******************************************************************
                 **append the link to the source table in the destination file
                 ******************************************************************/
                p_DaoDatabaseDestination.TableDefs.Append(this.m_DaoTableDef);

            }
            catch (Exception caught)
            {
                this.m_intError = -1;
                this.m_strError = caught.Message;
                MessageBox.Show(this.m_strError);
                return;
            }
        }

        /// <summary>
        /// create a link to a table
        /// </summary>
        /// <param name="p_DaoDatabaseDestination">open mdb destination file that will contain the table link</param>
        /// <param name="strTableDestination">table link name</param>
        /// <param name="strMDBFileSource">mdb file that contains the table to be linked</param>
        /// <param name="strTableSource">name of the table</param>
        public void CreateTableLink(Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabaseDestination, string strTableDestination, string strMDBFileSource, string strTableSource)
        {
            string strConnect = ";DATABASE=" + strMDBFileSource;

            this.m_intError = 0;
            this.m_strError = "";

            try
            {
                this.m_DaoTableDef = p_DaoDatabaseDestination.CreateTableDef(strTableDestination, 0, strTableSource, strConnect);
                p_DaoDatabaseDestination.TableDefs.Append(this.m_DaoTableDef);
            }
            catch (Exception caught)
            {
                this.m_intError = -1;
                this.m_strError = caught.Message;
                MessageBox.Show(this.m_strError);
                return;
            }
            this.m_DaoTableDef = null;
        }
        /// <summary>
        /// Create tables links from one mdb file to another
        /// </summary>
        /// <param name="p_strDbFileDestination"></param>
        /// <param name="p_strDbFileSource"></param>
        public void CreateTableLinks(string p_strDbFileDestination, string p_strDbFileSource)
        {
            string[] strTableNames;
            strTableNames = new string[1];

            int intCount = getTableNames(p_strDbFileSource, ref strTableNames);
            if (m_intError == 0)
            {
                if (intCount > 0)
                {
                    for (int x = 0; x <= intCount - 1; x++)
                    {

                        CreateTableLink(p_strDbFileDestination, strTableNames[x], p_strDbFileSource, strTableNames[x]);
                        if (m_intError != 0)
                        {
                            m_strError = strTableNames[x] + " !!Error Creating Table Link!!!";
                            break;
                        }
                    }

                }
            }
        }
        /// <summary>
        /// Determine the type of access table whether the table is an
        /// external link to an oracle table, external link to an access table, or
        /// a regular access table.
        /// The routine will return one of these values:
        /// O - Oracle linked table
        /// L - Access linked table
        /// T - Regular table
        /// </summary>
        /// <param name="strMDBPathAndFile"></param>
        /// <param name="strTable"></param>
        /// <returns></returns>
        public string TableType(string strMDBPathAndFile, string strTable)
        {
            string strConnect = "";

            this.OpenDb(strMDBPathAndFile);
            if (this.m_intError == 0)
            {
                for (int x = 0; x <= this.m_DaoDatabase.TableDefs.Count - 1; x++)
                {
                    if (this.m_DaoDatabase.TableDefs[x].Name.Trim().ToLower() == strTable.Trim().ToLower())
                    {
                        strConnect = this.m_DaoDatabase.TableDefs[x].Connect.ToString().Trim();

                        if (strConnect.ToUpper().IndexOf("ODBC;", 0) >= 0)
                        {
                            return "LINKEDODBC";
                        }
                        if (strConnect.Trim().Length > 0)
                        {
                            return "L";
                        }

                    }
                }
            }
            return "T";
        }
        /// <summary>
        /// get access table link properties
        /// </summary>
        /// <param name="strMDBPathAndFile">full path and mdb file name</param>
        /// <param name="strTable">table name in the strMDBPathAndFile param</param>
        /// <param name="strSourceDatabase">Source database location of the linked table</param>
        /// <param name="strSourceTable">table name found in the source database</param>
        /// <returns>gives a value to the strSourceDatabase and strSourceTable variables</returns>
        public void GetLinkedTableInfo(string strMDBPathAndFile, string strTable, ref string strSourceDatabase, ref string strSourceTable)
        {
            string strConnect = "";
            strSourceDatabase = "";
            strSourceTable = "";
            this.OpenDb(strMDBPathAndFile);
            if (this.m_intError == 0)
            {
                for (int x = 0; x <= this.m_DaoDatabase.TableDefs.Count - 1; x++)
                {

                    if (this.m_DaoDatabase.TableDefs[x].Name.Trim().ToLower() == strTable.Trim().ToLower())
                    {
                        strConnect = this.m_DaoDatabase.TableDefs[x].Connect.ToString().Trim();
                        if (strConnect.Trim().Length > 0)
                        {
                            strSourceTable = Convert.ToString(this.m_DaoDatabase.TableDefs[x].SourceTableName);
                            int intStart = strConnect.ToUpper().IndexOf("DATABASE=", 0);
                            if (intStart >= 0)
                            {
                                int intEnd = strConnect.IndexOf(";", intStart);
                                if (intEnd >= 0)
                                {
                                    strSourceDatabase = strConnect.Substring(intStart + 9, (intEnd + 1) - (intStart + 9));
                                }
                                else if (intEnd < 0)
                                {
                                    strSourceDatabase = strConnect.Substring(intStart + 9, strConnect.Length - (intStart + 9));
                                }
                            }
                        }
                        else
                        {
                            strSourceDatabase = "";
                            strSourceTable = "";

                        }
                        break;
                    }
                }
            }
        }
        public void CreateOracleTableLink(string strMDBFileDestination, string strTableSource, string strTableDestination, string strDsn, string strTableOwner)
        {

            string strConnect = "ODBC;DSN=" + strDsn + ";DBQ=" + strDsn + ";DATABASE=";

            this.m_intError = 0;
            this.m_strError = "";
            /********************************************************
             **open the destination file that will contain the link
             ********************************************************/
            this.OpenDb(strMDBFileDestination);
            if (this.m_intError == 0)
            {
                try
                {
                    /****************************************************
                     **get the table definition of the source table
                     ****************************************************/
                    this.m_DaoTableDef = this.m_DaoDatabase.CreateTableDef(strTableDestination, 0, strTableOwner.Trim() + "." + strTableSource.Trim(), strConnect);


                    /******************************************************************
                     **append the link to the source table in the destination file
                     ******************************************************************/
                    this.m_DaoDatabase.TableDefs.Append(this.m_DaoTableDef);

                }
                catch (Exception caught)
                {
                    m_DaoDatabase.Close();
                    m_DaoTableDef = null;
                    m_DaoDatabase = null;
                    this.m_intError = -1;
                    this.m_strError = caught.Message;
                    if (_bDisplayErrors)
                        MessageBox.Show(this.m_strError);
                    return;
                }
                this.m_DaoDatabase.Close();
                this.m_DaoTableDef = null;

            }
            this.m_DaoDatabase = null;
        }
        /// <summary>
        /// Create an Oracle XE Table Link
        /// </summary>
        /// <param name="strDsn">ODBC Data Source Name</param>
        /// <param name="strUserId"></param>
        /// <param name="strPW"></param>
        /// <param name="strTableSchema"></param>
        /// <param name="strTableSource"></param>
        /// <param name="strMDBFileDestination"></param>
        /// <param name="strTableDestination"></param>

        public void CreateOracleXETableLink(Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabase,string strDsn, string strUserId, string strPW, string strTableSchema, string strTableSource, string strTableDestination)
        {

            string strConnect = "ODBC;DSN=" + strDsn + ";UID=" + strUserId + ";PWD=" + strPW + ";DBQ=XE;";
            //string strConnect = "ODBC;DSN=FIA Biosum Oracle Services;UID=FCS;PWD=fcs;DBQ=XE;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;";
            this.m_intError = 0;
            this.m_strError = "";
            /********************************************************
             **open the destination file that will contain the link
             ********************************************************/
          
                try
                {
                    /****************************************************
                     **get the table definition of the source table
                     ****************************************************/
                    this.m_DaoTableDef = p_DaoDatabase.CreateTableDef(strTableDestination, 0, strTableSchema.Trim() + "." + strTableSource.Trim(), strConnect);


                    /******************************************************************
                     **append the link to the source table in the destination file
                     ******************************************************************/
                    p_DaoDatabase.TableDefs.Append(this.m_DaoTableDef);

                }
                catch (Exception caught)
                {
                    m_DaoTableDef = null;
                    this.m_intError = -1;
                    this.m_strError = caught.Message;
                    if (_bDisplayErrors)
                        MessageBox.Show(this.m_strError);
                    return;
                }
                this.m_DaoTableDef = null;
        }
        public void EditOracleTableLinkConnectionString(string strMDBPathAndFile, string strTable)
        {

            //int z = 0;
            string strConnect = "";
            this.m_intError = 0;
            this.m_strError = "";
            this.OpenDb(strMDBPathAndFile);
            int intCount = 0;
            if (this.m_intError == 0)
            {
                for (int x = 0; x <= this.m_DaoDatabase.TableDefs.Count - 1; x++)
                {
                    if (this.m_DaoDatabase.TableDefs[x].Name.Trim().ToLower() == strTable.Trim().ToLower())
                    {
                        strConnect = this.m_DaoDatabase.TableDefs[x].Connect.ToString().Trim();

                        if (strConnect.ToUpper().IndexOf("IFALLSUCCESSFUL", 0) > 0)
                        {
                            if (strConnect.ToUpper().IndexOf("UID", 0) > 0)
                            {
                            }
                            else
                            {
                                MessageBox.Show("before:" + strConnect);
                                strConnect = strConnect + "UID=nims_club;PWD=dteam;";

                                this.m_DaoDatabase.TableDefs[x].Connect = strConnect;
                                this.m_DaoDatabase.TableDefs[x].RefreshLink();

                                strConnect = this.m_DaoDatabase.TableDefs[x].Connect.ToString().Trim();
                                MessageBox.Show("after:" + strConnect);



                            }
                        }
                    }

                }
                if (this.m_DaoDatabase.Updatable) MessageBox.Show("true");
                else MessageBox.Show("false");

                this.m_DaoDatabase.TableDefs.Refresh();
                this.m_DaoDatabase.Close();
            }
        }
        public void CreateOracleXETableLink(string strDsn, string strUserId, string strPW, string strTableSchema, string strTableSource, string strMDBFileDestination, string strTableDestination)
        {
            if (m_bDebug)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//dao_data_access.CreateOracleXETableLink\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n\r\n");

                frmMain.g_oUtils.WriteText(m_strDebugFile, "Passed Parameters\r\n------------------\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "DSN=" + strDsn + "\r\nUserId=" + strUserId + "\r\nPassword=" + strPW + "\r\nOracleTableSchema=" + strTableSchema + "\r\nOracleTableSource=" + strTableSource + "\r\nDestinationMDBFile=" + strMDBFileDestination + "\r\nMDBTableLinkName=" + strTableDestination + "\r\n");
            }
            string strConnect = "ODBC;DSN=" + strDsn + ";UID=" + strUserId + ";PWD=" + strPW + ";DBQ=XE;";
            if (m_bDebug) frmMain.g_oUtils.WriteText(m_strDebugFile, "odbc connection string=" + strConnect + "\r\n");
            this.m_intError = 0;
            this.m_strError = "";
            /********************************************************
             **open the destination file that will contain the link
             ********************************************************/
            this.OpenDb(strMDBFileDestination);
            if (this.m_intError == 0)
            {
                try
                {
                    /****************************************************
                     **get the table definition of the source table
                     ****************************************************/
                    this.m_DaoTableDef = this.m_DaoDatabase.CreateTableDef(strTableDestination, 0, strTableSchema.Trim() + "." + strTableSource.Trim(), strConnect);


                    /******************************************************************
                     **append the link to the source table in the destination file
                     ******************************************************************/
                    this.m_DaoDatabase.TableDefs.Append(this.m_DaoTableDef);

                }
                catch (Exception caught)
                {
                    m_DaoDatabase.Close();
                    m_DaoTableDef = null;
                    m_DaoDatabase = null;
                    this.m_intError = -1;
                    this.m_strError = caught.Message;
                    if (_bDisplayErrors)
                        MessageBox.Show(this.m_strError);
                    return;
                }
                this.m_DaoDatabase.Close();
                this.m_DaoTableDef = null;

            }
            this.m_DaoDatabase = null;
        }

        public void CreateSQLiteTableLink(string strSQLiteFileDestination, string strTableSource, string strTableDestination, string strDsn, 
            string strSQLiteDbFile, bool bDeleteExistingLink)
        {

            string strConnect = "ODBC;DSN=" + strDsn + ";DATABASE=" + strSQLiteDbFile;

            this.m_intError = 0;
            this.m_strError = "";
            /********************************************************
             **open the destination file that will contain the link
             ********************************************************/
            this.OpenDb(strSQLiteFileDestination);
            if (this.m_intError == 0)
            {
                try
                {
                    if (this.TableExists(this.m_DaoDatabase, strTableDestination) == true)
                    {
                        if (bDeleteExistingLink == true)
                        {
                            this.m_DaoDatabase.TableDefs.Delete(strTableDestination);
                        }
                    }

                    /****************************************************
                     **get the table definition of the source table
                     ****************************************************/
                    this.m_DaoTableDef = this.m_DaoDatabase.CreateTableDef(strTableDestination, 0, strTableSource.Trim(), strConnect);


                    /******************************************************************
                     **append the link to the source table in the destination file
                     ******************************************************************/
                    this.m_DaoDatabase.TableDefs.Append(this.m_DaoTableDef);

                }
                catch (Exception caught)
                {
                    m_DaoDatabase.Close();
                    m_DaoTableDef = null;
                    m_DaoDatabase = null;
                    this.m_intError = -1;
                    this.m_strError = caught.Message;
                    if (_bDisplayErrors)
                        MessageBox.Show(this.m_strError);
                    return;
                }
                this.m_DaoDatabase.Close();
                this.m_DaoTableDef = null;

            }
            this.m_DaoDatabase = null;
        }

        public void CreateSQLiteTableLink(string strSQLiteFileDestination, string strTableSource, string strTableDestination, string strDsn, string strSQLiteDbFile)
        {

            string strConnect = "ODBC;DSN=" + strDsn + ";DATABASE=" + strSQLiteDbFile;

            this.m_intError = 0;
            this.m_strError = "";
            /********************************************************
             **open the destination file that will contain the link
             ********************************************************/
            this.OpenDb(strSQLiteFileDestination);
            if (this.m_intError == 0)
            {
                try
                {
                    /****************************************************
                     **get the table definition of the source table
                     ****************************************************/
                    this.m_DaoTableDef = this.m_DaoDatabase.CreateTableDef(strTableDestination, 0, strTableSource.Trim(), strConnect);


                    /******************************************************************
                     **append the link to the source table in the destination file
                     ******************************************************************/
                    this.m_DaoDatabase.TableDefs.Append(this.m_DaoTableDef);

                }
                catch (Exception caught)
                {
                    m_DaoDatabase.Close();
                    m_DaoTableDef = null;
                    m_DaoDatabase = null;
                    this.m_intError = -1;
                    this.m_strError = caught.Message;
                    if (_bDisplayErrors)
                        MessageBox.Show(this.m_strError);
                    return;
                }
                this.m_DaoDatabase.Close();
                this.m_DaoTableDef = null;

            }
            this.m_DaoDatabase = null;
        }

        /// <summary>
        /// delete a table from an mdb file
        /// </summary>
        /// <param name="strMDBPathAndFile">full path and mdb file name</param>
        /// <param name="strTable">table name to delete</param>
        public void DeleteTableFromMDB(string strMDBPathAndFile, string strTable)
        {
            this.OpenDb(strMDBPathAndFile);
            if (this.m_intError == 0)
            {
                //see if the table name already exists within the destination mdb 
                if (this.TableExists(this.m_DaoDatabase, strTable) == true)
                {
                    this.m_DaoDatabase.TableDefs.Delete(strTable);
                }

            }
            this.m_DaoDatabase.Close();
            this.m_DaoDatabase = null;

        }
        public bool DeleteTableFromMDB(Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabase, string strTable)
        {
            bool p_bResult = false;

            //see if the table name already exists within the destination mdb 
            if (this.TableExists(this.m_DaoDatabase, strTable) == true)
            {
                this.m_DaoDatabase.TableDefs.Delete(strTable);
            }
            return p_bResult;
        }

        public void MoveTableToMDB2(Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabaseDestination, string strTableDestination, string strMDBFileSource, string strTableSource)
        {
            this.CreateTableLink(p_DaoDatabaseDestination,
                strTableDestination,
                strMDBFileSource,
                strTableSource);
            if (this.m_intError == 0)
            {

            }
        }


        /// <summary>
        /// move a table to a different mdb file
        /// </summary>
        /// <param name="p_uc_datasource_edit">datasource user control</param>
        /// <param name="lDeleteTableSource">if true then delete table source once copy is finished</param>
        public void MoveTableToMDB(FIA_Biosum_Manager.uc_datasource_edit p_uc_datasource_edit,
                                   bool lDeleteTableSource)
        {

            string strMDBFileDestination = p_uc_datasource_edit.lblNewMDBFile.Text.Trim();
            string strTableDestination = p_uc_datasource_edit.lblNewTable.Text.Trim();
            string strMDBFileSource = p_uc_datasource_edit.lblMDBFile.Text.Trim();
            string strTableSource = p_uc_datasource_edit.lblTable.Text.Trim();
            System.Windows.Forms.ProgressBar p_pbar = p_uc_datasource_edit._ProgressBar;
            System.Windows.Forms.Label lblProgress = p_uc_datasource_edit._lblProgress;
            Datasource p_datasource;
            if (p_uc_datasource_edit.m_strScenarioId.Trim().Length > 0)
            {
                p_datasource = new Datasource(p_uc_datasource_edit.m_strProjectDirectory, p_uc_datasource_edit.m_strScenarioId);
            }
            else
            {
                p_datasource = new Datasource(p_uc_datasource_edit.m_strProjectDirectory);
            }

            Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabaseSource;
            Microsoft.Office.Interop.Access.Dao.Field p_DaoFieldDestination;
            Microsoft.Office.Interop.Access.Dao.Recordset p_DaoRsSource;
            Microsoft.Office.Interop.Access.Dao.Recordset p_DaoRsDestination;


            int x = 0;
            int intFieldSize = 0;
            this.m_intError = 0;
            this.m_strError = "";


            //open destination MDB file that the table will be moved to
            this.OpenDb(strMDBFileDestination);
            if (this.m_intError == 0)
            {
                //see if the table name already exists within the destination mdb 
                if (this.TableExists(this.m_DaoDatabase, strTableDestination) == true)
                {
                    //the table exists so lets add a number onto the
                    //table name and see if it exists
                    for (x = 1; x <= 99; x++)
                    {
                        strTableDestination = strTableSource.Trim() + x.ToString().Trim();
                        if (this.TableExists(this.m_DaoDatabase, strTableDestination) == false)
                        {
                            break;
                        }
                    }
                    //i should check if table exists here but i am
                    //going to assume there will never be a 
                    //table name from 1 to 99
                }
                //open the source mdb file and table
                this.m_strError = "";
                this.m_intError = 0;
                try
                {
                    p_DaoDatabaseSource = this.m_DaoDbEngine.Workspaces[0].OpenDatabase(strMDBFileSource, false, false, string.Empty);
                }
                catch
                {
                    this.m_strError = "DAO Error Opening Database " + strMDBFileSource.ToString();
                    this.m_intError = -1;
                    MessageBox.Show(this.m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    p_datasource = null;
                    return;
                }

                try
                {
                    //create a table and structure in the destination mdb
                    this.m_DaoTableDef = this.m_DaoDatabase.CreateTableDef(strTableDestination, 0, "", "");
                    for (x = 0; x <= p_DaoDatabaseSource.TableDefs[strTableSource].Fields.Count - 1; x++)
                    {
                        //dao does not like dbDecimal data type so had to convert dbDecimal fields to double
                        if (p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Type !=
                            (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDecimal)
                        {
                            intFieldSize = DataWidth(p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x]);
                            p_DaoFieldDestination = this.m_DaoTableDef.CreateField(
                                p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Name,
                                p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Type,
                                intFieldSize);
                        }
                        else
                        {
                            p_DaoFieldDestination = this.m_DaoTableDef.CreateField(
                                p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Name,
                                (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDouble, 17);
                        }
                        //allow text field lengths to be zero
                        if (p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Type ==
                            (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText) p_DaoFieldDestination.AllowZeroLength = true;
                        this.m_DaoTableDef.Fields.Append(p_DaoFieldDestination);
                    }
                    this.m_DaoDatabase.TableDefs.Append(this.m_DaoTableDef);
                }
                catch (Exception Caught)
                {
                    this.m_intError = -1;
                    this.m_strError = Caught.Message;
                    MessageBox.Show(m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    this.m_DaoTableDef = null;
                    p_DaoDatabaseSource = null;
                    p_datasource = null;
                    return;
                }

                //okay now lets move the contents from the source table to the destination table
                try
                {

                    p_datasource.SetPrimaryIndexesAndAutoNumbers(p_uc_datasource_edit.lblTableType.Text.Trim(),
                        strTableDestination,
                        this);
                    p_DaoRsSource = p_DaoDatabaseSource.OpenRecordset(
                                            strTableSource,
                                            Microsoft.Office.Interop.Access.Dao.RecordsetTypeEnum.dbOpenTable,
                                            Microsoft.Office.Interop.Access.Dao.RecordsetOptionEnum.dbReadOnly,
                                            Microsoft.Office.Interop.Access.Dao.RecordsetOptionEnum.dbReadOnly);
                    p_DaoRsDestination = this.m_DaoDatabase.OpenRecordset(
                                            strTableDestination,
                                            Microsoft.Office.Interop.Access.Dao.RecordsetTypeEnum.dbOpenTable,
                                            Microsoft.Office.Interop.Access.Dao.RecordsetOptionEnum.dbSeeChanges,
                                            Microsoft.Office.Interop.Access.Dao.LockTypeEnum.dbPessimistic);
                    if (p_DaoRsSource.RecordCount > 0)
                    {
                        string strTtlRecs = p_DaoRsSource.RecordCount.ToString().Trim();
                        p_pbar.Maximum = p_DaoRsSource.RecordCount;
                        p_pbar.Minimum = 0;
                        p_pbar.Visible = true;
                        lblProgress.Visible = true;
                        p_DaoRsSource.MoveFirst();
                        int y = 0;
                        while (p_DaoRsSource.EOF == false)
                        {
                            y++;
                            p_pbar.Value = y;
                            p_pbar.Refresh();
                            string str = "Adding Record#" + y.ToString().Trim() + " Of " + strTtlRecs + "  Records";

                            lblProgress.Text = str;
                            lblProgress.Refresh();

                            p_DaoRsDestination.AddNew();
                            for (x = 0; x <= p_DaoRsDestination.Fields.Count - 1; x++)
                            {
                                p_DaoRsDestination.Fields[x].Value = p_DaoRsSource.Fields[x].Value;

                            }
                            p_DaoRsDestination.Update((int)Microsoft.Office.Interop.Access.Dao.UpdateTypeEnum.dbUpdateRegular, false);
                            p_DaoRsSource.MoveNext();
                        }
                    }
                }

                catch (Exception caught)
                {
                    this.m_intError = -1;
                    this.m_strError = caught.Message;
                    MessageBox.Show(this.m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    p_DaoDatabaseSource.Close();
                    p_DaoDatabaseSource = null;
                    this.m_DaoTableDef = null;
                    p_DaoFieldDestination = null;
                    p_DaoRsSource = null;
                    p_DaoRsDestination = null;
                    p_datasource = null;
                    return;
                }

                p_DaoRsDestination.Close();
                p_DaoRsDestination = null;
                p_DaoRsSource.Close();
                p_DaoRsSource = null;
                this.m_DaoTableDef = null;
                p_DaoFieldDestination = null;

                if (lDeleteTableSource == true)
                {
                    p_DaoDatabaseSource.TableDefs.Delete(strTableSource);
                }

                this.m_DaoDatabase.Close();
                this.m_DaoDatabase = null;
                p_DaoDatabaseSource.Close();
                p_DaoDatabaseSource = null;
                p_datasource = null;


            }

        }
        public void MoveTableToMDB(FIA_Biosum_Manager.uc_datasource_edit p_uc_datasource_edit,
                                   bool lDeleteTableSource, string strTemp)
        {
            string strMDBFileDestination = p_uc_datasource_edit.lblNewMDBFile.Text.Trim();
            string strTableDestination = p_uc_datasource_edit.lblNewTable.Text.Trim();
            string strMDBFileSource = p_uc_datasource_edit.lblMDBFile.Text.Trim();
            string strTableSource = p_uc_datasource_edit.lblTable.Text.Trim();
            System.Windows.Forms.ProgressBar p_pbar = p_uc_datasource_edit._ProgressBar;
            System.Windows.Forms.Label lblProgress = p_uc_datasource_edit._lblProgress;
            Datasource p_datasource;
            if (p_uc_datasource_edit.m_strScenarioId.Trim().Length > 0)
            {
                p_datasource = new Datasource(p_uc_datasource_edit.m_strProjectDirectory, p_uc_datasource_edit.m_strScenarioId);
            }
            else
            {
                p_datasource = new Datasource(p_uc_datasource_edit.m_strProjectDirectory);
            }

            Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabaseSource;
            Microsoft.Office.Interop.Access.Dao.Field p_DaoFieldDestination;
            Microsoft.Office.Interop.Access.Dao.Recordset p_DaoRsSource;
            Microsoft.Office.Interop.Access.Dao.Recordset p_DaoRsDestination;

            int x = 0;
            int intFieldSize = 0;
            this.m_intError = 0;
            this.m_strError = "";


            //open destination MDB file that the table will be moved to
            this.OpenDb(strMDBFileDestination);
            if (this.m_intError == 0)
            {
                this.m_strError = "";
                this.m_intError = 0;
                try
                {
                    p_DaoDatabaseSource = this.m_DaoDbEngine.Workspaces[0].OpenDatabase(strMDBFileSource, false, false, string.Empty);
                }
                catch
                {
                    this.m_strError = "DAO Error Opening Database " + strMDBFileSource.ToString();
                    this.m_intError = -1;
                    MessageBox.Show(this.m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    p_datasource = null;
                    return;
                }

                try
                {
                    //create a table and structure in the destination mdb
                    this.m_DaoTableDef = this.m_DaoDatabase.CreateTableDef(strTableDestination, 0, "", "");
                    for (x = 0; x <= p_DaoDatabaseSource.TableDefs[strTableSource].Fields.Count - 1; x++)
                    {
                        //dao does not like dbDecimal data type so had to convert dbDecimal fields to double
                        if (p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Type !=
                            (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDecimal)
                        {
                            intFieldSize = DataWidth(p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x]);
                            p_DaoFieldDestination = this.m_DaoTableDef.CreateField(
                                p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Name,
                                p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Type,
                                intFieldSize);
                        }
                        else
                        {
                            p_DaoFieldDestination = this.m_DaoTableDef.CreateField(
                                p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Name,
                                (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDouble, 17);
                        }
                        //allow text field lengths to be zero
                        if (p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Type ==
                            (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText) p_DaoFieldDestination.AllowZeroLength = true;
                        this.m_DaoTableDef.Fields.Append(p_DaoFieldDestination);
                    }
                    this.m_DaoDatabase.TableDefs.Append(this.m_DaoTableDef);
                }
                catch (Exception Caught)
                {
                    this.m_intError = -1;
                    this.m_strError = Caught.Message;
                    MessageBox.Show(m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    this.m_DaoTableDef = null;
                    p_DaoDatabaseSource = null;
                    p_datasource = null;
                    return;
                }

                //okay now lets move the contents from the source table to the destination table
                try
                {
                    //create primary indexes and autonumbers
                    p_datasource.SetPrimaryIndexesAndAutoNumbers(p_uc_datasource_edit.lblTableType.Text.Trim(),
                                                                 strTableSource,
                                                                 this);

                    p_DaoRsSource = p_DaoDatabaseSource.OpenRecordset(
                        strTableSource,
                        Microsoft.Office.Interop.Access.Dao.RecordsetTypeEnum.dbOpenTable,
                        Microsoft.Office.Interop.Access.Dao.RecordsetOptionEnum.dbReadOnly,
                        Microsoft.Office.Interop.Access.Dao.RecordsetOptionEnum.dbReadOnly);
                    p_DaoRsDestination = this.m_DaoDatabase.OpenRecordset(
                        strTableDestination,
                        Microsoft.Office.Interop.Access.Dao.RecordsetTypeEnum.dbOpenTable,
                        Microsoft.Office.Interop.Access.Dao.RecordsetOptionEnum.dbSeeChanges,
                        Microsoft.Office.Interop.Access.Dao.LockTypeEnum.dbPessimistic);
                    if (p_DaoRsSource.RecordCount > 0)
                    {
                        string strTtlRecs = p_DaoRsSource.RecordCount.ToString().Trim();
                        x = 0;
                        int y = 0;
                        p_pbar.Maximum = p_DaoRsSource.RecordCount;
                        p_pbar.Minimum = 0;
                        lblProgress.Visible = true;
                        p_pbar.Visible = true;
                        p_DaoRsSource.MoveFirst();
                        while (p_DaoRsSource.EOF == false)
                        {
                            y++;
                            p_pbar.Value = y;
                            p_pbar.Refresh();
                            string str = "Adding Record#" + y.ToString().Trim() + " Of " + strTtlRecs + "  Records";
                            lblProgress.Text = str;
                            lblProgress.Refresh();
                            p_DaoRsDestination.AddNew();
                            for (x = 0; x <= p_DaoRsDestination.Fields.Count - 1; x++)
                            {
                                p_DaoRsDestination.Fields[x].Value = p_DaoRsSource.Fields[x].Value;

                            }
                            p_DaoRsDestination.Update((int)Microsoft.Office.Interop.Access.Dao.UpdateTypeEnum.dbUpdateRegular, false);
                            p_DaoRsSource.MoveNext();

                        }
                    }

                }

                catch (Exception caught)
                {
                    this.m_intError = -1;
                    this.m_strError = caught.Message;
                    MessageBox.Show(this.m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    p_DaoDatabaseSource.Close();
                    p_DaoDatabaseSource = null;
                    this.m_DaoTableDef = null;
                    p_DaoFieldDestination = null;
                    p_DaoRsSource = null;
                    p_DaoRsDestination = null;
                    p_datasource = null;
                    return;
                }

                p_DaoRsDestination.Close();
                p_DaoRsDestination = null;
                p_DaoRsSource.Close();
                p_DaoRsSource = null;
                this.m_DaoTableDef = null;
                p_DaoFieldDestination = null;

                if (lDeleteTableSource == true)
                {
                    p_DaoDatabaseSource.TableDefs.Delete(strTableSource);
                }

                this.m_DaoDatabase.Close();
                this.m_DaoDatabase = null;
                p_DaoDatabaseSource.Close();
                p_DaoDatabaseSource = null;
                p_datasource = null;


            }

        }

        /// <summary>
        /// copy table to a different mdb file
        /// </summary>
        /// <param name="strMDBFileDestination">mdb destination</param>
        /// <param name="strTableDestination">name to give destination table</param>
        /// <param name="strMDBFileSource">mdb file to copy strTableSource</param>
        /// <param name="strTableSource">name of table</param>
        /// <param name="lDeleteTableSource">if true then delete table once copy is done</param>
        public void MoveTableToMDB(string strMDBFileDestination,
                                   string strTableDestination,
                                   string strMDBFileSource,
                                   string strTableSource,
                                   bool lDeleteTableSource)
        {
            Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabaseSource;
            Microsoft.Office.Interop.Access.Dao.Field p_DaoFieldDestination;
            Microsoft.Office.Interop.Access.Dao.Recordset p_DaoRsSource;
            Microsoft.Office.Interop.Access.Dao.Recordset p_DaoRsDestination;

            int x = 0;
            int intFieldSize = 0;
            this.m_intError = 0;
            this.m_strError = "";


            //open destination MDB file that the table will be moved to
            this.OpenDb(strMDBFileDestination);
            if (this.m_intError == 0)
            {
                this.m_strError = "";
                this.m_intError = 0;
                try
                {
                    p_DaoDatabaseSource = this.m_DaoDbEngine.Workspaces[0].OpenDatabase(strMDBFileSource, false, false, string.Empty);
                }
                catch
                {
                    this.m_strError = "DAO Error Opening Database " + strMDBFileSource.ToString();
                    this.m_intError = -1;
                    MessageBox.Show(this.m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    return;
                }

                try
                {
                    //create a table and structure in the destination mdb
                    this.m_DaoTableDef = this.m_DaoDatabase.CreateTableDef(strTableDestination, 0, "", "");
                    for (x = 0; x <= p_DaoDatabaseSource.TableDefs[strTableSource].Fields.Count - 1; x++)
                    {
                        //dao does not like dbDecimal data type so had to convert dbDecimal fields to double
                        if (p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Type !=
                            (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDecimal)
                        {
                            intFieldSize = DataWidth(p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x]);
                            p_DaoFieldDestination = this.m_DaoTableDef.CreateField(
                                p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Name,
                                p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Type,
                                intFieldSize);
                        }
                        else
                        {
                            p_DaoFieldDestination = this.m_DaoTableDef.CreateField(
                                p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Name,
                                (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDouble, 17);
                        }
                        //allow text field lengths to be zero
                        if (p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Type ==
                            (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText) p_DaoFieldDestination.AllowZeroLength = true;
                        this.m_DaoTableDef.Fields.Append(p_DaoFieldDestination);
                    }
                    this.m_DaoDatabase.TableDefs.Append(this.m_DaoTableDef);
                }
                catch (Exception Caught)
                {
                    this.m_intError = -1;
                    this.m_strError = Caught.Message;
                    MessageBox.Show(m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    this.m_DaoTableDef = null;
                    p_DaoDatabaseSource = null;
                    return;
                }

                //okay now lets move the contents from the source table to the destination table
                try
                {
                    p_DaoRsSource = p_DaoDatabaseSource.OpenRecordset(
                        strTableSource,
                        Microsoft.Office.Interop.Access.Dao.RecordsetTypeEnum.dbOpenTable,
                        Microsoft.Office.Interop.Access.Dao.RecordsetOptionEnum.dbReadOnly,
                        Microsoft.Office.Interop.Access.Dao.RecordsetOptionEnum.dbReadOnly);
                    p_DaoRsDestination = this.m_DaoDatabase.OpenRecordset(
                        strTableDestination,
                        Microsoft.Office.Interop.Access.Dao.RecordsetTypeEnum.dbOpenTable,
                        Microsoft.Office.Interop.Access.Dao.RecordsetOptionEnum.dbSeeChanges,
                        Microsoft.Office.Interop.Access.Dao.LockTypeEnum.dbPessimistic);
                    if (p_DaoRsSource.RecordCount > 0)
                    {

                        p_DaoRsSource.MoveFirst();
                        while (p_DaoRsSource.EOF == false)
                        {


                            p_DaoRsDestination.AddNew();
                            for (x = 0; x <= p_DaoRsDestination.Fields.Count - 1; x++)
                            {
                                p_DaoRsDestination.Fields[x].Value = p_DaoRsSource.Fields[x].Value;

                            }
                            p_DaoRsDestination.Update((int)Microsoft.Office.Interop.Access.Dao.UpdateTypeEnum.dbUpdateRegular, false);
                            p_DaoRsSource.MoveNext();

                        }
                    }

                }

                catch (Exception caught)
                {
                    this.m_intError = -1;
                    this.m_strError = caught.Message;
                    MessageBox.Show(this.m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    p_DaoDatabaseSource.Close();
                    p_DaoDatabaseSource = null;
                    this.m_DaoTableDef = null;
                    p_DaoFieldDestination = null;
                    p_DaoRsSource = null;
                    p_DaoRsDestination = null;
                    return;
                }

                p_DaoRsDestination.Close();
                p_DaoRsDestination = null;
                p_DaoRsSource.Close();
                p_DaoRsSource = null;
                this.m_DaoTableDef = null;
                p_DaoFieldDestination = null;

                if (lDeleteTableSource == true)
                {
                    p_DaoDatabaseSource.TableDefs.Delete(strTableSource);
                }

                this.m_DaoDatabase.Close();
                this.m_DaoDatabase = null;
                p_DaoDatabaseSource.Close();
                p_DaoDatabaseSource = null;


            }

        }


        /// <summary>
        /// copy table to a different MDB file
        /// </summary>
        /// <param name="strMDBFileDestination">mdb file destination of table</param>
        /// <param name="strMDBFileSource">mdb file containing the table source</param>
        /// <param name="strTableSource">table name</param>
        /// <param name="lDeleteTableSource">if true then delete the table once the copy is finished</param>
        public void MoveTableToMDB(string strMDBFileDestination, string strMDBFileSource, string strTableSource, bool lDeleteTableSource)
        {
            Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabaseSource;
            Microsoft.Office.Interop.Access.Dao.Field p_DaoFieldDestination;
            Microsoft.Office.Interop.Access.Dao.Recordset p_DaoRsSource;
            Microsoft.Office.Interop.Access.Dao.Recordset p_DaoRsDestination;

            string strTableDestination = strTableSource;

            int x = 0;
            int intFieldSize = 0;
            this.m_intError = 0;
            this.m_strError = "";


            //open destination MDB file that the table will be moved to
            this.OpenDb(strMDBFileDestination);
            if (this.m_intError == 0)
            {
                //see if the table name already exists within the destination mdb 
                if (this.TableExists(this.m_DaoDatabase, strTableDestination) == true)
                {
                    //the table exists so lets add a number onto the
                    //table name and see if it exists
                    for (x = 1; x <= 99; x++)
                    {
                        strTableDestination = strTableSource.Trim() + x.ToString().Trim();
                        if (this.TableExists(this.m_DaoDatabase, strTableDestination) == false)
                        {
                            break;
                        }
                    }
                    //i should check if table exists here but i am
                    //going to assume there will never be a 
                    //table name from 1 to 99
                }
                //open the source mdf file and table
                this.m_strError = "";
                this.m_intError = 0;
                try
                {
                    p_DaoDatabaseSource = this.m_DaoDbEngine.Workspaces[0].OpenDatabase(strMDBFileSource, false, false, string.Empty);
                }
                catch
                {
                    this.m_strError = "DAO Error Opening Database " + strMDBFileSource.ToString();
                    this.m_intError = -1;
                    MessageBox.Show(this.m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    return;
                }

                try
                {
                    //create a table and structure in the destination mdb
                    this.m_DaoTableDef = this.m_DaoDatabase.CreateTableDef(strTableDestination, 0, "", "");
                    for (x = 0; x <= p_DaoDatabaseSource.TableDefs[strTableSource].Fields.Count - 1; x++)
                    {
                        //dao does not like dbDecimal data type so had to convert dbDecimal fields to double
                        if (p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Type !=
                            (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDecimal)
                        {
                            intFieldSize = DataWidth(p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x]);
                            p_DaoFieldDestination = this.m_DaoTableDef.CreateField(
                                p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Name,
                                p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Type,
                                intFieldSize);
                        }
                        else
                        {
                            p_DaoFieldDestination = this.m_DaoTableDef.CreateField(
                                p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Name,
                                (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDouble, 17);
                        }
                        //allow text field lengths to be zero
                        if (p_DaoDatabaseSource.TableDefs[strTableSource].Fields[x].Type ==
                            (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText) p_DaoFieldDestination.AllowZeroLength = true;
                        this.m_DaoTableDef.Fields.Append(p_DaoFieldDestination);
                    }
                    this.m_DaoDatabase.TableDefs.Append(this.m_DaoTableDef);
                }
                catch (Exception Caught)
                {
                    this.m_intError = -1;
                    this.m_strError = Caught.Message;
                    MessageBox.Show(m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    this.m_DaoTableDef = null;
                    p_DaoDatabaseSource = null;
                    return;
                }

                //okay now lets move the contents from the source table to the destination table
                try
                {
                    p_DaoRsSource = p_DaoDatabaseSource.OpenRecordset(
                        strTableSource,
                        Microsoft.Office.Interop.Access.Dao.RecordsetTypeEnum.dbOpenTable,
                        Microsoft.Office.Interop.Access.Dao.RecordsetOptionEnum.dbReadOnly,
                        Microsoft.Office.Interop.Access.Dao.RecordsetOptionEnum.dbReadOnly);
                    p_DaoRsDestination = this.m_DaoDatabase.OpenRecordset(
                        strTableDestination,
                        Microsoft.Office.Interop.Access.Dao.RecordsetTypeEnum.dbOpenTable,
                        Microsoft.Office.Interop.Access.Dao.RecordsetOptionEnum.dbSeeChanges,
                        Microsoft.Office.Interop.Access.Dao.LockTypeEnum.dbPessimistic);
                    if (p_DaoRsSource.RecordCount > 0)
                    {
                        p_DaoRsSource.MoveFirst();
                        while (p_DaoRsSource.EOF == false)
                        {


                            p_DaoRsDestination.AddNew();
                            for (x = 0; x <= p_DaoRsDestination.Fields.Count - 1; x++)
                            {
                                p_DaoRsDestination.Fields[x].Value = p_DaoRsSource.Fields[x].Value;

                            }
                            p_DaoRsDestination.Update((int)Microsoft.Office.Interop.Access.Dao.UpdateTypeEnum.dbUpdateRegular, false);
                            p_DaoRsSource.MoveNext();
                        }
                    }
                }

                catch (Exception caught)
                {
                    this.m_intError = -1;
                    this.m_strError = caught.Message;
                    MessageBox.Show(this.m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    p_DaoDatabaseSource.Close();
                    p_DaoDatabaseSource = null;
                    this.m_DaoTableDef = null;
                    p_DaoFieldDestination = null;
                    p_DaoRsSource = null;
                    p_DaoRsDestination = null;
                    return;
                }

                p_DaoRsDestination.Close();
                p_DaoRsDestination = null;
                p_DaoRsSource.Close();
                p_DaoRsSource = null;
                this.m_DaoTableDef = null;
                p_DaoFieldDestination = null;

                if (lDeleteTableSource == true)
                {
                    p_DaoDatabaseSource.TableDefs.Delete(strTableSource);
                }

                this.m_DaoDatabase.Close();
                this.m_DaoDatabase = null;
                p_DaoDatabaseSource.Close();
                p_DaoDatabaseSource = null;


            }

        }

        /// <summary>
        /// remove fields from the table listed in the strFields array
        /// </summary>
        /// <param name="strMDBFile">full path and mdb file name</param>
        /// <param name="strTable">table name</param>
        /// <param name="strFields">string array containing list of field names to remove from table</param>
        public void DeleteField(string strMDBFile, string strTable, string[] strFields)
        {

            int x = 0;
            int y = 0;
            this.m_intError = 0;
            this.m_strError = "";


            //open destination MDB file that the table will be moved to
            this.OpenDb(strMDBFile);
            if (this.m_intError == 0)
            {


                //see if the table name already exists within the destination mdb 
                try
                {
                    if (this.TableExists(this.m_DaoDatabase, strTable) == true)
                    {
                        for (y = 0; y <= strFields.Length - 1; y++)
                        {
                            if (strFields[y].Trim().Length > 0)
                            {
                                for (x = 0; x <= this.m_DaoDatabase.TableDefs[strTable].Fields.Count - 1; x++)
                                {
                                    if (this.m_DaoDatabase.TableDefs[strTable].Fields[x].Name.ToString().Trim().ToUpper() ==
                                        strFields[y].Trim().ToUpper())
                                    {
                                        this.m_DaoDatabase.TableDefs[strTable].Fields.Delete(strFields[y].Trim());
                                    }
                                }
                            }
                        }
                    }
                    //open the source mdf file and table

                }
                catch (Exception Caught)
                {
                    this.m_intError = -1;
                    this.m_strError = Caught.Message;
                    MessageBox.Show(m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    this.m_DaoTableDef = null;
                    return;
                }
                this.m_DaoDatabase.Close();
                this.m_DaoDatabase = null;
                this.m_DaoTableDef = null;

            }
        }
        /// <summary>
        /// create an mdb file
        /// </summary>
        /// <param name="strFullPath">full path and name of mdb file</param>
        public void CreateMDB(string strFullPath)
        {
            Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabase;
            if (m_bDebug)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//dao_data_access.CreateMDB\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            this.m_intError = 0;
            this.m_strError = "";
            try
            {
                p_DaoDatabase = this.m_DaoDbEngine.Workspaces[0].CreateDatabase(strFullPath, Microsoft.Office.Interop.Access.Dao.LanguageConstants.dbLangGeneral, Microsoft.Office.Interop.Access.Dao.DatabaseTypeEnum.dbVersion120); //,false,string.Empty);
            }
            catch (Exception caught)
            {
                this.m_intError = -1;
                this.m_strError = caught.Message;
                MessageBox.Show(this.m_strError);
                p_DaoDatabase = null;
                return;
            }
            p_DaoDatabase.Close();
            p_DaoDatabase = null;
        }


        /// <summary>
        /// return the default field length depending on its data type
        /// </summary>
        /// <param name="field_def">dao field definition</param>
        /// <returns>dao field length</returns>
        private int DataWidth(Microsoft.Office.Interop.Access.Dao.Field field_def)
        {
            switch (field_def.Type)
            {
                case (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbByte:
                    return 3;
                case (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbInteger:
                    return 6;
                case (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbLong:
                    return 11;
                case 14:          //Auto Number
                    return 11;
                case (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbSingle:
                    return 9;
                case (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDouble:
                    return 17;
                case (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDecimal:
                    return 17;
                case (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText:
                    return field_def.Size;
                case (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbMemo:
                    return 4;
                case (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDate:
                    return 19;
                case (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbCurrency:
                    return 7;
                case (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbBoolean:
                    return 3;
                default:
                    return 3;
            }
        }


        /// <summary>
        /// loads the string array strTableNames with the table names found in the mdb file
        /// </summary>
        /// <param name="strMDBPathAndFile">full path and name of mdb file</param>
        /// <param name="strTableNames">table names in strMDBPathAndFile are loaded into this array</param>
        /// <returns></returns>
        public int getTableNames(string strMDBPathAndFile, ref string[] strTableNames)
        {
            int x = 0;
            this.m_intError = 0;
            this.m_strError = "";
            this.OpenDb(strMDBPathAndFile);
            if (this.m_intError == 0)
            {

                strTableNames = new string[this.m_DaoDatabase.TableDefs.Count];
                foreach (Microsoft.Office.Interop.Access.Dao.TableDef table in this.m_DaoDatabase.TableDefs)
                {
                    if (table.Name.ToString().ToUpper().IndexOf("MSYS") != 0)
                    {
                        strTableNames[x] = table.Name.ToString();
                        x++;
                    }

                }


                this.m_DaoDatabase.Close();

            }
            return x;
        }
        public int getTableNames(string strMDBPathAndFile, ref string[] strTableNames, bool bGetLinkedTableNames)
        {
            int x = 0;
            this.m_intError = 0;
            this.m_strError = "";
            this.OpenDb(strMDBPathAndFile);
            if (this.m_intError == 0)
            {

                strTableNames = new string[this.m_DaoDatabase.TableDefs.Count];
                foreach (Microsoft.Office.Interop.Access.Dao.TableDef table in this.m_DaoDatabase.TableDefs)
                {
                    if (table.Name.ToString().ToUpper().IndexOf("MSYS") != 0)
                    {
                        if (bGetLinkedTableNames)
                        {
                            strTableNames[x] = table.Name.ToString();
                            x++;
                        }
                        else
                        {
                            //MessageBox.Show(table.Name.ToString() + " " + table.Attributes.ToString());
                            if (table.Attributes == 0)
                            {
                                strTableNames[x] = table.Name.ToString();
                                x++;
                            }
                        }
                    }
                }



                this.m_DaoDatabase.Close();

            }
            return x;
        }
        /// <summary>
        /// get a table field names and load them into the string array strFieldNames
        /// </summary>
        /// <param name="strMDBPathAndFile">full path and mdb file name</param>
        /// <param name="strTableName">table name</param>
        /// <param name="strFieldNames">string array</param>
        public void getFieldNames(string strMDBPathAndFile, string strTableName, ref string[] strFieldNames)
        {


            int x = 0;
            this.m_intError = 0;
            this.m_strError = "";


            //open destination MDB file that the table will be moved to
            this.OpenDb(strMDBPathAndFile);
            if (this.m_intError == 0)
            {
                try
                {
                    strFieldNames = new string[this.m_DaoDatabase.TableDefs[strTableName].Fields.Count];
                    for (x = 0; x <= this.m_DaoDatabase.TableDefs[strTableName].Fields.Count - 1; x++)
                    {
                        strFieldNames[x] = this.m_DaoDatabase.TableDefs[strTableName].Fields[x].Name;
                    }
                }
                catch (Exception caught)
                {
                    this.m_strError = "dao_data_access.getFieldNames() error: " + caught.Message;
                    this.m_intError = -1;
                    MessageBox.Show(this.m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    return;
                }
                this.m_DaoDatabase.Close();
                this.m_DaoDatabase = null;
            }

        }
        /// <summary>
        /// get a table field names and load them into the string array strFieldNames
        /// </summary>
        /// <param name="strMDBPathAndFile">full path and mdb file name</param>
        /// <param name="strTableName">table name</param>
        /// <param name="strFieldNames">string array</param>
        public void getFieldNames(Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabase, string strTableName, ref string[] strFieldNames)
        {


            int x = 0;
            this.m_intError = 0;
            this.m_strError = "";


            try
            {
                strFieldNames = new string[p_DaoDatabase.TableDefs[strTableName].Fields.Count];
                for (x = 0; x <= p_DaoDatabase.TableDefs[strTableName].Fields.Count - 1; x++)
                {
                    strFieldNames[x] = p_DaoDatabase.TableDefs[strTableName].Fields[x].Name;
                }
            }
            catch (Exception caught)
            {
                this.m_strError = "dao_data_access.getFieldNames() error: " + caught.Message;
                this.m_intError = -1;
                MessageBox.Show(this.m_strError);
                return;
            }
        }
        public bool ColumnExist(Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabase, string p_strTableName, string p_strColumnName)
        {
            int x = 0;
            this.m_intError = 0;
            this.m_strError = "";
            string[] strColumns = null;
            bool bFound = false;
            if (TableExists(p_DaoDatabase, p_strTableName))
            {
                getFieldNames(p_DaoDatabase, p_strTableName, ref strColumns);
                if (strColumns != null)
                {
                    for (x = 0; x <= strColumns.Length - 1; x++)
                    {
                        if (strColumns[x] != null && strColumns[x].Trim().Length > 0)
                        {
                            if (strColumns[x].Trim().ToUpper() ==
                                p_strColumnName.Trim().ToUpper())
                            {
                                bFound = true;
                                break;
                            }
                        }
                    }
                }
            }
            else
            {
            }
            return bFound;
        }

        /**********************************************************
         **create a table from an ado oledb datatable and return
         **the name of the table to the client
         **********************************************************/
        /// <summary>
        /// create an MDB Access table from an ado oledb datatable and return
        /// the name of the table to the client
        /// </summary>
        /// <param name="strMDBPathAndFile">directory location and name of the mdb file to create the table</param>
        /// <param name="p_strMDBTableName">table name</param>
        /// <param name="p_AdoOleDbDataTable">table containing rows that describe the table schema</param>
        /// <param name="p_bDeleteExistingTable">True: If the table name in variable p_strMDBTable exists then delete it. False: Do not delete the table. Append numbers to the end of p_strMDBTableName until a unique table name is found</param>
        /// <returns></returns>
        public string CreateMDBTableFromDataSetTable(string strMDBPathAndFile, string p_strMDBTableName, System.Data.DataTable p_AdoOleDbDataTable, bool p_bDeleteExistingTable)
        {
            Microsoft.Office.Interop.Access.Dao.Field p_DaoField;



            string strMDBTableName = p_strMDBTableName;
            int x = 0;
            int intFieldSize = 0;
            this.m_intError = 0;
            this.m_strError = "";

            this.m_strTable = strMDBTableName;
            //open destination MDB file that will contain the created table
            this.OpenDb(strMDBPathAndFile);
            if (this.m_intError == 0)
            {

                //see if the table name already exists within the destination mdb 
                if (this.TableExists(this.m_DaoDatabase, strMDBTableName) == true)
                {
                    if (p_bDeleteExistingTable == false)
                    {
                        //the table exists so lets add a number onto the
                        //table name and see if it exists
                        for (x = 1; x <= 99; x++)
                        {
                            strMDBTableName = p_strMDBTableName.Trim() + x.ToString().Trim();
                            if (this.TableExists(this.m_DaoDatabase, strMDBTableName) == false)
                            {
                                this.m_strTable = strMDBTableName;
                                break;
                            }
                        }
                        //i should check if table exists here but i am
                        //going to assume there will never be a 
                        //table name from 1 to 99
                    }
                    else
                    {
                        m_DaoDatabase.TableDefs.Delete(strMDBTableName);
                    }
                }

                try
                {
                    //create a table and structure in the destination mdb
                    //first create the table definition
                    this.m_DaoTableDef = this.m_DaoDatabase.CreateTableDef(strMDBTableName, 0, "", "");

                    for (x = 0; x <= p_AdoOleDbDataTable.Rows.Count - 1; x++)
                    {
                        intFieldSize = 0;
                        //create the field based on the field name, datatype, and size
                        switch (p_AdoOleDbDataTable.Rows[x]["DataType"].ToString().Trim())
                        {
                            case "System.Single":
                                intFieldSize = 9;
                                p_DaoField =
                                    m_DaoTableDef.CreateField(p_AdoOleDbDataTable.Rows[x]["ColumnName"].ToString().Trim(),
                                    (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbSingle, intFieldSize);

                                break;
                            case "System.Double":
                                intFieldSize = 17;
                                p_DaoField =
                                    m_DaoTableDef.CreateField(p_AdoOleDbDataTable.Rows[x]["ColumnName"].ToString().Trim(),
                                    (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDouble, intFieldSize);
                                break;
                            case "System.Decimal":
                                intFieldSize = 17;
                                p_DaoField =
                                    m_DaoTableDef.CreateField(p_AdoOleDbDataTable.Rows[x]["ColumnName"].ToString().Trim(),
                                    (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDouble, intFieldSize);
                                break;
                            case "System.String":
                                intFieldSize = Convert.ToInt32(p_AdoOleDbDataTable.Rows[x]["ColumnSize"]);
                                if (intFieldSize <= 255)
                                {
                                    p_DaoField =
                                        m_DaoTableDef.CreateField(p_AdoOleDbDataTable.Rows[x]["ColumnName"].ToString().Trim(),
                                        (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText, intFieldSize);
                                }
                                else
                                {
                                    p_DaoField =
                                        m_DaoTableDef.CreateField(p_AdoOleDbDataTable.Rows[x]["ColumnName"].ToString().Trim(),
                                        (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbMemo, intFieldSize);

                                }

                                if (Convert.ToBoolean(p_AdoOleDbDataTable.Rows[x]["AllowDBNull"]) == true)
                                {
                                    p_DaoField.AllowZeroLength = true;
                                }
                                else
                                {
                                    p_DaoField.AllowZeroLength = false;
                                }



                                break;
                            case "System.Int16":
                                intFieldSize = 6;
                                p_DaoField =
                                    m_DaoTableDef.CreateField(p_AdoOleDbDataTable.Rows[x]["ColumnName"].ToString().Trim(),
                                    (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbInteger, intFieldSize);
                                break;
                            case "System.Char":
                                intFieldSize = Convert.ToInt32(p_AdoOleDbDataTable.Rows[x]["ColumnSize"]);
                                p_DaoField =
                                    m_DaoTableDef.CreateField(p_AdoOleDbDataTable.Rows[x]["ColumnName"].ToString().Trim(),
                                    (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText, intFieldSize);
                                if (Convert.ToBoolean(p_AdoOleDbDataTable.Rows[x]["AllowDBNull"]) == true)
                                {
                                    p_DaoField.AllowZeroLength = true;
                                }
                                else
                                {
                                    p_DaoField.AllowZeroLength = false;
                                }
                                break;
                            case "System.Int32":
                                intFieldSize = 11;
                                p_DaoField =
                                    m_DaoTableDef.CreateField(p_AdoOleDbDataTable.Rows[x]["ColumnName"].ToString().Trim(),
                                    (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbLong, intFieldSize);
                                break;
                            case "System.DateTime":
                                intFieldSize = 19;
                                p_DaoField =
                                    m_DaoTableDef.CreateField(p_AdoOleDbDataTable.Rows[x]["ColumnName"].ToString().Trim(),
                                    (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDate, intFieldSize);
                                break;

                            case "System.Int64":
                                intFieldSize = 17;
                                p_DaoField =
                                    m_DaoTableDef.CreateField(p_AdoOleDbDataTable.Rows[x]["ColumnName"].ToString().Trim(),
                                    (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbDouble, intFieldSize);
                                break;
                            case "System.Byte":
                                intFieldSize = 3;
                                p_DaoField =
                                    m_DaoTableDef.CreateField(p_AdoOleDbDataTable.Rows[x]["ColumnName"].ToString().Trim(),
                                    (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbByte, intFieldSize);
                                if (Convert.ToBoolean(p_AdoOleDbDataTable.Rows[x]["AllowDBNull"]) == true)
                                {

                                }
                                else
                                {

                                }
                                break;
                            case "System.Boolean":
                                intFieldSize = 3;
                                p_DaoField =
                                    m_DaoTableDef.CreateField(p_AdoOleDbDataTable.Rows[x]["ColumnName"].ToString().Trim(),
                                    (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbBoolean, intFieldSize);
                                break;
                            default:

                                MessageBox.Show(p_AdoOleDbDataTable.Rows[x]["DataType"].ToString() + " is undefined");
                                this.m_intError = -1;
                                this.m_DaoDatabase.Close();
                                this.m_DaoDatabase = null;
                                return "";

                        }
                        if (intFieldSize > 0)
                        {
                            this.m_DaoTableDef.Fields.Append(p_DaoField);
                        }

                    }

                    this.m_DaoDatabase.TableDefs.Append(this.m_DaoTableDef);



                }
                catch (Exception Caught)
                {
                    this.m_intError = -1;
                    this.m_strError = Caught.Message;
                    MessageBox.Show(m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    this.m_DaoTableDef = null;
                    return "";
                }


                this.m_DaoDatabase.Close();
                this.m_DaoDatabase = null;


            }
            return strMDBTableName;





        }
        /// <summary>
        /// create a primary index
        /// </summary>
        /// <param name="p_strMDBFile">full path and file name</param>
        /// <param name="p_strTable"></param>
        /// <param name="p_strFields">one or more field names separated by commas</param>
        public void CreatePrimaryKeyIndex(string p_strMDBFile, string p_strTable, string p_strFields)
        {
            this.m_intError = 0;
            this.m_strError = "";

            //open destination MDB file that will contain the created table
            this.OpenDb(p_strMDBFile);
            if (this.m_intError == 0)
            {

                //see if the table name already exists within the destination mdb 
                if (this.TableExists(this.m_DaoDatabase, p_strTable) == true)
                {
                    try
                    {
                        string cmd = "CREATE INDEX _PrimaryKey ON " + p_strTable + " (" + p_strFields + ") WITH PRIMARY";
                        this.m_DaoDatabase.Execute(cmd, null);
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message, "DAO Error",
                                         System.Windows.Forms.MessageBoxButtons.OK,
                                         System.Windows.Forms.MessageBoxIcon.Exclamation);
                        this.m_intError = -1;
                    }
                }
                this.m_DaoDatabase.Close();
                this.m_DaoDatabase = null;

            }

        }
        public void CreatePrimaryKeyIndex(Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabase, string p_strTable, string p_strFields)
        {
            this.m_intError = 0;
            this.m_strError = "";


            //see if the table name already exists within the destination mdb 
            if (this.TableExists(p_DaoDatabase, p_strTable) == true)
            {
                try
                {
                    string cmd = "CREATE INDEX _PrimaryKey ON " + p_strTable + " (" + p_strFields + ") WITH PRIMARY";
                    p_DaoDatabase.Execute(cmd, null);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "DAO Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Exclamation);
                    this.m_intError = -1;
                }
            }
        }

        /// <summary>
        ///  designate a field as autonumber datatype
        /// </summary>
        /// <param name="p_strMDBFile">full path and mdb file name</param>
        /// <param name="p_strTable">table name</param>
        /// <param name="p_strField">field name</param>
        public void CreateAutoNumber(string p_strMDBFile, string p_strTable, string p_strField)
        {
            this.m_intError = 0;
            this.m_strError = "";

            //open destination MDB file that will contain the created table
            this.OpenDb(p_strMDBFile);
            if (this.m_intError == 0)
            {

                //see if the table name already exists within the destination mdb 
                if (this.TableExists(this.m_DaoDatabase, p_strTable) == true)
                {
                    try
                    {
                        string cmd = "ALTER TABLE " + p_strTable + " ALTER COLUMN " + p_strField + " COUNTER(1,1)";
                        this.m_DaoDatabase.Execute(cmd, null);
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message, "DAO Error",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                        this.m_intError = -1;
                    }
                }
                this.m_DaoDatabase.Close();
                this.m_DaoDatabase = null;
            }

        }

        /// <summary>
        /// designate a field as autonumber datatype
        /// </summary>
        /// <param name="p_DaoDatabase">open mdb file variable</param>
        /// <param name="p_strTable">table name</param>
        /// <param name="p_strField">field name</param>
        public void CreateAutoNumber(Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabase, string p_strTable, string p_strField)
        {
            this.m_intError = 0;
            this.m_strError = "";

            //open destination MDB file that will contain the created table

            //see if the table name already exists within the destination mdb 
            if (this.TableExists(p_DaoDatabase, p_strTable) == true)
            {
                try
                {
                    string cmd = "ALTER TABLE " + p_strTable + " ALTER COLUMN " + p_strField + " COUNTER(1,1)";
                    p_DaoDatabase.Execute(cmd, null);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "DAO Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Exclamation);
                    this.m_intError = -1;
                }
            }



        }

        /// <summary>
        /// add the table names to a listbox
        /// </summary>
        /// <param name="strMDBPathAndFile"></param>
        /// <param name="listbox1"></param>
        public void LoadTablesIntoListBox(string strMDBPathAndFile, System.Windows.Forms.ListBox listbox1)
        {

            this.m_intError = 0;
            this.m_strError = "";
            this.OpenDb(strMDBPathAndFile);
            if (this.m_intError == 0)
            {
                foreach (Microsoft.Office.Interop.Access.Dao.TableDef table in this.m_DaoDatabase.TableDefs)
                {
                    if (table.Name.ToString().ToUpper().IndexOf("MSYS") != 0)
                    {
                        listbox1.Items.Add(table.Name.ToString());
                    }

                }
                this.m_DaoDatabase.Close();

            }
        }



        public void LoadTablesIntoListBox(Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabase, System.Windows.Forms.ListBox listbox1)
        {

            this.m_intError = 0;
            this.m_strError = "";
            if (this.m_intError == 0)
            {
                foreach (Microsoft.Office.Interop.Access.Dao.TableDef table in p_DaoDatabase.TableDefs)
                {
                    if (table.Name.ToString().ToUpper().IndexOf("MSYS") != 0)
                    {
                        listbox1.Items.Add(table.Name.ToString());
                    }

                }
            }
        }

        /// <summary>
        /// add the field names to a checked listbox
        /// </summary>
        /// <param name="strMDBPathAndFile"></param>
        /// <param name="strTableName"></param>
        /// <param name="CheckedListBox1"></param>
        public void LoadFieldsIntoCheckedListBox(string strMDBPathAndFile, string strTableName, System.Windows.Forms.CheckedListBox CheckedListBox1)
        {
            int x = 0;
            this.m_intError = 0;
            this.m_strError = "";
            //open destination MDB file that the table will be moved to
            this.OpenDb(strMDBPathAndFile);
            if (this.m_intError == 0)
            {
                try
                {
                    for (x = 0; x <= this.m_DaoDatabase.TableDefs[strTableName].Fields.Count - 1; x++)
                    {
                        CheckedListBox1.Items.Add(this.m_DaoDatabase.TableDefs[strTableName].Fields[x].Name);
                    }
                }
                catch (Exception caught)
                {
                    this.m_strError = "dao_data_access.getFieldNames() error: " + caught.Message;
                    this.m_intError = -1;
                    MessageBox.Show(this.m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    return;
                }
                this.m_DaoDatabase.Close();
                this.m_DaoDatabase = null;
            }

        }
        public void CompactMDB(string strMDBFile)
        {

            try
            {
                m_intError = 0;
                m_strError = "";
                string File_Path;
                string compact_file;
                utils p_utils = new utils();
                env p_env = new env();
                compact_file = p_utils.getRandomFile(p_env.strTempDir, "accdb");
                p_utils = null;
                p_env = null;

                File_Path = strMDBFile;

                if (System.IO.File.Exists(File_Path))
                {
                    this.m_DaoDbEngine.CompactDatabase(File_Path, compact_file, null, null, null);
                }
                if (System.IO.File.Exists(compact_file))
                {
                    System.IO.File.Delete(File_Path);
                    System.IO.File.Move(compact_file, File_Path);
                }
            }
            catch (Exception e)
            {
                if (DisplayErrors)
                {
                    MessageBox.Show(e.Message, "DAO Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Exclamation);
                }
                this.m_intError = -1;
            }
        }
        public void AddColumn_TextDataType(Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabase,
                              string p_strTableName,
                              string p_strColumnName,
                              int p_intLength,
                              string p_strDefaultValue)
        {
            m_intError = 0;
            m_strError = "";

            if (TableExists(p_DaoDatabase, p_strTableName) &&
                  ColumnExist(p_DaoDatabase, p_strTableName, p_strColumnName) == false)
            {
                m_DaoTableDef = p_DaoDatabase.TableDefs[p_strTableName];
                m_DaoField =
                   m_DaoTableDef.CreateField(p_strColumnName,
                   (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText, p_intLength);
                if (p_strDefaultValue.Trim().Length > 0)
                    m_DaoField.DefaultValue = p_strDefaultValue;
                m_DaoTableDef.Fields.Append(m_DaoField);
            }
            m_DaoField = null;
            m_DaoTableDef = null;
        }
        public bool RenameField(string p_strDatabase, string p_strTable, string p_strOldFieldName, string p_strNewFieldName)
        {

            this.m_intError = 0;
            this.m_strError = "";


            try
            {
                //If the table doesn't exist, we can't delete the column
                if (!TableExists(p_strDatabase, p_strTable))
                {
                    return false;
                }
                this.OpenDb(p_strDatabase);
                Microsoft.Office.Interop.Access.Dao.Field fld = this.m_DaoDatabase.TableDefs[p_strTable].Fields[p_strOldFieldName];
                if (fld != null)
                {
                    fld.Name = p_strNewFieldName;
                }
                else
                {
                    return false;
                }
                return true;
            }
            catch (Exception err)
            {
                m_intError = -1;
                m_strError = err.Message;
                MessageBox.Show(this.m_strError);
                return false;
            }
            finally
            {
                this.m_DaoDatabase.Close();
                this.m_DaoDatabase = null;
            }
        }


        public bool DisplayErrors
        {
            set { this._bDisplayErrors = value; }
            get { return this._bDisplayErrors; }
        }

        private void WriteText(string p_strTextFile, string p_strText)
        {
            System.IO.FileStream oTextFileStream;
            System.IO.StreamWriter oTextStreamWriter;

            if (!System.IO.File.Exists(p_strTextFile))
            {
                oTextFileStream = new System.IO.FileStream(p_strTextFile, System.IO.FileMode.Create,
                    System.IO.FileAccess.Write);
            }
            else
            {
                oTextFileStream = new System.IO.FileStream(p_strTextFile, System.IO.FileMode.Append,
                    System.IO.FileAccess.Write);
            }

            oTextStreamWriter = new System.IO.StreamWriter(oTextFileStream);
            oTextStreamWriter.Write(p_strText);
            oTextStreamWriter.Close();
            oTextFileStream.Close();
        }

        public void CopyQueriesToDatabase(string strMDBFileSource, string strMDBFileDestination)
        {
            Microsoft.Office.Interop.Access.Dao.Database p_DaoDatabaseSource;

            this.m_intError = 0;
            this.m_strError = "";


            //open destination DB file that the queries will be copied to
            this.OpenDb(strMDBFileDestination);
            if (this.m_intError == 0)
            {
                //open the source db file and table
                this.m_strError = "";
                this.m_intError = 0;
                try
                {
                    p_DaoDatabaseSource = this.m_DaoDbEngine.Workspaces[0].OpenDatabase(strMDBFileSource, false, false, string.Empty);
                }
                catch
                {
                    this.m_strError = "DAO Error Opening Database " + strMDBFileSource.ToString();
                    this.m_intError = -1;
                    MessageBox.Show(this.m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    return;
                }

                try
                {
                    // Retrieve queries from source database
                    QueryDefs destQdfs = this.m_DaoDatabase.QueryDefs;
                    foreach (QueryDef qdfLoop in p_DaoDatabaseSource.QueryDefs)
                    {
                        // Check to see if the query already exists, if it does we don't try to overwrite it
                        bool bQueryExists = false;
                        foreach (QueryDef testDef in destQdfs)
                        {
                            if (testDef.Name.Trim().ToLower().Equals(qdfLoop.Name.Trim().ToLower()))
                            {
                                bQueryExists = true;
                                break;
                            }

                        }
                        if (bQueryExists == false)
                        {
                            QueryDef qdfNew = this.m_DaoDatabase.CreateQueryDef(qdfLoop.Name, qdfLoop.SQL);
                        }
                    }
                }         
                catch (Exception Caught)
                {
                    this.m_intError = -1;
                    this.m_strError = Caught.Message;
                    MessageBox.Show(m_strError);
                    this.m_DaoDatabase.Close();
                    this.m_DaoDatabase = null;
                    this.m_DaoTableDef = null;
                    p_DaoDatabaseSource.Close();
                    p_DaoDatabaseSource = null;
                    return;
                }

                this.m_DaoTableDef = null;
                this.m_DaoDatabase.Close();
                this.m_DaoDatabase = null;
                p_DaoDatabaseSource.Close();
                p_DaoDatabaseSource = null;
            }

        }

        /// <summary>
        /// see if a index exists in an mdb file
        /// </summary>
        /// <param name="strMDBPathAndFile">full path and MDB file name</param>
        /// <param name="strTable">name of the table</param>
        /// <returns></returns>
        public bool IndexExists(string strMDBPathAndFile, string strTable, string strIndex)
        {
            bool p_bResult = false;
            //int z = 0;
            this.m_intError = 0;
            this.m_strError = "";
            this.OpenDb(strMDBPathAndFile);
            if (this.m_intError == 0)
            {
                foreach (Microsoft.Office.Interop.Access.Dao.TableDef table in this.m_DaoDatabase.TableDefs)
                {
                    if (table.Name.ToLower() == strTable.ToLower())
                    {
                        foreach(Microsoft.Office.Interop.Access.Dao.Index idx in table.Indexes)
                        {
                            if (idx.Name.ToLower() == strIndex.ToLower())
                            {
                                p_bResult = true;
                                break;
                            }
                        }
                    }
                }
            }

            return p_bResult;
        }

    }



}


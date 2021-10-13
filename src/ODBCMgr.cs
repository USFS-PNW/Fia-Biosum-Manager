using System;
using Microsoft.Win32;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    /// <summary>
    /// Summary description for Class1.
    /// </summary>
    public class ODBCMgr
    {
        private const string ODBC_INI_REG_PATH = "SOFTWARE\\ODBC\\ODBC.INI\\";
        Microsoft.Win32.RegistryKey m_oCurrentUserRegKey;
        Microsoft.Win32.RegistryKey m_oLocalMachineRegKey;
        public int m_intError = 0;
        public string m_strError = "";

        public ODBCMgr()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        /// <summary>
        /// Determine if the dsn name exists as a key in the current users SOFTWARE\ODBC\ODBC.INI\ registry path
        /// </summary>
        /// <param name="p_strDsn"></param>
        /// <returns></returns>
        public bool CurrentUserDSNKeyExist(string p_strDsn)
        {

            m_oCurrentUserRegKey = Registry.CurrentUser.OpenSubKey(ODBC_INI_REG_PATH, false);
            // get string name in string array
            // then use getValue to get data value.
            //save the odbc datasources in strDsnNames
            string[] strSubKeyNames = m_oCurrentUserRegKey.GetSubKeyNames(); // for the 1st time it's only 2 names
            m_oCurrentUserRegKey.Close();
            foreach (string strSubKeyName in strSubKeyNames)
            {
                if (strSubKeyName.Trim().ToUpper() == p_strDsn.Trim().ToUpper())
                {
                    return true;
                }
            }
            return false;
        }
        public bool LocalMachineDSNKeyExist(string p_strDsn)
        {

            m_oLocalMachineRegKey = Registry.LocalMachine.OpenSubKey(ODBC_INI_REG_PATH, false);
            string[] strSubKeyNames = m_oLocalMachineRegKey.GetSubKeyNames();
            m_oLocalMachineRegKey.Close();
            foreach (string strSubKeyName in strSubKeyNames)
            {
                if (strSubKeyName.Trim().ToUpper() == p_strDsn.Trim().ToUpper())
                {
                    return true;
                }
            }
            return false;
        }

        public void CreateUserSQLiteDSN(string p_strDSN, string p_strDBFileName)
        {
            m_intError = 0;
            m_strError = "";
            try
            {
                m_oCurrentUserRegKey = Registry.CurrentUser.OpenSubKey(ODBC_INI_REG_PATH + "\\ODBC Data Sources", true);
                m_oCurrentUserRegKey.SetValue(p_strDSN, "SQLite3 ODBC Driver");
                m_oCurrentUserRegKey.Close();

                m_oCurrentUserRegKey = Registry.CurrentUser.CreateSubKey(ODBC_INI_REG_PATH + "\\" + p_strDSN, true);
                m_oCurrentUserRegKey.SetValue("BigInt", "0");
                m_oCurrentUserRegKey.SetValue("Database", p_strDBFileName);
                m_oCurrentUserRegKey.SetValue("Description", "");
                m_oCurrentUserRegKey.SetValue("Driver", @"C:\windows\system32\sqlite3odbc.dll");
                m_oCurrentUserRegKey.SetValue("FKSupport", "0");
                m_oCurrentUserRegKey.SetValue("JDConv", "0");
                m_oCurrentUserRegKey.SetValue("LoadExt", "");
                m_oCurrentUserRegKey.SetValue("LongNames", "0");
                m_oCurrentUserRegKey.SetValue("NoCreat", "0");
                m_oCurrentUserRegKey.SetValue("NoTXN", "0");
                m_oCurrentUserRegKey.SetValue("NoWCHAR", "0");
                m_oCurrentUserRegKey.SetValue("OEMCP", "0");
                m_oCurrentUserRegKey.SetValue("PWD", "");
                m_oCurrentUserRegKey.SetValue("ShortNames", "0");
                m_oCurrentUserRegKey.SetValue("StepAPI", "0");
                m_oCurrentUserRegKey.SetValue("SyncPragma", "");
                m_oCurrentUserRegKey.SetValue("Timeout", "");

            }
            catch (Exception err)
            {
                m_intError = -1;
                m_strError = err.Message;
            }
            finally
            {
                m_oCurrentUserRegKey.Close();
            }
        }
        public void RemoveUserDSN(string p_strDSN)
        {
            m_intError = 0;
            m_strError = "";
            try
            {
                m_oCurrentUserRegKey = Registry.CurrentUser.OpenSubKey(ODBC_INI_REG_PATH + "\\ODBC Data Sources", true);
                m_oCurrentUserRegKey.DeleteValue(p_strDSN);
                Registry.CurrentUser.DeleteSubKeyTree(ODBC_INI_REG_PATH + "\\" + p_strDSN, false);
            }
            catch (Exception err)
            {
                m_intError = -1;
                m_strError = err.Message;
            }
            finally
            {
                m_oCurrentUserRegKey.Close();
            }
        }
    

		public void CurrentUserEditDsnUseridPassword(string p_strDsn,string p_strUserId, string p_strPW)
		{
			m_intError=0;
			m_strError="";
			try
			{
				m_oCurrentUserRegKey = Registry.CurrentUser.OpenSubKey(ODBC_INI_REG_PATH + "\\" + p_strDsn,true);
				m_oCurrentUserRegKey.SetValue("DSN",p_strDsn);
				m_oCurrentUserRegKey.SetValue("UserID",p_strUserId);
				m_oCurrentUserRegKey.SetValue("Password",p_strPW);
				m_oCurrentUserRegKey.Close();
			}
			catch (Exception e)
			{
				m_strError=e.Message;
				m_intError=-1;
				MessageBox.Show("ODBCMgr.CurrentUserEditDsnUser:" + m_strError,"ODBCManager",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);
			}
 
		}
        /// <summary>
        /// Copy all properties associated with a DataSourceName from local machine to local user
        /// </summary>
        /// <param name="p_strDsn"></param>
		public void CopyLocalMachineDsn(string p_strDsn)
		{

			//
			if (!this.LocalMachineDSNKeyExist(p_strDsn))
			{
				m_intError=-1;
				MessageBox.Show("Registry does not have a local machine ODBC.INI key for DSN name " + p_strDsn,"ODBCManager",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);
				return;
			}
			//delete the current user DSN subkey if it exists
			if (this.CurrentUserDSNKeyExist(p_strDsn))
			{
				m_oCurrentUserRegKey = Registry.CurrentUser.OpenSubKey(ODBC_INI_REG_PATH,true);
				m_oCurrentUserRegKey.DeleteSubKey(p_strDsn);
			}
			else
			{
				m_oCurrentUserRegKey = Registry.CurrentUser.OpenSubKey(ODBC_INI_REG_PATH,true);
			}
			m_oCurrentUserRegKey.CreateSubKey(p_strDsn);

			m_oLocalMachineRegKey = Registry.LocalMachine.OpenSubKey(ODBC_INI_REG_PATH + "\\" + p_strDsn ,false); 
			m_oCurrentUserRegKey = Registry.CurrentUser.OpenSubKey(ODBC_INI_REG_PATH+ "\\" + p_strDsn,true);
			string[] strValueNames = m_oLocalMachineRegKey.GetValueNames();
			foreach(string strValueName in strValueNames)
			{
				m_oCurrentUserRegKey.SetValue(strValueName,(string)m_oLocalMachineRegKey.GetValue(strValueName));
			}
			m_oLocalMachineRegKey.Close();
			m_oCurrentUserRegKey.Close();
			



			

		}

        public class DSN_KEYS
        {
            public DSN_KEYS()
            {
            }
            static public string Fia2FvsInputDsnName { get { return "FIABIOSUM_FIA2FVS_INPUT"; } }
            static public string Fia2FvsOutputDsnName { get { return "FIABIOSUM_FIA2FVS_OUTPUT"; } }
            static public string PlotInputDsnName { get { return "FIABIOSUM_PLOT_INPUT"; } }
            static public string OptimizerCalcVariableDsnName { get { return "FIABIOSUM_OPTIMIZER_CALC_VAR"; } }
            static public string ProcessorRuleDefinitionsDsnName { get { return "PROCESSOR_RULE_DEFINITIONS"; } }
            static public string ProcessorResultsDsnName { get { return "PROCESSOR_RESULTS"; } }
            static public string ProcessorTemporaryDsnName { get { return "PROCESSOR_TEMPORARY"; } }
        }

    }
}

using System;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    /// <summary>
	/// Summary description for ado_optimizer_tables.
	/// </summary>
    class sqlite_optimizer_tables
    {
		public System.Data.DataSet m_dsOptimizerTables;
		public int m_intNumberOfOptimizerTablesLoaded;
		public string[] m_strOptimizerTables;
		public System.Data.SQLite.SQLiteDataAdapter m_SQLiteDataAdapter;
		public System.Data.SQLite.SQLiteDataAdapter m_daPlot;
		public System.Data.SQLite.SQLiteDataAdapter m_daCond;
		public System.Data.SQLite.SQLiteDataAdapter m_daFFE;
		public System.Data.SQLite.SQLiteDataAdapter m_daHarvestCosts;
		public System.Data.SQLite.SQLiteDataAdapter m_daOwnerGroups;
		public System.Data.SQLite.SQLiteDataAdapter m_daTreeClass;
		public System.Data.SQLite.SQLiteDataAdapter m_daRX;
		public System.Data.SQLite.SQLiteDataAdapter m_daTreeSpecies;
		public System.Data.SQLite.SQLiteDataAdapter m_daTreeVol;
		public System.Data.SQLite.SQLiteDataAdapter m_daTravelTimes;
		public System.Data.SQLite.SQLiteDataAdapter m_daPSites;

		public System.Data.SQLite.SQLiteCommand m_SQLiteCommand;

		public System.Data.SQLite.SQLiteConnection m_SQLiteConnection;
		public System.Data.OleDb.OleDbConnection m_connPlot;
		public System.Data.OleDb.OleDbConnection m_connCond;
		public System.Data.SQLite.SQLiteConnection m_connMasterLink;
		public System.Data.SQLite.SQLiteConnection m_SQLiteConnectionFFE;
		public System.Data.SQLite.SQLiteConnection m_SQLiteConnectionTravelTimes;
		public System.Data.SQLite.SQLiteConnection m_SQLiteConnectionProcessingSites;
		public System.Data.SQLite.SQLiteConnection m_SQLiteConnectionScenario;

		public System.Data.SQLite.SQLiteDataReader m_sQLiteDataReader;
		public System.Data.DataTable m_DataTable;
		public string m_strError;
		public int m_intError;
		public const int NUMBER_OF_OPTIMIZER_TABLES = 11;
		public string m_strRandomFileName = "";
		public int m_intNumberOfOptimizerTables;

		public sqlite_optimizer_tables()
        {
			this.m_strOptimizerTables = new string[50];
			for (int x = 0; x <= NUMBER_OF_OPTIMIZER_TABLES - 1; x++)
			{
				this.m_strOptimizerTables[x] = "";
			}
			this.m_intNumberOfOptimizerTablesLoaded = 0;

			this.m_SQLiteDataAdapter = new System.Data.SQLite.SQLiteDataAdapter();
			this.m_daPlot = new System.Data.SQLite.SQLiteDataAdapter();
			this.m_daCond = new System.Data.SQLite.SQLiteDataAdapter();
			this.m_daFFE = new System.Data.SQLite.SQLiteDataAdapter();
			this.m_daHarvestCosts = new System.Data.SQLite.SQLiteDataAdapter();
			this.m_daOwnerGroups = new System.Data.SQLite.SQLiteDataAdapter();
			this.m_daTreeClass = new System.Data.SQLite.SQLiteDataAdapter();
			this.m_daRX = new System.Data.SQLite.SQLiteDataAdapter();
			this.m_daTreeSpecies = new System.Data.SQLite.SQLiteDataAdapter();
			this.m_daTreeVol = new System.Data.SQLite.SQLiteDataAdapter();
			this.m_daTravelTimes = new System.Data.SQLite.SQLiteDataAdapter();
			this.m_daPSites = new System.Data.SQLite.SQLiteDataAdapter();

			this.m_connPlot = new System.Data.OleDb.OleDbConnection();
			this.m_connCond = new System.Data.OleDb.OleDbConnection();


			this.m_dsOptimizerTables = new DataSet();
			this.m_dsOptimizerTables.Clear();
		}
	}
}

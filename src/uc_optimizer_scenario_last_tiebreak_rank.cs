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
	/// Summary description for uc_optimizer_scenario_last_tiebreak_rank.
	/// </summary>
	public class uc_optimizer_scenario_last_tiebreak_rank : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.DataGrid dataGrid1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.ImageList imgSize;
		public string strImagePath = "";
		public string m_strProjectDirectory;
		public int m_intFullHt = 280;
		public System.Data.OleDb.OleDbDataAdapter m_OleDbDataAdapter;
		public System.Data.SQLite.SQLiteDataAdapter m_SQLiteDataAdapter;
		public System.Data.DataSet m_DataSet;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionMaster;
		public System.Data.OleDb.OleDbConnection m_OleDbRxConn;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionScenario;
		public System.Data.OleDb.OleDbCommand m_OleDbCommand;
		public System.Data.SQLite.SQLiteCommand m_SQLiteCommand;
		public System.Data.DataRelation m_DataRelation;
		public System.Data.DataTable m_DataTable;
		public System.Data.DataRow m_DataRow;
		public string strRxPackageTableName;
		public string strRxConn;
		public string strRxMDBFile;
		public string strScenarioId;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario = null;


		public uc_optimizer_scenario_last_tiebreak_rank()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			// TODO: Add any initialization after the InitializeComponent call

		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(uc_optimizer_scenario_last_tiebreak_rank));
			this.dataGrid1 = new System.Windows.Forms.DataGrid();
			this.imgSize = new System.Windows.Forms.ImageList(this.components);
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// dataGrid1
			// 
			this.dataGrid1.DataMember = "";
			this.dataGrid1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.dataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dataGrid1.Location = new System.Drawing.Point(3, 16);
			this.dataGrid1.Name = "dataGrid1";
			this.dataGrid1.Size = new System.Drawing.Size(738, 261);
			this.dataGrid1.TabIndex = 13;
			this.dataGrid1.Click += new System.EventHandler(this.dataGrid1_Click);
			this.dataGrid1.CurrentCellChanged += new System.EventHandler(this.dataGrid1_CurrentCellChanged);
			// 
			// imgSize
			// 
			this.imgSize.ImageSize = new System.Drawing.Size(16, 16);
			this.imgSize.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgSize.ImageStream")));
			this.imgSize.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// groupBox1
			// 
			this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox1.Controls.Add(this.dataGrid1);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(744, 280);
			this.groupBox1.TabIndex = 16;
			this.groupBox1.TabStop = false;
			// 
			// uc_scenario_treatment_intensity
			// 
			this.BackColor = System.Drawing.SystemColors.Control;
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_scenario_treatment_intensity";
			this.Size = new System.Drawing.Size(744, 280);
			this.Resize += new System.EventHandler(this.uc_scenario_last_tiebreak_rank_Resize);
			((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void uc_scenario_last_tiebreak_rank_Resize(object sender, System.EventArgs e)
		{
			try
			{

				this.dataGrid1.Width = this.ClientSize.Width - this.dataGrid1.Left * 2;

				this.dataGrid1.Height = this.ClientSize.Height - this.dataGrid1.Top;

			}
			catch
			{
			}


		}

		private void btnSize_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{


		}


		public int savevalues()
		{
			int x = 0;
			string strSQL = "";
			ado_data_access p_ado = new ado_data_access();
			p_ado.m_intError = 0;
			try
			{
				for (x = 0; x <= this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows.Count - 1; x++)
				{
					if (m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows[x]["last_tiebreak_rank"] != System.DBNull.Value &&
						m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows[x]["last_tiebreak_rank"].ToString().Trim().Length > 0)
					{
						strSQL = "UPDATE scenario_last_tiebreak_rank SET last_tiebreak_rank = " +
							this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows[x]["last_tiebreak_rank"] +
							" WHERE TRIM(rxpackage) = '" + this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows[x]["rxpackage"].ToString().Trim() + "';";
					}
					else
					{
						strSQL = "UPDATE scenario_last_tiebreak_rank SET last_tiebreak_rank = null " +
								 " WHERE TRIM(rxpackage) = '" + this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows[x]["rxpackage"].ToString().Trim() + "';";
					}
					p_ado.SqlNonQuery(this.m_OleDbConnectionScenario, strSQL);
					if (p_ado.m_intError < 0) break;

				}
			}
			catch (Exception caught)
			{
				MessageBox.Show("Function: uc_optimizer_scenario_last_tiebreak_rank.savevalues ErrMsg:" + caught.Message + " Failed updating scenario_last_tiebreak_rank table with last tiebreak rank ratings");
			}
			x = p_ado.m_intError;
			p_ado = null;
			return x;


		}

		public void loadgrid(bool p_bScenarioCopy)
		{
			string[] strDeleteSQL = new string[25];
			string strSQL = "";
			int intArrayCount;
			int x = 0;
			string strConn = "";

			string strScenarioMDB = "";

			ado_data_access p_ado = new ado_data_access();

			this.strScenarioId = ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();

			/***************************************************
			 **scenario mdb connection
			 ***************************************************/
			if (p_ado.m_intError != 0) return;
			p_ado.getScenarioConnStringAndMDBFile(ref strScenarioMDB,
										  ref strConn,
										 frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim());

			this.m_OleDbConnectionScenario = new System.Data.OleDb.OleDbConnection();
			p_ado.OpenConnection(strConn, ref this.m_OleDbConnectionScenario);

			/*************************************************************************
			 **get the treatment prescription mdb file,table, and connection strings
			 *************************************************************************/
			p_ado.getScenarioDataSourceConnStringAndTable(ref this.strRxMDBFile,
											ref this.strRxPackageTableName, ref this.strRxConn,
											"Treatment Packages",
											this.strScenarioId,
											this.m_OleDbConnectionScenario);


			this.m_OleDbRxConn = new System.Data.OleDb.OleDbConnection();
			p_ado.OpenConnection(this.strRxConn, ref this.m_OleDbRxConn);
			if (p_ado.m_intError != 0)
			{
				this.m_OleDbConnectionScenario.Close();
				this.m_OleDbConnectionScenario = null;
				this.m_OleDbRxConn = null;
				return;
			}
			strSQL = "select * from " + this.strRxPackageTableName;
			p_ado.SqlQueryReader(this.m_OleDbRxConn, strSQL);

			/********************************************************************************
             **insert records into the scenario_last_tiebreak_rank table from the master rxpackage table
             ********************************************************************************/
			if (p_ado.m_intError == 0)
			{
				this.m_DataSet = new System.Data.DataSet();
				this.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
				while (p_ado.m_OleDbDataReader.Read())
				{
					strSQL = "select * from scenario_last_tiebreak_rank " +
						" where scenario_id = '" + this.strScenarioId + "' and " +
						"rxpackage = '" + p_ado.m_OleDbDataReader["rxpackage"].ToString() + "';";
					this.m_OleDbCommand = this.m_OleDbConnectionScenario.CreateCommand();
					this.m_OleDbCommand.CommandText = strSQL;
					this.m_OleDbDataAdapter.SelectCommand = this.m_OleDbCommand;
					this.m_OleDbDataAdapter.Fill(this.m_DataSet, "scenario_last_tiebreak_rank");

					/*******************************************************************************
					 **if the master treatment record is not found in the scenario db than insert it
					 *******************************************************************************/
					if (this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows.Count == 0)
					{
						strSQL = "INSERT INTO scenario_last_tiebreak_rank (scenario_id," +
							"rxpackage, last_tiebreak_rank) VALUES " +
							"('" + this.strScenarioId + "'," +
							"'" + p_ado.m_OleDbDataReader["rxpackage"].ToString() + "'," +
							"0);";

						p_ado.SqlNonQuery(this.m_OleDbConnectionScenario, strSQL);
					}
					this.m_DataSet.Tables.Clear();

				}
				p_ado.m_OleDbDataReader.Close();
				this.m_DataSet.Dispose();
				this.m_OleDbDataAdapter.Dispose();


				intArrayCount = 0;
				/****************************************************************************************************
				 **delete any prescriptions from the scenario db that do not exist in the master
				 ****************************************************************************************************/
				strSQL = "select * from scenario_last_tiebreak_rank where scenario_id = '" + this.strScenarioId + "';";
				p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);


				if (p_ado.m_intError == 0)
				{
					this.m_DataSet = new System.Data.DataSet();
					this.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
					while (p_ado.m_OleDbDataReader.Read())
					{
						/************************************************************************
						 **query the scenario treatment in the master db. If it is not found
						 **in the master db then delete the scenario treatment record
						 ************************************************************************/
						strSQL = "select * from " + this.strRxPackageTableName +
							" where rxpackage = '" + p_ado.m_OleDbDataReader["rxpackage"].ToString() + "';";
						this.m_OleDbCommand = this.m_OleDbRxConn.CreateCommand();
						this.m_OleDbCommand.CommandText = strSQL;
						this.m_OleDbDataAdapter.SelectCommand = this.m_OleDbCommand;
						this.m_OleDbDataAdapter.Fill(this.m_DataSet, this.strRxPackageTableName);
						if (this.m_DataSet.Tables[this.strRxPackageTableName].Rows.Count == 0)
						{
							strDeleteSQL[intArrayCount] = "DELETE FROM scenario_last_tiebreak_rank" +
								" WHERE scenario_id = '" + this.strScenarioId + "'" +
								" AND rxpackage = '" + p_ado.m_OleDbDataReader["rxpackage"] + "';";
							intArrayCount++;
						}
						this.m_DataSet.Tables.Clear();
					}

					p_ado.m_OleDbDataReader.Close();
					this.m_DataSet.Dispose();
					this.m_OleDbDataAdapter.Dispose();
					/**********************************************************************************
					 **if there were any treatments that were loaded into sql delete 
					 **arrays then perform the sql to delete the treatments out of the table
					 **********************************************************************************/
					if (intArrayCount > 0)
					{
						for (x = 0; x <= intArrayCount - 1; x++)
						{
							p_ado.SqlNonQuery(this.m_OleDbConnectionScenario, strDeleteSQL[x].ToString());
						}
					}
				}

				/***************************************************************************************
				 **okay, now that the table has been validated and updated lets load the grid to
				 **display the treatments to the user
				 ***************************************************************************************/
				this.m_DataSet = new System.Data.DataSet();
				this.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
				this.m_OleDbCommand = this.m_OleDbRxConn.CreateCommand();
				this.m_OleDbCommand.CommandText = "select * from " + this.strRxPackageTableName;
				this.m_OleDbDataAdapter.SelectCommand = this.m_OleDbCommand;
				this.m_OleDbDataAdapter.Fill(this.m_DataSet, this.strRxPackageTableName);

				this.m_OleDbCommand = this.m_OleDbConnectionScenario.CreateCommand();
				strSQL = "";
				for (x = 0; x <= this.m_DataSet.Tables[this.strRxPackageTableName].Rows.Count - 1; x++)
				{
					if (this.m_DataSet.Tables[this.strRxPackageTableName].Rows[x]["rxpackage"].ToString().Length > 0)
					{
						strSQL = "select scenario_id,rxpackage,last_tiebreak_rank from scenario_last_tiebreak_rank where scenario_id = '" + this.strScenarioId + "';";
						break;
					}
				}
				/************************************************************
				 **if no records in the master prescription table then return
				 ************************************************************/
				if (strSQL.Length == 0)
				{
					this.m_DataSet.Clear();
					this.m_DataSet.Dispose();
					this.m_OleDbDataAdapter.Dispose();
					this.m_OleDbCommand.Dispose();
					this.m_OleDbConnectionScenario.Close();
					this.m_OleDbConnectionScenario = null;
					this.m_OleDbRxConn.Close();
					this.m_OleDbRxConn = null;
					return;
				}

				/*******************************
				 **create the data sets
				 *******************************/
				this.m_OleDbCommand = this.m_OleDbConnectionScenario.CreateCommand();
				this.m_OleDbCommand.CommandText = "select scenario_id,rxpackage,last_tiebreak_rank from scenario_last_tiebreak_rank where scenario_id = '" + this.strScenarioId + "';";
				this.m_OleDbDataAdapter.SelectCommand = this.m_OleDbCommand;
				this.m_OleDbDataAdapter.Fill(this.m_DataSet, "scenario_last_tiebreak_rank");

				/*****************************************************************************
				 **add the description column to the scenario last tiebreak rank dataset
				 *****************************************************************************/
				this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Columns.Add("Description");

				/********************************************************************************
				 **add the treatment description value to the scenrario rx last tiebreak rank data set.
				 **the description is only in the master rx table and is added to the
				 **scenario last tiebreak rank data set for information purposes.
				 ********************************************************************************/

				/***********************************************************************************
                 **for loop through the master db rx dataset adding the description field to the
                 **scenenario db scenario_last_tiebreak_rank dataset
                 ***********************************************************************************/
				for (x = 0; x <= this.m_DataSet.Tables[this.strRxPackageTableName].Rows.Count - 1; x++)
				{
					if (this.m_DataSet.Tables[this.strRxPackageTableName].Rows[x]["rxpackage"].ToString().Length > 0)
					{
						/***************************************************************************************
						 **build the expression to filter only the scenario row that meets the expression
						 ***************************************************************************************/
						strSQL = "rxpackage = '" + this.m_DataSet.Tables[this.strRxPackageTableName].Rows[x]["rxpackage"] + "'";

						/***************************************************************************************
						 **create a datarow that will hold the results from the query expression
						 ***************************************************************************************/
						System.Data.DataRow[] dr1;
						dr1 = this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Select(strSQL);

						/***************************************************************************************
						 **check to see if it found the master rx treatment in the sceanrio last tiebreak rank dataset
						 ***************************************************************************************/
						if (dr1.Length != 0)
						{
							/***************************************************************************************
							 **it found it, loop through the dataset and find the row that matches the row 
							 **returned from the search expression
							 ***************************************************************************************/
							for (int y = 0; y <= this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows.Count - 1; y++)
							{
								if (dr1[0]["rxpackage"] ==
									this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows[y]["rxpackage"])
								{
									/**********************************************************************************
									 **update the description row/column with the master db rx table description value
									 **********************************************************************************/
									this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows[y]["description"] =
										this.m_DataSet.Tables[this.strRxPackageTableName].Rows[x]["description"];
									break;
								}
							}
							this.m_DataSet.Tables["scenario_last_tiebreak_rank"].AcceptChanges();

						}

					}
				}


				/**************************************************************************************************
				 **place the dataset table into a view class so as to not allow new records to be appended
				 **************************************************************************************************/
				DataView firstView = new DataView(this.m_DataSet.Tables["scenario_last_tiebreak_rank"]);
				firstView.AllowNew = false;       //cannot append new records
				firstView.AllowDelete = false;    //cannot delete records

				/***************************************************************
				 **custom define the grid style
				 ***************************************************************/
				DataGridTableStyle tableStyle = new DataGridTableStyle();

				/***********************************************************************
				 **map the data grid table style to the last tiebreak rank dataset
				 ***********************************************************************/
				tableStyle.MappingName = "scenario_last_tiebreak_rank";
				tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
				tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
				tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
				tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;

				/******************************************************************************
				 **since the dataset has things like field name and number of columns,
				 **we will use those to create new columnstyles for the columns in our grid
				 ******************************************************************************/
				//get the number of columns from the scenario_last_tiebreak_rank data set
				int numCols = this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Columns.Count;

				/***********************************************************************************
				 **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
				 ***********************************************************************************/
				DataGridColoredTextBoxColumn aColumnTextColumn;

				//loop through all the columns in the dataset	
				for (int i = 0; i < numCols; ++i)
				{
					//create a new instance of the DataGridColoredTextBoxColumn class
					aColumnTextColumn = new DataGridColoredTextBoxColumn();
					aColumnTextColumn.HeaderText = this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Columns[i].ColumnName;
					//all columns are read-only except the last_tiebreak_rank column
					if (aColumnTextColumn.HeaderText != "last_tiebreak_rank") aColumnTextColumn.ReadOnly = true;
					//assign the mappingname property the data sets column name
					aColumnTextColumn.MappingName = this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Columns[i].ColumnName;
					//add the datagridcoloredtextboxcolumn object to the data grid table style object
					tableStyle.GridColumnStyles.Add(aColumnTextColumn);
					//set wider width for some columns
					switch (aColumnTextColumn.HeaderText)
					{
						case "scenario_id":
							aColumnTextColumn.Width = 150;
							break;
						case "Description":
							aColumnTextColumn.Width = 475;
							break;
					}
				}
				dataGrid1.BackgroundColor = frmMain.g_oGridViewBackgroundColor;
				dataGrid1.BackColor = frmMain.g_oGridViewRowBackgroundColor;
				if (frmMain.g_oGridViewFont != null) dataGrid1.Font = frmMain.g_oGridViewFont;

				// make the dataGrid use our new tablestyle and bind it to our table
				this.dataGrid1.TableStyles.Clear();
				this.dataGrid1.TableStyles.Add(tableStyle);

				// If this is a copied scenario, we will have a reference form to get the values
				if (p_bScenarioCopy == true)
				{
					if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oLastTieBreakRankItem_Collection != null)
					{
						OptimizerScenarioItem.LastTieBreakRankItem_Collection oLastTieBreakRankItem_Collection = ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oLastTieBreakRankItem_Collection;
						for (int i = 0; i < firstView.Count - 1; ++i)
						{
							for (x = 0; x <= oLastTieBreakRankItem_Collection.Count - 1; x++)
							{
								if (oLastTieBreakRankItem_Collection.Item(x).RxPackage.Equals(firstView[i]["rxpackage"]))
								{
									firstView[i]["last_tiebreak_rank"] = oLastTieBreakRankItem_Collection.Item(x).LastTieBreakRank;
								}
							}
						}
					}
				}
				this.dataGrid1.DataSource = firstView;
				this.dataGrid1.Expand(-1);

			}
		}
			
		public void loadgridsqlite(bool p_bScenarioCopy)
        {
			string[] strDeleteSQL = new string[25];
			int intArrayCount;
			int x = 0;

			DataMgr oDataMgr = new DataMgr();

			this.strScenarioId = ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();

			/***************************************************
			 **scenario db connection
			 ***************************************************/
			if (oDataMgr.m_intError != 0) return;
			string strScenarioDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "//" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
			string strConn = oDataMgr.GetConnectionString(strScenarioDB);
			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
				conn.Open();
				oDataMgr.m_strSQL = "SELECT path, file, table_name from scenario_datasource " +
						  "where TRIM(UPPER(table_type)) = 'TREATMENT PACKAGES' " +
						  "and TRIM(UPPER(scenario_id)) = '" + strScenarioId.Trim().ToUpper() + "';";

				oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
				if (oDataMgr.m_intError == 0)
                {
					while (oDataMgr.m_DataReader.Read())
                    {
						this.strRxPackageTableName = oDataMgr.m_DataReader["table_name"].ToString();
						string strRxDBFile = oDataMgr.m_DataReader["path"].ToString().Trim() + "\\" + oDataMgr.m_DataReader["file"].ToString().Trim();
						this.strRxConn = "Provider=Microsoft.Ace.OLEDB.12.0;" + oDataMgr.GetConnectionString(strRxDBFile) + ";User Id=admin;Password=;";
						break;
					}
					oDataMgr.m_DataReader.Close();
                }
				if (oDataMgr.m_intError != 0)
				{
					conn.Close();
					return;
				}

				ado_data_access p_ado = new ado_data_access();
				this.m_OleDbRxConn = new System.Data.OleDb.OleDbConnection();
				p_ado.OpenConnection(this.strRxConn, ref this.m_OleDbRxConn);
				if (p_ado.m_intError != 0)
				{
					this.m_OleDbConnectionScenario.Close();
					this.m_OleDbConnectionScenario = null;
					this.m_OleDbRxConn = null;
					return;
				}
				p_ado.m_strSQL = "select * from " + this.strRxPackageTableName;
				p_ado.SqlQueryReader(this.m_OleDbRxConn, p_ado.m_strSQL);

				/********************************************************************************
				**insert records into the scenario_last_tiebreak_rank table from the master rxpackage table
					********************************************************************************/
				if (p_ado.m_intError == 0)
				{
					this.m_DataSet = new System.Data.DataSet();
					this.m_SQLiteDataAdapter = new System.Data.SQLite.SQLiteDataAdapter();
					while (p_ado.m_OleDbDataReader.Read())
					{
						oDataMgr.m_strSQL = "SELECT * FROM scenario_last_tiebreak_rank " +
						" where scenario_id = '" + this.strScenarioId + "' and " +
						"rxpackage = '" + p_ado.m_OleDbDataReader["rxpackage"].ToString() + "';";

						this.m_SQLiteCommand = conn.CreateCommand();
						this.m_SQLiteCommand.CommandText = oDataMgr.m_strSQL;
						this.m_SQLiteDataAdapter.SelectCommand = this.m_SQLiteCommand;
						this.m_SQLiteDataAdapter.Fill(this.m_DataSet, "scenario_last_tiebreak_rank");

						/*******************************************************************************
						**if the master treatment record is not found in the scenario db than insert it
						*******************************************************************************/
						if (this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows.Count == 0)
						{
							oDataMgr.m_strSQL = "INSERT INTO scenario_last_tiebreak_rank (scenario_id," +
								"rxpackage, last_tiebreak_rank) VALUES " +
								"('" + this.strScenarioId + "'," +
								"'" + p_ado.m_OleDbDataReader["rxpackage"].ToString() + "'," +
								"0);";

							oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
						}
						this.m_DataSet.Tables.Clear();
					}
					p_ado.m_OleDbDataReader.Close();
					this.m_DataSet.Dispose();
					this.m_SQLiteDataAdapter.Dispose();

					intArrayCount = 0;
					/****************************************************************************************************
						**delete any prescriptions from the scenario db that do not exist in the master
						****************************************************************************************************/
					oDataMgr.m_strSQL = "SELECT * FROM scenario_last_tiebreak_rank where scenario_id = '" + this.strScenarioId + "';";
					oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);

					if (oDataMgr.m_intError == 0)
					{
						this.m_DataSet = new System.Data.DataSet();
						this.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
						while (oDataMgr.m_DataReader.Read())
						{
							/************************************************************************
								**query the scenario treatment in the master db. If it is not found
								**in the master db then delete the scenario treatment record
								************************************************************************/
							p_ado.m_strSQL = "SELECT * FROM " + this.strRxPackageTableName +
								" where rxpackage = '" + oDataMgr.m_DataReader["rxpackage"].ToString() + "';";
							this.m_OleDbCommand = this.m_OleDbRxConn.CreateCommand();
							this.m_OleDbCommand.CommandText = p_ado.m_strSQL;
							this.m_OleDbDataAdapter.SelectCommand = this.m_OleDbCommand;
							this.m_OleDbDataAdapter.Fill(this.m_DataSet, this.strRxPackageTableName);
							if (this.m_DataSet.Tables[this.strRxPackageTableName].Rows.Count == 0)
							{
								strDeleteSQL[intArrayCount] = "DELETE FROM scenario_last_tiebreak_rank" +
									" WHERE scenario_id = '" + this.strScenarioId + "'" +
									" AND rxpackage = '" + oDataMgr.m_DataReader["rxpackage"] + "';";
								intArrayCount++;
							}
							this.m_DataSet.Tables.Clear();
						}
						oDataMgr.m_DataReader.Close();
						this.m_DataSet.Dispose();
						this.m_OleDbDataAdapter.Dispose();
						/**********************************************************************************
							**if there were any treatments that were loaded into sql delete 
							**arrays then perform the sql to delete the treatments out of the table
							**********************************************************************************/
						if (intArrayCount > 0)
						{
							for (x = 0; x <= intArrayCount - 1; x++)
							{
								oDataMgr.SqlNonQuery(conn, strDeleteSQL[x].ToString());
							}
						}
					}
					/***************************************************************************************
					**okay, now that the table has been validated and updated lets load the grid to
					**display the treatments to the user
					***************************************************************************************/
					this.m_DataSet = new System.Data.DataSet();
					this.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
					this.m_OleDbCommand = this.m_OleDbRxConn.CreateCommand();
					this.m_OleDbCommand.CommandText = "SELECT * FROM " + this.strRxPackageTableName;
					this.m_OleDbDataAdapter.SelectCommand = this.m_OleDbCommand;
					this.m_OleDbDataAdapter.Fill(this.m_DataSet, this.strRxPackageTableName);

					this.m_SQLiteCommand = conn.CreateCommand();
					for (x = 0; x <= this.m_DataSet.Tables[this.strRxPackageTableName].Rows.Count - 1; x++)
					{
						if (this.m_DataSet.Tables[this.strRxPackageTableName].Rows[x]["rxpackage"].ToString().Length > 0)
						{
							oDataMgr.m_strSQL = "SELECT scenario_id, rxpackage, last_tiebreak_rank from scenario_last_tiebreak_rank where scenario_id = '" + this.strScenarioId + "';";
							break;
						}
					}
					/************************************************************
						**if no records in the master prescription table then return
						************************************************************/
					if (oDataMgr.m_strSQL.Length == 0)
					{
						this.m_DataSet.Clear();
						this.m_DataSet.Dispose();
						this.m_OleDbDataAdapter.Dispose();
						this.m_OleDbCommand.Dispose();
						conn.Close();
						this.m_OleDbRxConn.Close();
						return;
					}

					/*******************************
						**create the data sets
						*******************************/
					this.m_SQLiteDataAdapter = new System.Data.SQLite.SQLiteDataAdapter();
					this.m_SQLiteCommand = conn.CreateCommand();
					this.m_SQLiteCommand.CommandText = "SELECT scenario_id, rxpackage, last_tiebreak_rank from scenario_last_tiebreak_rank WHERE scenario_id = '" + this.strScenarioId + "';";
					this.m_SQLiteDataAdapter.SelectCommand = this.m_SQLiteCommand;
					this.m_SQLiteDataAdapter.Fill(this.m_DataSet, "scenario_last_tiebreak_rank");

					/*****************************************************************************
						**add the description column to the scenario last tiebreak rank dataset
						*****************************************************************************/
					this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Columns.Add("Description");

					/********************************************************************************
						**add the treatment description value to the scenrario rx last tiebreak rank data set.
						**the description is only in the master rx table and is added to the
						**scenario last tiebreak rank data set for information purposes.
						********************************************************************************/

					/***********************************************************************************
						**for loop through the master db rx dataset adding the description field to the
						**scenenario db scenario_last_tiebreak_rank dataset
						***********************************************************************************/
					for (x = 0; x <= this.m_DataSet.Tables[this.strRxPackageTableName].Rows.Count - 1; x++)
					{
						if (this.m_DataSet.Tables[this.strRxPackageTableName].Rows[x]["rxpackage"].ToString().Length > 0)
						{
							/***************************************************************************************
								**build the expression to filter only the scenario row that meets the expression
								***************************************************************************************/
							oDataMgr.m_strSQL = "rxpackage = '" + this.m_DataSet.Tables[this.strRxPackageTableName].Rows[x]["rxpackage"] + "'";

							/***************************************************************************************
								**create a datarow that will hold the results from the query expression
								***************************************************************************************/
							System.Data.DataRow[] dr1;
							dr1 = this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Select(oDataMgr.m_strSQL);

							/***************************************************************************************
								**check to see if it found the master rx treatment in the sceanrio last tiebreak rank dataset
								***************************************************************************************/
							if (dr1.Length != 0)
							{
								/***************************************************************************************
									**it found it, loop through the dataset and find the row that matches the row 
									**returned from the search expression
									***************************************************************************************/
								for (int y = 0; y <= this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows.Count - 1; y++)
								{
									if (dr1[0]["rxpackage"] ==
										this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows[y]["rxpackage"])
									{
										/**********************************************************************************
											**update the description row/column with the master db rx table description value
											**********************************************************************************/
										this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows[y]["description"] =
											this.m_DataSet.Tables[this.strRxPackageTableName].Rows[x]["description"];
										break;
									}
								}
								this.m_DataSet.Tables["scenario_last_tiebreak_rank"].AcceptChanges();
							}
						}
					}
					/**************************************************************************************************
						**place the dataset table into a view class so as to not allow new records to be appended
						**************************************************************************************************/
					DataView firstView = new DataView(this.m_DataSet.Tables["scenario_last_tiebreak_rank"]);
					firstView.AllowNew = false;       //cannot append new records
					firstView.AllowDelete = false;    //cannot delete records

					/***************************************************************
						**custom define the grid style
						***************************************************************/
					DataGridTableStyle tableStyle = new DataGridTableStyle();

					/***********************************************************************
						**map the data grid table style to the last tiebreak rank dataset
						***********************************************************************/
					tableStyle.MappingName = "scenario_last_tiebreak_rank";
					tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
					tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
					tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
					tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;

					/******************************************************************************
						**since the dataset has things like field name and number of columns,
						**we will use those to create new columnstyles for the columns in our grid
						******************************************************************************/
					//get the number of columns from the scenario_last_tiebreak_rank data set
					int numCols = this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Columns.Count;

					/***********************************************************************************
						**assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
						***********************************************************************************/
					DataGridColoredTextBoxColumn aColumnTextColumn;

					//loop through all the columns in the dataset	
					for (int i = 0; i < numCols; ++i)
					{
						//create a new instance of the DataGridColoredTextBoxColumn class
						aColumnTextColumn = new DataGridColoredTextBoxColumn();
						aColumnTextColumn.HeaderText = this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Columns[i].ColumnName;
						//all columns are read-only except the last_tiebreak_rank column
						if (aColumnTextColumn.HeaderText != "last_tiebreak_rank") aColumnTextColumn.ReadOnly = true;
						//assign the mappingname property the data sets column name
						aColumnTextColumn.MappingName = this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Columns[i].ColumnName;
						//add the datagridcoloredtextboxcolumn object to the data grid table style object
						tableStyle.GridColumnStyles.Add(aColumnTextColumn);
						//set wider width for some columns
						switch (aColumnTextColumn.HeaderText)
						{
							case "scenario_id":
								aColumnTextColumn.Width = 150;
								break;
							case "Description":
								aColumnTextColumn.Width = 475;
								break;
						}
					}
					dataGrid1.BackgroundColor = frmMain.g_oGridViewBackgroundColor;
					dataGrid1.BackColor = frmMain.g_oGridViewRowBackgroundColor;
					if (frmMain.g_oGridViewFont != null) dataGrid1.Font = frmMain.g_oGridViewFont;

					// make the dataGrid use our new tablestyle and bind it to our table
					this.dataGrid1.TableStyles.Clear();
					this.dataGrid1.TableStyles.Add(tableStyle);

					// If this is a copied scenario, we will have a reference form to get the values
					if (p_bScenarioCopy == true)
					{
						if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oLastTieBreakRankItem_Collection != null)
						{
							OptimizerScenarioItem.LastTieBreakRankItem_Collection oLastTieBreakRankItem_Collection = ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oLastTieBreakRankItem_Collection;
							for (int i = 0; i < firstView.Count - 1; ++i)
							{
								for (x = 0; x <= oLastTieBreakRankItem_Collection.Count - 1; x++)
								{
									if (oLastTieBreakRankItem_Collection.Item(x).RxPackage.Equals(firstView[i]["rxpackage"]))
									{
										firstView[i]["last_tiebreak_rank"] = oLastTieBreakRankItem_Collection.Item(x).LastTieBreakRank;
									}
								}
							}
						}
					}
					this.dataGrid1.DataSource = firstView;
					this.dataGrid1.Expand(-1);
				}
				this.m_OleDbRxConn.Close();
				conn.Close();
			}
		}

		private void dataGrid1_Click(object sender, System.EventArgs e)
		{
			
		}

		private void dataGrid1_CurrentCellChanged(object sender, System.EventArgs e)
		{

		}
		public int Val_Last_Tiebreak_Rank(bool p_bDisplayMessage)
		{
			int x;
			int y;
			ado_data_access p_ado = new ado_data_access();
			if (this.m_DataSet.Tables[this.strRxPackageTableName].Rows.Count==0)
			{
				if (p_bDisplayMessage) MessageBox.Show("Run Scenario Failed: No treatments defined");
				return -1;
			}
			//if (((frmScenario)this.ParentForm).m_bSave)
			 //   this.savevalues();

            for (x = this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows.Count - 1;
				x >= 0; x--)
			{
                //check for null value
                if (m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows[x]["last_tiebreak_rank"] == System.DBNull.Value)
                {
                    if (p_bDisplayMessage) MessageBox.Show("Run Scenario Failed: Last Tie-Break Rankings cannot be null in value");
                    return -3;
                }
				//check for duplicates
				for (y=x-1;y >= 0; y--)
				{
                    if (this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows[x]["last_tiebreak_rank"].ToString().Trim() ==
                        this.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows[y]["last_tiebreak_rank"].ToString().Trim())
					{
                        if (p_bDisplayMessage) MessageBox.Show("Run Scenario Failed: Last Tie-Break Rankings must be unique");
						return -2;
					}
				}
			}

			return 0;

		}
	
		public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
		
	}
	public class DataGridColoredTextBoxColumn : DataGridTextBoxColumn
	{
		protected override void Paint(System.Drawing.Graphics g, System.Drawing.Rectangle bounds, System.Windows.Forms.CurrencyManager source, int rowNum, System.Drawing.Brush backBrush, System.Drawing.Brush foreBrush, bool alignToRight)
		{
			
			// color only the columns that can be edited by the user
			try
			{
                if (this.HeaderText == "last_tiebreak_rank")
					{
						// could be as simple as
						// or something fnacier...
						backBrush = new System.Drawing.Drawing2D.LinearGradientBrush(bounds,
							Color.FromArgb(255, 200, 200), 
							Color.FromArgb(128, 20, 20),
							System.Drawing.Drawing2D.LinearGradientMode.BackwardDiagonal);
						foreBrush = new SolidBrush(Color.White);
					}
				//}
			}
			catch { /* empty catch */ }
			finally
			{
				// make sure the base class gets called to do the drawing with
				// the possibly changed brushes
				base.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight);
			}
		}
		
	}
}

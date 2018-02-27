using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario_treatment_intensity.
	/// </summary>
	public class uc_scenario_treatment_intensity : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.DataGrid dataGrid1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.ImageList imgSize;
		public string strImagePath="";
		public string m_strProjectDirectory;
		public int m_intFullHt = 280;
		public System.Data.OleDb.OleDbDataAdapter m_OleDbDataAdapter;
		public System.Data.DataSet m_DataSet;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionMaster;
		public System.Data.OleDb.OleDbConnection m_OleDbRxConn;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionScenario;
		public System.Data.OleDb.OleDbCommand m_OleDbCommand;
		public System.Data.DataRelation m_DataRelation;
		public System.Data.DataTable m_DataTable;
		public System.Data.DataRow m_DataRow;
		public string strRxTableName;
		public string strRxConn;
		public string strRxMDBFile;
		public string strScenarioId;
		private FIA_Biosum_Manager.frmCoreScenario _frmScenario=null;
		

		public uc_scenario_treatment_intensity()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
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

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(uc_scenario_treatment_intensity));
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
			this.Resize += new System.EventHandler(this.uc_scenario_treatment_intensity_Resize);
			((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void uc_scenario_treatment_intensity_Resize(object sender, System.EventArgs e)
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
		    int x=0;
			string strSQL="";
			ado_data_access p_ado = new ado_data_access();
            p_ado.m_intError=0;
			try
			{
				for (x=0 ; x <=this.m_DataSet.Tables["scenario_rx_intensity"].Rows.Count-1; x++)
				{
                    if (m_DataSet.Tables["scenario_rx_intensity"].Rows[x]["rx_intensity"] != System.DBNull.Value &&
                        m_DataSet.Tables["scenario_rx_intensity"].Rows[x]["rx_intensity"].ToString().Trim().Length > 0)
                    {
                        strSQL = "UPDATE scenario_rx_intensity SET rx_intensity = " +
                            this.m_DataSet.Tables["scenario_rx_intensity"].Rows[x]["rx_intensity"] +
                            " WHERE TRIM(rx) = '" + this.m_DataSet.Tables["scenario_rx_intensity"].Rows[x]["rx"].ToString().Trim() + "';";
                    }
                    else
                    {
                        strSQL = "UPDATE scenario_rx_intensity SET rx_intensity = null " +
                                 " WHERE TRIM(rx) = '" + this.m_DataSet.Tables["scenario_rx_intensity"].Rows[x]["rx"].ToString().Trim() + "';";
                    }
					p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
					if (p_ado.m_intError < 0) break;

				}
			}
			catch (Exception caught)
			{
				MessageBox.Show("Function: uc_scenario_treatment_intensity.savevalues ErrMsg:" + caught.Message + " Failed updating scenario_rx_intensity table with intensity ratings");
			}
			x=p_ado.m_intError;
			p_ado=null;
			return x;

			
		}

        public void loadgrid(bool p_bScenarioCopy)
		{
			string[]  strDeleteSQL = new string[25];
			string strSQL="";
			int intArrayCount;
			int x=0;
			string strConn="";

			string strScenarioMDB = "";

			ado_data_access p_ado = new ado_data_access();

			this.strScenarioId = ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();

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
				                            ref this.strRxTableName,ref this.strRxConn,
                                            "Treatment Prescriptions",  
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
            strSQL = "select * from " + this.strRxTableName;
			p_ado.SqlQueryReader(this.m_OleDbRxConn, strSQL);
			
			/********************************************************************************
			 **insert records into the scenario_rx_intensity table from the master rx table
			 ********************************************************************************/
			if (p_ado.m_intError == 0)
			{
				this.m_DataSet = new System.Data.DataSet();
				this.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
				while (p_ado.m_OleDbDataReader.Read())
				{
					strSQL = "select * from scenario_rx_intensity " + 
						" where scenario_id = '" + this.strScenarioId + "' and " + 
						"rx = '" + p_ado.m_OleDbDataReader["rx"].ToString()  + "';";
					this.m_OleDbCommand = this.m_OleDbConnectionScenario.CreateCommand();
					this.m_OleDbCommand.CommandText = strSQL;
					this.m_OleDbDataAdapter.SelectCommand = this.m_OleDbCommand;
					this.m_OleDbDataAdapter.Fill(this.m_DataSet,"scenario_rx_intensity");

					/*******************************************************************************
					 **if the master treatment record is not found in the scenario db than insert it
					 *******************************************************************************/
					if (this.m_DataSet.Tables["scenario_rx_intensity"].Rows.Count == 0)
					{
						strSQL = "INSERT INTO scenario_rx_intensity (scenario_id," +
							"rx, rx_intensity) VALUES " + 
							"('" + this.strScenarioId + "'," + 
							"'" + p_ado.m_OleDbDataReader["rx"].ToString() + "'," + 
							"0);";

						p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strSQL);
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
				strSQL = "select * from scenario_rx_intensity where scenario_id = '" + this.strScenarioId + "';";
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
						strSQL = "select * from " + this.strRxTableName + 
							" where rx = '" + p_ado.m_OleDbDataReader["rx"].ToString()  + "';";
						this.m_OleDbCommand = this.m_OleDbRxConn.CreateCommand();
						this.m_OleDbCommand.CommandText = strSQL;
						this.m_OleDbDataAdapter.SelectCommand = this.m_OleDbCommand;
						this.m_OleDbDataAdapter.Fill(this.m_DataSet, this.strRxTableName);
						if (this.m_DataSet.Tables[this.strRxTableName].Rows.Count == 0)
						{
							strDeleteSQL[intArrayCount] = "DELETE FROM scenario_rx_intensity " +
								" WHERE scenario_id = '" + this.strScenarioId + "'" + 
								" AND rx = '" + p_ado.m_OleDbDataReader["rx"] + "';";
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
					if (intArrayCount > 0 )
					{
						for (x=0; x<=intArrayCount-1; x++)
						{
							p_ado.SqlNonQuery(this.m_OleDbConnectionScenario,strDeleteSQL[x].ToString());
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
				this.m_OleDbCommand.CommandText = "select * from " + this.strRxTableName;
				this.m_OleDbDataAdapter.SelectCommand = this.m_OleDbCommand;
				this.m_OleDbDataAdapter.Fill(this.m_DataSet,this.strRxTableName);
                     
				this.m_OleDbCommand = this.m_OleDbConnectionScenario.CreateCommand();
				strSQL="";
				for (x=0 ; x <=this.m_DataSet.Tables[this.strRxTableName].Rows.Count-1; x++)
				{
					if (this.m_DataSet.Tables[this.strRxTableName].Rows[x]["rx"].ToString().Length > 0)
					{
						strSQL = "select scenario_id,rx,rx_intensity from scenario_rx_intensity where scenario_id = '" + this.strScenarioId + "';";
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
				this.m_OleDbCommand.CommandText = "select scenario_id,rx,rx_intensity from scenario_rx_intensity where scenario_id = '" + this.strScenarioId + "';";
				this.m_OleDbDataAdapter.SelectCommand = this.m_OleDbCommand;
				this.m_OleDbDataAdapter.Fill(this.m_DataSet,"scenario_rx_intensity");

				/*****************************************************************************
				 **add the description column to the scenario rx intensity dataset
				 *****************************************************************************/
				this.m_DataSet.Tables["scenario_rx_intensity"].Columns.Add("Description");
                
              /********************************************************************************
			   **add the treatment description value to the scenrario rx intensity data set.
			   **the description is only in the master rx table and is added to the
			   **scenario rx intensity data set for information purposes.
			   ********************************************************************************/

				/***********************************************************************************
				 **for loop through the master db rx dataset adding the description field to the
				 **scenenario db scenario_rx_intensity dataset
				 ***********************************************************************************/
				for (x=0 ; x <=this.m_DataSet.Tables[this.strRxTableName].Rows.Count-1; x++)
				{
					if (this.m_DataSet.Tables[this.strRxTableName].Rows[x]["rx"].ToString().Length > 0)
					{
						/***************************************************************************************
						 **build the expression to filter only the scenario row that meets the expression
						 ***************************************************************************************/
						strSQL = "rx = '" + this.m_DataSet.Tables[this.strRxTableName].Rows[x]["rx"] + "'";

						/***************************************************************************************
						 **create a datarow that will hold the results from the query expression
						 ***************************************************************************************/
						System.Data.DataRow[] dr1;
						dr1 = this.m_DataSet.Tables["scenario_rx_intensity"].Select(strSQL);
						
						/***************************************************************************************
						 **check to see if it found the master rx treatment in the sceanrio rx intensity dataset
						 ***************************************************************************************/
						if (dr1.Length != 0)
						{
							/***************************************************************************************
							 **it found it, loop through the dataset and find the row that matches the row 
							 **returned from the search expression
							 ***************************************************************************************/
							for (int y = 0; y <= this.m_DataSet.Tables["scenario_rx_intensity"].Rows.Count -1  ; y++)
							{
								if (dr1[0]["rx"] == 
									this.m_DataSet.Tables["scenario_rx_intensity"].Rows[y]["rx"])
								{
									/**********************************************************************************
									 **update the description row/column with the master db rx table description value
									 **********************************************************************************/
									this.m_DataSet.Tables["scenario_rx_intensity"].Rows[y]["description"] = 
										this.m_DataSet.Tables[this.strRxTableName].Rows[x]["description"];
									break;
								}
							}
							this.m_DataSet.Tables["scenario_rx_intensity"].AcceptChanges();

						}
					
					}	
				}
                
                 
        /**************************************************************************************************
         **place the dataset table into a view class so as to not allow new records to be appended
         **************************************************************************************************/
				DataView firstView = new DataView(this.m_DataSet.Tables["scenario_rx_intensity"]);
				firstView.AllowNew = false;       //cannot append new records
				firstView.AllowDelete = false;    //cannot delete records

				/***************************************************************
				 **custom define the grid style
				 ***************************************************************/
				DataGridTableStyle tableStyle = new DataGridTableStyle();

				/***********************************************************************
				 **map the data grid table style to the scenario rx intensity dataset
				 ***********************************************************************/
				tableStyle.MappingName = "scenario_rx_intensity";   
				tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
				tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
				tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
				tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;

				/******************************************************************************
				 **since the dataset has things like field name and number of columns,
				 **we will use those to create new columnstyles for the columns in our grid
				 ******************************************************************************/
				//get the number of columns from the scenario_rx_intensity data set
				int numCols = this.m_DataSet.Tables["scenario_rx_intensity"].Columns.Count;
                
				/***********************************************************************************
				 **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
				 ***********************************************************************************/
				DataGridColoredTextBoxColumn aColumnTextColumn ;
                    
				//loop through all the columns in the dataset	
				for(int i = 0; i < numCols; ++i)
				{
					//create a new instance of the DataGridColoredTextBoxColumn class
					aColumnTextColumn = new DataGridColoredTextBoxColumn();
					aColumnTextColumn.HeaderText = this.m_DataSet.Tables["scenario_rx_intensity"].Columns[i].ColumnName;
					//all columns are read-only except the rx_intensity column
					if (aColumnTextColumn.HeaderText != "rx_intensity") aColumnTextColumn.ReadOnly=true;
					//assign the mappingname property the data sets column name
					aColumnTextColumn.MappingName = this.m_DataSet.Tables["scenario_rx_intensity"].Columns[i].ColumnName;
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
                dataGrid1.BackgroundColor=frmMain.g_oGridViewBackgroundColor;   	
				dataGrid1.BackColor=frmMain.g_oGridViewRowBackgroundColor;
				if (frmMain.g_oGridViewFont != null) dataGrid1.Font = frmMain.g_oGridViewFont;

				// make the dataGrid use our new tablestyle and bind it to our table
				this.dataGrid1.TableStyles.Clear();
				this.dataGrid1.TableStyles.Add(tableStyle);
                
                // If this is a copied scenario, we will have a reference form to get the values
                if (p_bScenarioCopy == true)
                {
                    if (ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oRxIntensityItem_Collection != null)
                    {
                        CoreAnalysisScenarioItem.RxIntensityItem_Collection oRxIntensityItem_Collection = ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oRxIntensityItem_Collection;
                        for (int i = 0; i < firstView.Count - 1; ++i)
                        {
                            for (x = 0; x <= oRxIntensityItem_Collection.Count - 1; x++)
                            {
                                if (oRxIntensityItem_Collection.Item(x).Rx.Equals(firstView[i]["rx"]))
                                {
                                    firstView[i]["rx_intensity"] = oRxIntensityItem_Collection.Item(x).RxIntensity;
                                }
                             }
                        }
                    }
                }
				this.dataGrid1.DataSource = firstView;
				this.dataGrid1.Expand(-1);
				


			}
           
				   
		

			
		}

		private void dataGrid1_Click(object sender, System.EventArgs e)
		{
			
		}

		private void dataGrid1_CurrentCellChanged(object sender, System.EventArgs e)
		{

		}
		public int Val_Intensity(bool p_bDisplayMessage)
		{
			int x;
			int y;
			ado_data_access p_ado = new ado_data_access();
			if (this.m_DataSet.Tables[this.strRxTableName].Rows.Count==0)
			{
				if (p_bDisplayMessage) MessageBox.Show("Run Scenario Failed: No treatments defined");
				return -1;
			}
			//if (((frmScenario)this.ParentForm).m_bSave)
			 //   this.savevalues();
             
			for (x=this.m_DataSet.Tables["scenario_rx_intensity"].Rows.Count-1;
				x >= 0; x--)
			{
                //check for null value
                if (m_DataSet.Tables["scenario_rx_intensity"].Rows[x]["rx_intensity"] == System.DBNull.Value)
                {
                    if (p_bDisplayMessage) MessageBox.Show("Run Scenario Failed: Treatment intensity ratings cannot be null in value");
                    return -3;
                }
				//check for duplicates
				for (y=x-1;y >= 0; y--)
				{
					if (this.m_DataSet.Tables["scenario_rx_intensity"].Rows[x]["rx_intensity"].ToString().Trim() ==
						this.m_DataSet.Tables["scenario_rx_intensity"].Rows[y]["rx_intensity"].ToString().Trim())
					{
						if (p_bDisplayMessage) MessageBox.Show("Run Scenario Failed: Treatment intensity ratings must be unique");
						return -2;
					}
				}
			}

			return 0;

		}

		private void cmdRxIntensity_Click(object sender, System.EventArgs e)
		{
			((frmCoreScenario)this.ParentForm).RulesRepositionControls();
		}

		private void grpboxRxIntensity_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void lblTitle_Click(object sender, System.EventArgs e)
		{
		
		}
	
		public FIA_Biosum_Manager.frmCoreScenario ReferenceCoreScenarioForm
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
					if(this.HeaderText == "rx_intensity")
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

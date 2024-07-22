using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_gis_psite.
	/// </summary>
	public class uc_gis_psite : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.DataGrid m_dg;
		public System.Windows.Forms.Button btnDelete;
		private System.Windows.Forms.Button btnNew;
		private System.Windows.Forms.Button btnEdit;
		public System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnHelp;

		private FIA_Biosum_Manager.Datasource m_DataSource;
		private string m_strPSiteTable;
		private string m_strProjDir;
		private string m_strFieldList="";
		private FIA_Biosum_Manager.ado_data_access m_ado;
        private SQLite.ADO.DataMgr m_dataMgr;
		private string m_strConn;
		private System.Data.DataView m_dv;
		public int m_intIndex=0;
		private int m_intCurrRow=0;
		public int m_intError=0;
		
		private System.Windows.Forms.VScrollBar m_vScrollBar;
		private FIA_Biosum_Manager.frmDialog m_frmDialog;

		private const int MENU_FILTERBYVALUE=0;
		private const int MENU_FILTERBYENTEREDVALUE=1;
		private const int MENU_REMOVEFILTER = 2;
		private const int MENU_UNIQUEVALUES = 3;
		private const int MENU_MODIFY = 5;
		private const int MENU_DELETE=6;
		private const int MENU_SELECTALL = 8;
		private const int MENU_IDXASC = 10;
		private const int MENU_IDXDESC = 11;
		private const int MENU_REMOVEIDX = 12;
		private const int MENU_MAX = 14;
		private const int MENU_MIN = 15;
		private const int MENU_AVG = 16;
		private const int MENU_SUM = 17;
		private const int MENU_COUNTBYVALUE=18;


		

		private int m_intPopupColumn=0;
		public System.Windows.Forms.ContextMenu m_mnuDataGridPopup;
		//private bool m_bDelete=false;
		private string m_strDeletedList="";
		private int m_intDeletedCount=0;
		private string m_strColumnFilterList="";
		private string m_strColumnSortList="";
		private int m_intScrollValue=0;
		//private int m_intScrollMove=0;
		private double m_dblOldPerc=0;
		private double m_dblNewPerc=0;
		private int m_intMaxSize=0;
		//private bool m_bScrollControls=true;

		private System.Data.DataTable m_dtTableSchema;

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultDatabaseXPSFile;

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_gis_psite(string p_strProjDir)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			this.m_strProjDir = p_strProjDir;


			m_DataSource = new Datasource();
			m_DataSource.LoadTableColumnNamesAndDataTypes=false;
			m_DataSource.LoadTableRecordCount=false;
			m_DataSource.m_strDataSourceMDBFile = m_strProjDir.Trim() + "\\db\\project.mdb";
			m_DataSource.m_strDataSourceTableName = "datasource";
			m_DataSource.m_strScenarioId="";
			m_DataSource.populate_datasource_array();

			this.m_strPSiteTable = this.m_DataSource.getValidDataSourceTableName("PROCESSING SITES");
			if (this.m_strPSiteTable.Trim().Length == 0) 
			{
				MessageBox.Show("!!Could Not Locate Processing Sites Table!!","FIA BioSum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}

            this.m_dataMgr = new SQLite.ADO.DataMgr();
            this.m_strConn = m_dataMgr.GetConnectionString(this.m_DataSource.getFullPathAndFile("PROCESSING SITES"));
            this.m_dataMgr.OpenConnection(this.m_strConn);
            if (m_dataMgr.m_intError != 0)
            {
                this.m_intError = m_dataMgr.m_intError;
                m_dataMgr.CloseConnection(this.m_dataMgr.m_Connection);
                m_dataMgr = null;
            }
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(this.m_strConn))
            {
                conn.Open();
                if (m_dataMgr.m_intError !=0)
                {
                    this.m_intError = m_dataMgr.m_intError;
                    m_dataMgr = null;
                }
            }
			
			// TODO: Add any initialization after the InitializeComponent call
            this.m_oEnv = new env();
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
			string strColumnName="";
			this.InitializePopup();
				                                         
			//this.m_ado.m_DataSet = new DataSet("processing_site");
			//this.m_ado.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
			//this.m_ado.m_strSQL = "select * from " + this.m_strPSiteTable + " order by psite_id;";
			//this.m_dtTableSchema = this.m_ado.getTableSchema(this.m_ado.m_OleDbConnection,
			//	this.m_ado.m_OleDbTransaction,
			//	this.m_ado.m_strSQL);
            m_dataMgr.m_DataSet = new DataSet("processing_site");
            m_dataMgr.m_DataAdapter = new System.Data.SQLite.SQLiteDataAdapter();
            m_dataMgr.m_strSQL = "select * from " + this.m_strPSiteTable + " order by psite_id;";
            this.m_dtTableSchema = m_dataMgr.getTableSchema(m_dataMgr.m_Connection, m_dataMgr.m_Transaction, m_dataMgr.m_strSQL);

            if (m_dataMgr.m_intError == 0)
			{

				for (int x=0; x<=m_dtTableSchema.Rows.Count-1;x++)
				{
					if (this.m_strFieldList.Trim().Length == 0)
					{
						if (m_dtTableSchema.Rows[x]["columnname"].ToString().Trim().ToUpper() != "PSITE_ID")
						{
							this.m_strFieldList = m_dtTableSchema.Rows[x]["columnname"].ToString().Trim();
						}
					}
					else
					{	
						if (m_dtTableSchema.Rows[x]["columnname"].ToString().Trim().ToUpper() != "PSITE_ID")
						{
							this.m_strFieldList += "," + m_dtTableSchema.Rows[x]["columnname"].ToString().Trim();
						}
					}
				}
                this.InitializeTransactionCommands();

				m_dataMgr.m_strSQL = "select * from " + this.m_strPSiteTable + " order by psite_id;";
				m_dataMgr.m_Command = m_dataMgr.m_Connection.CreateCommand();
                m_dataMgr.m_Command.CommandText = m_dataMgr.m_strSQL;
                m_dataMgr.m_DataAdapter.SelectCommand = m_dataMgr.m_Command;
                m_dataMgr.m_DataAdapter.SelectCommand.Transaction = m_dataMgr.m_Transaction;
				try 
				{

                    m_dataMgr.m_DataAdapter.Fill(m_dataMgr.m_DataSet,"processing_site");
					this.m_dv = new DataView(m_dataMgr.m_DataSet.Tables["processing_site"]);
				
					this.m_dv.AllowNew = false;       //cannot append new records
					this.m_dv.AllowDelete = false;    //cannot delete records
					this.m_dv.AllowEdit = false;
					this.m_dg.CaptionText = "processing_site";
					this.m_dg.BackgroundColor=frmMain.g_oGridViewBackgroundColor;
					/***********************************************************************************
					 **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
					 ***********************************************************************************/
					PSite_DataGridColoredTextBoxColumn aColumnTextColumn ;


					/***************************************************************
					 **custom define the grid style
					 ***************************************************************/
					DataGridTableStyle tableStyle = new DataGridTableStyle();

					/***********************************************************************
					 **map the data grid table style to the scenario rx intensity dataset
					 ***********************************************************************/
					tableStyle.MappingName = "processing_site";
					tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
					tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
					tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
					tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
   
					/******************************************************************************
					 **since the dataset has things like field name and number of columns,
					 **we will use those to create new columnstyles for the columns in our grid
					 ******************************************************************************/
					//get the number of columns from the scenario_rx_intensity data set
					int numCols = m_dataMgr.m_DataSet.Tables["processing_site"].Columns.Count;
                
                    
					/************************************************
					 **loop through all the columns in the dataset	
					 ************************************************/
					for(int i = 0; i < numCols; ++i)
					{
						strColumnName = m_dataMgr.m_DataSet.Tables["processing_site"].Columns[i].ColumnName;
					
						if (strColumnName.Trim().ToUpper() == "CONVERT_TO_FIA_SPCD")
						{
							/******************************************************************
							**create a new instance of the DataGridColoredTextBoxColumn class
							******************************************************************/
							aColumnTextColumn = new PSite_DataGridColoredTextBoxColumn(true,true,this);
							aColumnTextColumn.TextBox.MaxLength=3;
							aColumnTextColumn.ReadOnly=false;
						}
						else if (strColumnName.Trim().ToUpper() == "COMMENTS" || strColumnName.Trim().ToUpper() == "COMMON_NAME")
						{
							aColumnTextColumn = new PSite_DataGridColoredTextBoxColumn(true,false,this);
							aColumnTextColumn.TextBox.MaxLength=50;
							aColumnTextColumn.ReadOnly=false;
						}
						else
						{
							/******************************************************************
							 **create a new instance of the DataGridColoredTextBoxColumn class
							 ******************************************************************/
							aColumnTextColumn = new PSite_DataGridColoredTextBoxColumn(false,false,this);
							/***********************************
							 **all columns are read-only except
							 **the edit columns
							 ***********************************/
							aColumnTextColumn.ReadOnly=true;

						}
					
						aColumnTextColumn.HeaderText = strColumnName;
				 				    
						/********************************************************************
						 **assign the mappingname property the data sets column name
						 ********************************************************************/
						aColumnTextColumn.MappingName = strColumnName;
						//aColumnTextColumn
						aColumnTextColumn.TextBox.ContextMenu =  new ContextMenu(); //this.m_mnuDataGridPopup;
						aColumnTextColumn.TextBox.ContextMenu = this.m_mnuDataGridPopup;
						aColumnTextColumn.TextBox.MouseDown += new System.Windows.Forms.MouseEventHandler(this.m_dg_TextBox_MouseDown);
						/********************************************************************
						 **add the datagridcoloredtextboxcolumn object to the data grid 
						 **table style object
						 ********************************************************************/
						tableStyle.GridColumnStyles.Add(aColumnTextColumn);

					}
					/*********************************************************************
					 ** make the dataGrid use our new tablestyle and bind it to our table
					 *********************************************************************/

					if (frmMain.g_oGridViewFont!=null)this.m_dg.Font = frmMain.g_oGridViewFont;

					this.m_dg.TableStyles.Clear();
					this.m_dg.TableStyles.Add(tableStyle);

					this.m_dg.DataSource = this.m_dv;  

					if (m_dataMgr.m_DataSet.Tables["processing_site"].Rows.Count > 0)
					{
						this.btnDelete.Enabled=true;
						this.btnEdit.Enabled=true;
					}

				

					this.m_dg.Expand(-1);
				}
				catch (Exception e)
				{
					MessageBox.Show(e.Message,"Table",MessageBoxButtons.OK,MessageBoxIcon.Error);
					this.m_intError=-1;
					this.m_ado.m_OleDbConnection.Close();
                    m_dataMgr.m_Connection.Close();
                    m_dataMgr.m_Connection = null;
                    m_dataMgr.m_DataSet.Clear();
                    m_dataMgr.m_DataSet= null;
                    m_dataMgr.m_DataAdapter.Dispose();
                    m_dataMgr.m_DataAdapter = null;
					m_dataMgr = null;
					return;
				}

				if (m_dataMgr.m_DataSet.Tables["processing_site"].Rows.Count == 0)
				{
					this.m_intCurrRow = 1;
				}
				else
				{
					this.m_intCurrRow = 1;
				}
			

				//event handler to keep track of current row and cell movement
				this.m_dg.CurrentCellChanged += new
					System.EventHandler(this.m_dg_CurrentCellChanged);

			}
			if (m_dataMgr.m_intError < 0)
			{
				this.ParentForm.Close();
			}
		}
		private void InitializePopup()
		{
			this.m_mnuDataGridPopup = new ContextMenu();
			this.m_mnuDataGridPopup.MenuItems.Add("Filter By Value",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Filter By Entered Value", new EventHandler(Filter_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Remove Filter",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Unique Values",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;
			this.m_mnuDataGridPopup.MenuItems.Add("-");
			this.m_mnuDataGridPopup.MenuItems.Add("Modify",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=false;
			this.m_mnuDataGridPopup.MenuItems.Add("Delete",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems[MENU_DELETE].Enabled=false;
			this.m_mnuDataGridPopup.MenuItems.Add("-");
			this.m_mnuDataGridPopup.MenuItems.Add("Select All",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems[MENU_SELECTALL].Enabled=true;
			this.m_mnuDataGridPopup.MenuItems.Add("-");
			this.m_mnuDataGridPopup.MenuItems.Add("Index Ascending",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Index Descending",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Remove Index",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=false;
			this.m_mnuDataGridPopup.MenuItems.Add("-");
			this.m_mnuDataGridPopup.MenuItems.Add("Maximum",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Minimum",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Average",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Sum",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Count By Value",new EventHandler(PopUp_Clicked));
			
			

			
		}
		private void PopUp_Clicked(object sender,EventArgs e)
		{
			MenuItem miClicked = (MenuItem)sender;
			string strItem = miClicked.Text;
			string strCellValue = "";
			string strCol = "";
			string strFilter = "";
			string strExp="";
			int intLeft=0;
			int intTop=0;
			int x=0;
			int y=0;
			System.Windows.Forms.DialogResult dlgResult;
			System.Windows.Forms.CurrencyManager p_cm;
			System.Data.DataView p_dv;
			int intCurrRow=0;

			switch (strItem.ToUpper().Trim())
			{
				case "FILTER BY VALUE":

					strCellValue = this.m_dg[this.m_dg.CurrentCell.RowNumber,this.m_dg.CurrentCell.ColumnNumber].ToString();
					this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=true;
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					
				switch (this.m_dg[this.m_dg.CurrentCell.RowNumber,this.m_dg.CurrentCell.ColumnNumber].GetType().FullName.ToString().Trim())
				{
					case "System.String":
						if (strCellValue.Trim().Length==0)
						{
							strFilter = strCol + " IS NULL";
						}
						else

							strFilter =  strCol + "='" + strCellValue + "'";
						break;
					default:
						if (strCellValue.Trim().Length==0)
						{
							strFilter = strCol + " IS NULL";
						}
						else
							strFilter = strCol + "=" + strCellValue;

						break;
				}
					this.m_dv.RowFilter = strFilter;

					if (this.m_strColumnFilterList.Trim().Length > 0) this.m_strColumnFilterList+=  strCol.Trim() + ",";
					else this.m_strColumnFilterList = strCol.Trim() + ",";

					break;
				case "UNIQUE VALUES":
					this.uniquevalues();
					break;
				case "MODIFY":
					//---------------form----------------//
					//declare form and instatiate
					FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();

					
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					if (strCol.ToUpper()== "CONVERT_TO_FIA_SPCD")
					{

						//------------text box------------//
						//instatiate numeric text class
						FIA_Biosum_Manager.txtNumeric p_txtDefault = new FIA_Biosum_Manager.txtNumeric(3,0);

						
						//define form properties
						frmTemp.Width = 200;
						frmTemp.Height = 200;
						frmTemp.WindowState = System.Windows.Forms.FormWindowState.Normal;
						frmTemp.MaximizeBox = false;
						frmTemp.MinimizeBox = false;
						frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
						frmTemp.Text = "Modify";
		                    
						//define numeric text class properties
						p_txtDefault.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
						p_txtDefault.Name = "txtModify";
						p_txtDefault.TabIndex = 0;
						p_txtDefault.Tag = "";
						p_txtDefault.Visible = true;
						p_txtDefault.Enabled = true;
						frmTemp.Controls.Add(p_txtDefault);  //add the text box to the form
						p_txtDefault.Height = 100;

						p_txtDefault.Top = (int)(frmTemp.ClientSize.Height * .50) - (int)(p_txtDefault.Height * .50);
						p_txtDefault.Width = frmTemp.ClientSize.Width - 20;
						p_txtDefault.Left = 10;

						p_txtDefault.bEdit=true;
						p_txtDefault.ReadOnly=false;
						p_txtDefault.Text = "";
						intLeft = p_txtDefault.Left;
						intTop = p_txtDefault.Top;
						frmTemp.txtNumeric = p_txtDefault;
						frmTemp.strCallingFormType="TN";
					}
					else if (strCol.ToUpper() == "COMMENTS" || strCol.ToUpper() == "COMMON_NAME")
					{
						//------------text box------------//
						//instatiate numeric text class
						System.Windows.Forms.TextBox p_textbox = new TextBox();

						
						//define form properties
						frmTemp.Width = 200;
						frmTemp.Height = 200;
						frmTemp.WindowState = System.Windows.Forms.FormWindowState.Normal;
						frmTemp.MaximizeBox = false;
						frmTemp.MinimizeBox = false;
						frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
						frmTemp.Text = "Modify";
		                    
						//define numeric text class properties
						p_textbox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
						p_textbox.Name = "txtModify";
						p_textbox.TabIndex = 0;
						p_textbox.Tag = "";
						p_textbox.Visible = true;
						p_textbox.Enabled = true;
						frmTemp.Controls.Add(p_textbox);  //add the text box to the form
						p_textbox.Height = 100;

						p_textbox.Top = (int)(frmTemp.ClientSize.Height * .50) - (int)(p_textbox.Height * .50);
						p_textbox.Width = frmTemp.ClientSize.Width - 20;
						p_textbox.Left = 10;
						
						p_textbox.MaxLength = 50;
						
						p_textbox.Text = "";
						intLeft = p_textbox.Left;
						intTop = p_textbox.Top;
						frmTemp.txtBox = p_textbox;
						frmTemp.strCallingFormType="TS";
					}
					

					//-----------label---------------//
					System.Windows.Forms.Label lblTemp = new Label();
					lblTemp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
					lblTemp.Location = new System.Drawing.Point(168, 53);
					lblTemp.Name = "lblTemp";
					lblTemp.Size = new System.Drawing.Size(128, 15);
					lblTemp.TabIndex = 1;
					lblTemp.Text = "New Value";
					frmTemp.Controls.Add(lblTemp);
					lblTemp.Top = intTop - lblTemp.Height - 10;
					lblTemp.Left = intLeft;
					lblTemp.Visible=true;


					//----------------OK button-----------------//
					System.Windows.Forms.Button btnTempOK = new Button();
					btnTempOK.Location = new System.Drawing.Point(392, 328);
					btnTempOK.Name = "btnOK";
					btnTempOK.Size = new System.Drawing.Size(80, 32);
					btnTempOK.TabIndex = 2;
					btnTempOK.Text = "OK";
					btnTempOK.Click += new System.EventHandler(frmTemp.btnOK_Click);
					frmTemp.Controls.Add(btnTempOK);
					btnTempOK.Top = intTop + btnTempOK.Height + 10;
					btnTempOK.Left = (int)(frmTemp.Width * .50)  - btnTempOK.Width;
					btnTempOK.Visible=true;

					//----------------Cancel button-----------------//
					System.Windows.Forms.Button btnTempCancel = new Button();
					btnTempCancel.Location = new System.Drawing.Point(392, 328);
					btnTempCancel.Name = "btnCancel";
					btnTempCancel.Size = new System.Drawing.Size(80, 32);
					btnTempCancel.TabIndex = 2;
					btnTempCancel.Text = "CANCEL";
					btnTempCancel.Click += new System.EventHandler(frmTemp.btnCancel_Click);
					frmTemp.Controls.Add(btnTempCancel);
					btnTempCancel.Top = intTop + btnTempCancel.Height + 10;
					btnTempCancel.Left = (int)(frmTemp.Width * .50) ;
					btnTempCancel.Visible=true;
					frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
					dlgResult = frmTemp.ShowDialog();
					if (dlgResult.ToString().Trim().ToUpper() == "OK" )
					{
						try
						{
							p_cm = (CurrencyManager)this.BindingContext[this.m_dg.DataSource,this.m_dg.DataMember];
							p_dv = (DataView)p_cm.List;
							intCurrRow= this.m_intCurrRow-1;
							y = p_dv.Count;

							/***************************************************
								 **if the column being modified has a filter or sort
								 **applied to it then the rows will change when
								 **the edit value is applied to the row.
								 **Therefore, check if the edited column is also
								 **filtered/sorted to set the variables used to handle
								 **the row changes that will occur when the edit
								 **is applied to the column.
								 *************************************************/
							int intModifyCount=0;
							bool bApplyModifyCount=false;
							string strSearch=strCol.Trim() + ",";
							if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
							{
								if (this.m_strColumnFilterList.IndexOf(strSearch,0) >= 0)
									bApplyModifyCount = true;
							}
							if (bApplyModifyCount==false)
							{
								if (this.m_dv.Sort.ToString().Trim().Length > 0)
								{
									if (this.m_strColumnSortList.IndexOf(strSearch,0) >= 0)
										bApplyModifyCount = true;
								}
							}
							for (x=0; x <= y-1;x++)
							{

								if (m_dg.IsSelected(x)|| x == intCurrRow)
								{
									if (bApplyModifyCount==true)
									{
										m_dg.CurrentRowIndex = x-intModifyCount;
									}
									else
									{
										m_dg.CurrentRowIndex=x;
									}
								
									if (frmTemp.strCallingFormType=="TS")
									{
										m_dg[m_dg.CurrentRowIndex,this.m_intPopupColumn] = frmTemp.txtBox.Text ;

									}
									else if (frmTemp.strCallingFormType=="TN")
									{
										if (frmTemp.txtNumeric.Text.Trim().Length > 0)
										{
											if (frmTemp.txtNumeric.Text.Trim().Length == 1 &&
												frmTemp.txtNumeric.Text.Trim() == ".")
											{
												this.m_dg[x,this.m_intPopupColumn] = System.DBNull.Value;
											}
											else this.m_dg[x,this.m_intPopupColumn] = m_dg[x,this.m_intPopupColumn] = frmTemp.txtNumeric.Text;
										}
										else
										{
											this.m_dg[x,this.m_intPopupColumn] = System.DBNull.Value;
										}
									}
									else if (frmTemp.strCallingFormType=="TD")
									{
										if (frmTemp.txtNumeric.Text.Trim().Length > 0)
										{
											if (frmTemp.txtNumeric.Text.Trim().Length == 1 &&
												frmTemp.txtNumeric.Text.Trim() == ".")
											{
												this.m_dg[x,this.m_intPopupColumn] = System.DBNull.Value;
											}
											else this.m_dg[x,this.m_intPopupColumn] = m_dg[x,this.m_intPopupColumn] = frmTemp.txtNumeric.Text;
										}
										else
										{
											this.m_dg[x,this.m_intPopupColumn] = System.DBNull.Value;
										}
									}
									else if (frmTemp.strCallingFormType=="TM")
									{
										m_dg[x,this.m_intPopupColumn] = frmTemp.txtMoney.Text.Substring(1,frmTemp.txtMoney.Text.Trim().Length - 1);
									}
									if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
									intModifyCount++;

								}
							}
							this.m_dg.SetDataBinding(this.m_dv,"");
							this.m_dg.Update();
							
						}
						catch (Exception err)
						{
							MessageBox.Show(err.Message);
						}

						
					}
					frmTemp = null;
					break;
				case "REMOVE FILTER":
					this.m_dv.RowFilter="";
					this.m_strColumnFilterList="";
					this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;
					break;
				case "INDEX ASCENDING":
					this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=true;
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					this.m_dv.Sort = strCol + " ASC";
					this.m_strColumnSortList = strCol.Trim() + ",";
					break;
				case "INDEX DESCENDING":
					this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=true;
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					this.m_dv.Sort = strCol + " DESC";
					this.m_strColumnSortList = strCol.Trim() + ",";
					break;
				case "REMOVE INDEX":
					this.m_dv.Sort="";
					this.m_strColumnSortList="";
					break;
				case "MAXIMUM":
					try
					{
						strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
						strFilter = "Max(" + strCol + ")";
						this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
						if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
						{
							MessageBox.Show(strCol + " Maximum: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
						}
						else
						{
							MessageBox.Show(strCol + " Maximum: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter, null));
						}
					}
					catch
					{
					}
					break;
				case "MINIMUM":
					try
					{
						strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
						strFilter = "Min(" + strCol + ")";
						this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
						if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
						{
							MessageBox.Show(strCol + " Minimum: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
						}
						else
						{
							MessageBox.Show(strCol + " Minimum: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter, null));
						}
					}
					catch
					{
					}
					break;

				case "AVERAGE":
					try
					{
						strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
						strFilter = "Avg(" + strCol + ")";
						this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
						if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
						{
							MessageBox.Show(strCol + " Average: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
						}
						else
						{
							MessageBox.Show(strCol + " Average: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter, null));
						}
					}
					catch
					{
					}
					break;

				case "SUM":
					try
					{
						strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
						strFilter = "Sum(" + strCol + ")";
						this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
						if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
						{
							MessageBox.Show(strCol + " Sum: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
						}
						else
						{
							MessageBox.Show(strCol + " Sum: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter, null));
						}
					}
					catch
					{
					}
					break;

				case "COUNT BY VALUE":
					try
					{
						strCellValue = string.Format("{0:#.####################}",this.m_dg[this.m_dg.CurrentCell.RowNumber,this.m_dg.CurrentCell.ColumnNumber].ToString());
						strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
						switch (this.m_dg[this.m_dg.CurrentCell.RowNumber,this.m_dg.CurrentCell.ColumnNumber].GetType().FullName.ToString().Trim())
						{
							case "System.String":
								strExp =  strCol + "='" + strCellValue + "'";
								break;
							default:
								strExp = strCol + "=" + strCellValue;
								break;
						}
						strFilter = "Count(" + strCol + ")";
						this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
						if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
						{
							MessageBox.Show(strCol + " Count: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString() + " and " + strExp));
						}
						else
						{
							MessageBox.Show(strCol + " Count: " + this.m_ado.m_DataSet.Tables[0].Compute(strFilter, strExp));
						}
					}
					catch
					{
					}
					break;
				case "SELECT ALL":
					int intRows = this.m_dg.BindingContext[this.m_dg.DataSource,this.m_dg.DataMember].Count;
					for (x=0; x<=intRows-1;x++)
					{
						this.m_dg.Select(x);
					}
					break;
				case "DELETE":
					this.DeleteRecords();

					break;
					
					
		    
			}
		}
		private void Filter_Clicked(object sender,EventArgs e)
		{
			MenuItem miClicked = (MenuItem)sender;
			string strItem = miClicked.Text;
			string strCellValue = "";
			string strCol = "";
			string strDataType="";
			string strFilter = "";
			//string strCount1="";
			string strExp="";
			int intLeft=0;
			int intTop=0;
			int x=0;
			int y=0;
			System.Windows.Forms.DialogResult dlgResult;
			System.Windows.Forms.CurrencyManager p_cm;
			System.Data.DataView p_dv;
			int intCurrRow=0;
			//MessageBox.Show(strItem);
			switch (strItem.ToUpper().Trim())
			{
				case "FILTER BY ENTERED VALUE":
					strCellValue = this.m_dg[this.m_dg.CurrentCell.RowNumber,this.m_dg.CurrentCell.ColumnNumber].ToString();
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					strDataType = this.m_dg[this.m_dg.CurrentCell.RowNumber,this.m_dg.CurrentCell.ColumnNumber].GetType().FullName.ToString().Trim();
				switch (strDataType)
				{
					case "System.String":
						frmDialog frmTemp = new frmDialog();
						frmTemp.Initialize_Filter_Rows_Text_Datatype_User_Control();
						frmTemp.Text = strCol;
						frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
						dlgResult = frmTemp.ShowDialog();
						if (dlgResult.ToString().Trim().ToUpper() == "OK" )
						{
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=true;
							string strOperation = frmTemp.uc_filter_rows_text_datatype1.GetFilterType();
							string strText = frmTemp.uc_filter_rows_text_datatype1.GetText();

							switch (strOperation)
							{
								case "EQUAL":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NULL";
									else
										strFilter =  strCol + "='" + strText + "'";
									break;
								case "NOTEQUAL":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NOT NULL";
									else
										strFilter =  strCol + "<>'" + strText + "'";
									break;
								case "START":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NULL";
									else
										strFilter =  strCol + " LIKE '" + strText.Trim() + "*'";
									break;
								case "NOTSTART":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NOT NULL";
									else
										strFilter =  strCol + " NOT LIKE '" + strText.Trim() + "*'";
									break;
								case "CONTAIN":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NULL";
									else
										strFilter =  strCol + " LIKE '%" + strText.Trim() + "%'";
									break;
								case "NOTCONTAIN":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NOT NULL";
									else
										strFilter =  strCol + " NOT LIKE '%" + strText.Trim() + "%'";
									break;
								case "END":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NULL";
									else
										strFilter =  strCol + " LIKE '*" + strText.Trim() + "'";
									break;
								case "NOTEND":
									if (strText.Trim().Length==0)
										strFilter = strCol + " IS NULL";
									else
										strFilter =  strCol + " NOT LIKE '*" + strText.Trim() + "'";
									break;                                    
							}
						}
						break;
					default:
						if (strDataType=="System.Integer" || strDataType=="System.Double" ||
							strDataType=="System.Int16" || strDataType=="System.Int32" ||
							strDataType=="System.Int64" || strDataType=="System.Byte" ||
							strDataType=="System.Decimal")
						{
							frmDialog frmTemp2 = new frmDialog();
							frmTemp2.Initialize_Filter_Rows_Numeric_Datatype_User_Control();

							frmTemp2.Text = strCol;
							frmTemp2.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
							dlgResult = frmTemp2.ShowDialog();
							if (dlgResult.ToString().Trim().ToUpper() == "OK" )
							{
								this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=true;
								string strOperation = frmTemp2.uc_filter_rows_numeric_datatype1.GetFilterType();
								string strText = frmTemp2.uc_filter_rows_numeric_datatype1.GetSmallestText().Replace(",","");
								string strLargestText = frmTemp2.uc_filter_rows_numeric_datatype1.GetLargestText().Replace(",","");

								switch (strOperation)
								{
									case "EQUAL":
										if (strText.Trim().Length==0)
											strFilter = strCol + " IS NULL";
										else
											strFilter =  strCol + "=" + strText;
										break;
									case "NOTEQUAL":
										if (strText.Trim().Length==0)
											strFilter = strCol + " IS NOT NULL";
										else
											strFilter =  strCol + "<>" + strText;
										break;
									case "GREATERTHAN":
										if (strText.Trim().Length==0)
											strFilter = strCol + " IS NULL";
										else
											strFilter =  strCol + " > " + strText.Trim();
										break;
									case "LESSTHAN":
										if (strText.Trim().Length==0)
											strFilter = strCol + " IS NULL";
										else
											strFilter =  strCol + " < " + strText.Trim();
										break;
									case "BETWEEN":
										if (strText.Trim().Length==0)
											strFilter = strCol + " IS NULL";
										else
											strFilter =  strCol + " >= " + strText + " AND " + strCol + " <= " + strLargestText;
										break;
										              
								}
							}
						}
						break;

				}
			
					break;
			}
			if (strFilter.Trim().Length > 0)
			{
				this.m_dv.RowFilter = strFilter;

				if (this.m_strColumnFilterList.Trim().Length > 0) this.m_strColumnFilterList+=  strCol.Trim() + ",";
				else this.m_strColumnFilterList = strCol.Trim() + ",";
			}
			
					
			
		}

		public void DeleteRecords()
		{
			int x;
			int intCurrRow;
			string strValue="";
			System.Windows.Forms.CurrencyManager p_cm;
			System.Data.DataView p_dv;
			string strList="";

			this.m_intError=0;

			p_cm = (CurrencyManager)this.BindingContext[this.m_dg.DataSource,this.m_dg.DataMember];
			p_dv = (DataView)p_cm.List;
			intCurrRow= this.m_intCurrRow-1;
			//add selected plots to the delete list
			for (x=0; x < p_dv.Count;x++)
			{
				if (m_dg.IsSelected(x) || x == intCurrRow)
				{
					if (this.m_strDeletedList.Trim().Length > 0)
					{
						this.m_strDeletedList += "," +  p_dv[x]["psite_id"].ToString().Trim();
					}
					else
					{
						this.m_strDeletedList = p_dv[x]["psite_id"].ToString().Trim();
					}
					strList += "," + p_dv[x]["psite_id"].ToString().Trim() + ",";
					this.m_intDeletedCount++;
						
				}
			}
			if (this.m_intDeletedCount > 0)
			{
				string strMsg = this.m_intDeletedCount.ToString().Trim() + " record(s) will be deleted from the table. \n Do you want to delete the record(s) from the table? (Y/N)";
				DialogResult result = MessageBox.Show(strMsg,"Delete Selected Processing Sites", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
				if (result == DialogResult.Yes)
				{
					this.m_dv.AllowDelete = true;

					if (this.m_intError==0 && m_dataMgr.m_intError==0)
					{
						//delete any plots that are in the delete list
						for (x=0; x < p_dv.Count;x++)
						{
							strValue = "," + p_dv[x]["psite_id"].ToString().Trim() + ",";
							if (strList.IndexOf(strValue) >=0)
							{
								p_dv[x].Delete();
								x=x-1;
							}
						}
						if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
					}
				}
			}
			this.m_strDeletedList="";
			this.m_intDeletedCount=0;
			this.m_dv.AllowDelete=false;
			
			return;
		}


		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnHelp = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnNew = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.m_dg = new System.Windows.Forms.DataGrid();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.m_dg)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.btnDelete);
            this.groupBox1.Controls.Add(this.btnNew);
            this.groupBox1.Controls.Add(this.btnEdit);
            this.groupBox1.Controls.Add(this.btnSave);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.m_dg);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(704, 352);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(598, 312);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 58;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(12, 312);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(96, 32);
            this.btnHelp.TabIndex = 57;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Enabled = false;
            this.btnDelete.Location = new System.Drawing.Point(336, 280);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(64, 32);
            this.btnDelete.TabIndex = 56;
            this.btnDelete.Text = "Delete";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnNew
            // 
            this.btnNew.Location = new System.Drawing.Point(208, 280);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(64, 32);
            this.btnNew.TabIndex = 52;
            this.btnNew.Text = "New";
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(272, 280);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(64, 32);
            this.btnEdit.TabIndex = 53;
            this.btnEdit.Text = "Edit";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnSave
            // 
            this.btnSave.Enabled = false;
            this.btnSave.Location = new System.Drawing.Point(400, 280);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(64, 32);
            this.btnSave.TabIndex = 54;
            this.btnSave.Text = "Save";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(464, 280);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(64, 32);
            this.btnCancel.TabIndex = 55;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // m_dg
            // 
            this.m_dg.DataMember = "";
            this.m_dg.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.m_dg.Location = new System.Drawing.Point(8, 64);
            this.m_dg.Name = "m_dg";
            this.m_dg.Size = new System.Drawing.Size(688, 208);
            this.m_dg.TabIndex = 27;
            this.m_dg.CurrentCellChanged += new System.EventHandler(this.m_dg_CurrentCellChanged);
            this.m_dg.MouseDown += new System.Windows.Forms.MouseEventHandler(this.m_dg_MouseDown);
            this.m_dg.MouseUp += new System.Windows.Forms.MouseEventHandler(this.m_dg_MouseUp);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(698, 32);
            this.lblTitle.TabIndex = 26;
            this.lblTitle.Text = "Processing Sites";
            // 
            // uc_gis_psite
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_gis_psite";
            this.Size = new System.Drawing.Size(704, 352);
            this.Resize += new System.EventHandler(this.uc_gis_psite_Resize);
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.m_dg)).EndInit();
            this.ResumeLayout(false);

		}
		#endregion

		private void m_dg_CurrentCellChanged(object sender, System.EventArgs e)
		{
			/*********************************************************************
			 **this one command causes the whole row to be highlighted with 
			 **the selected row background and foreground color
			 *********************************************************************/
			if (m_dg.CurrentRowIndex >=0) m_dg.Select(m_dg.CurrentRowIndex);
			if (this.m_intCurrRow > 0)
			{
				if (this.m_dg.CurrentRowIndex != this.m_intCurrRow - 1)
				{
					this.m_intCurrRow = this.m_dg.CurrentRowIndex + 1;
				}
			}
		}

		private void m_dg_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
            switch (e.Button)
            {
                case System.Windows.Forms.MouseButtons.Right:
                    if (this.m_dv.RowFilter.Trim().Length > 0)
                    {
                        if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled == false)
                            this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled = true;
                        if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled == true)
                            this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled = false;
                        if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled == true)
                            this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled = false;

                    }
                    else
                    {
                        if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled == false)
                            this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled = true;
                        if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled == true)
                            this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled = false;
                        if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled == false)
                            this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled = true;

                    }
                    if (this.m_dv.Sort.Trim().Length > 0)
                    {
                        if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled == false)
                            this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled = true;
                        //if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled==true)
                        //	this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled=false;
                        //if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled==true)
                        //	this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled=false;

                    }
                    else
                    {
                        if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled == false)
                            this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled = true;
                        if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled == false)
                            this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled = true;
                        if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled == true)
                            this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled = false;

                    }

                    if (this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled == false)
                        this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled = true;
                    if (this.m_mnuDataGridPopup.MenuItems[MENU_SELECTALL].Enabled == false)
                        this.m_mnuDataGridPopup.MenuItems[MENU_SELECTALL].Enabled = true;
                    if (this.m_mnuDataGridPopup.MenuItems[MENU_MIN].Enabled == false)
                        this.m_mnuDataGridPopup.MenuItems[MENU_MIN].Enabled = true;
                    if (this.m_mnuDataGridPopup.MenuItems[MENU_MAX].Enabled == false)
                        this.m_mnuDataGridPopup.MenuItems[MENU_MAX].Enabled = true;
                   

                    break;
            }
            

		}
		private void uniquevalues()
		{
			string strCellValue;
			string strBuild="";
			int count=0;
			int x=0;
			
			string strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();

			System.Windows.Forms.CurrencyManager p_cm = (CurrencyManager)this.BindingContext[this.m_dg.DataSource,this.m_dg.DataMember];
			System.Data.DataView p_dv = (DataView)p_cm.List;
			string[] strValues = new string[p_dv.Count];

			for (x = 0; x<= p_dv.Count-1; x++)
			{
				strCellValue = this.m_dg[x,this.m_dg.CurrentCell.ColumnNumber].ToString();
				if (strBuild.IndexOf(strCellValue,0,strBuild.Length) == -1)
				{
					strValues[count] = strCellValue;
					
					strBuild = strBuild + "'" + strCellValue + "',";
					count++;
				}

			}																																					 
 					
			
			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
			System.Windows.Forms.ListBox p_lst = new ListBox();
         
						
			//define form properties
			frmTemp.Width = 200;
			frmTemp.Height = 300;
			frmTemp.WindowState = System.Windows.Forms.FormWindowState.Normal;
			frmTemp.MaximizeBox = false;
			frmTemp.MinimizeBox = false;
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.Text = "Unique Values (" + strCol + ")";
		                    
			//define numeric text class properties
			p_lst.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			p_lst.Name = "lstUniqueValues";
			p_lst.TabIndex = 0;
			p_lst.Tag = "";
			p_lst.Enabled = true;
			frmTemp.Controls.Add(p_lst);  //add the text box to the form
			

			p_lst.Top = 5;
			p_lst.Width = frmTemp.ClientSize.Width - 20;
			p_lst.Left = 10;
			p_lst.Items.Clear();
			int intLeft = p_lst.Left;
			int intTop = p_lst.Top;

					
			//----------------OK button-----------------//
			System.Windows.Forms.Button btnTempOK = new Button();
			btnTempOK.Location = new System.Drawing.Point(392, 328);
			btnTempOK.Name = "btnOK";
			btnTempOK.Size = new System.Drawing.Size(80, 32);
			btnTempOK.TabIndex = 2;
			btnTempOK.Text = "OK";
			btnTempOK.Click += new System.EventHandler(frmTemp.btnOK_Click);
			frmTemp.Controls.Add(btnTempOK);
			btnTempOK.Left = (int)(frmTemp.Width * .50)  - (int)(btnTempOK.Width * .50);
			p_lst.Height = frmTemp.ClientSize.Height - btnTempOK.Height -15;
			btnTempOK.Top = intTop + p_lst.Height; // + 10;
			p_lst.Visible = true;
			btnTempOK.Visible=true;
			

			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			for (x = 0; x<= count - 1; x++)
				p_lst.Items.Add(strValues[x].ToString().Trim());

			frmTemp.ShowDialog();
			strValues = null;
			frmTemp = null;
		}
		/// <summary>
		/// return the row of the containing the column definitions of 
		/// column name, datatype, length, default value, allow nulls, etc.
		/// </summary>
		/// <param name="strColumnName">column definition to look up</param>
		/// <returns></returns>
		private int getTableSchemaColumnDefinition(string strColumnName)
		{
			int x;
			for (x=0;x<=this.m_dtTableSchema.Rows.Count-1;x++)
			{
				if (this.m_dtTableSchema.Rows[x]["COLUMNNAME"].ToString().Trim().ToUpper()==
					strColumnName.Trim().ToUpper())
				{
					return x;
					
				}

			}
			//could not find the column
			return -1;
		}

		private void m_dg_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			switch (e.Button)
			{
				case System.Windows.Forms.MouseButtons.Right :
					Point pt = new Point(e.X,e.Y);
					try
					{
					
                     
						System.Windows.Forms.DataGrid.HitTestInfo hitTestInfo = this.m_dg.HitTest(pt);
						this.m_intPopupColumn = hitTestInfo.Column;
						switch (this.m_dg[0,this.m_intPopupColumn].GetType().FullName.ToString().Trim())
						{
							case "System.String":
								this.m_mnuDataGridPopup.MenuItems[MENU_AVG].Enabled=false;
								this.m_mnuDataGridPopup.MenuItems[MENU_SUM].Enabled=false;
								break;
							default:
								this.m_mnuDataGridPopup.MenuItems[MENU_AVG].Enabled=true;
								this.m_mnuDataGridPopup.MenuItems[MENU_SUM].Enabled=true;
								break;
						}

						switch (hitTestInfo.Type)
						{
							case DataGrid.HitTestType.ColumnHeader:
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=false;
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=false;
								this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=false;
								this.m_mnuDataGridPopup.MenuItems[MENU_COUNTBYVALUE].Enabled=false;

								this.m_mnuDataGridPopup.Show(this, new Point(pt.X,pt.Y));
								break;
							case DataGrid.HitTestType.Cell:
							
								this.m_dg.CurrentCell = new DataGridCell(hitTestInfo.Row,hitTestInfo.Column);
								if (this.m_dg.TableStyles[0].GridColumnStyles[this.m_dg.CurrentCell.ColumnNumber].ReadOnly==true)
								{
									this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=false;
								}
								else
								{
									this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=true;
								}

								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=true;
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;
								this.m_mnuDataGridPopup.MenuItems[MENU_COUNTBYVALUE].Enabled=true;
								this.m_mnuDataGridPopup.Show(this, new Point(pt.X,pt.Y + this.m_dg.Top));
								break;
					  
						}

						if (this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled=true;


						if (this.m_dv.RowFilter.Trim().Length > 0)
						{
							if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=true;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled==true)
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=false;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==true)
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=false;

						}
						else
						{
							if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=true;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==true)
								this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;

						}
						if (this.m_dv.Sort.Trim().Length > 0)
						{
							if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=true;
						
						}
						else
						{
							if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled=true;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled=true;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled==true)
								this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=false;

						}
                        
					
					}
					catch
					{
						this.m_mnuDataGridPopup.MenuItems[MENU_COUNTBYVALUE].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_AVG].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_SUM].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_MAX].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_MIN].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_SELECTALL].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_DELETE].Enabled=false;

						this.m_mnuDataGridPopup.Show(this, new Point(pt.X,pt.Y + this.m_dg.Top));
					}

					break;
				case System.Windows.Forms.MouseButtons.Left :
					Point pt2 = new Point(e.X,e.Y);
					try
					{
					
						// DataGrid.HitTestInfo hitTestInfo ;
                     
						System.Windows.Forms.DataGrid.HitTestInfo hitTestInfo = this.m_dg.HitTest(pt2);
						this.m_intPopupColumn = hitTestInfo.Column;

						switch (hitTestInfo.Type)
						{
							case DataGrid.HitTestType.ColumnHeader:
								this.m_strColumnSortList = 
									this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim() + ",";
								break;
							case DataGrid.HitTestType.Cell:
								break;
						}

					
					}
					catch
					{
					}

					break;

			}
		}
		/// <summary>
		/// configure the data adapter to insert, update, and delete data using oledb sql
		/// </summary>
		private void InitializeTransactionCommands()
		{
			m_dataMgr.m_strSQL = "select * from " + this.m_strPSiteTable + " order by psite_id;";
            //initialize the transaction object with the connection
            m_dataMgr.m_Transaction = m_dataMgr.m_Connection.BeginTransaction();

            m_dataMgr.ConfigureDataAdapterInsertCommand(m_dataMgr.m_Connection,
                m_dataMgr.m_DataAdapter,
                m_dataMgr.m_Transaction,
                m_dataMgr.m_strSQL,
				this.m_strPSiteTable);

            m_dataMgr.m_strSQL = "select " + this.m_strFieldList  +  " from " + this.m_strPSiteTable + " order by name;";
            m_dataMgr.ConfigureDataAdapterUpdateCommand(m_dataMgr.m_Connection,
                m_dataMgr.m_DataAdapter,
                m_dataMgr.m_Transaction,
                m_dataMgr.m_strSQL,
                "select psite_id from " + this.m_strPSiteTable,
				this.m_strPSiteTable);

            m_dataMgr.m_strSQL = "select " + this.m_strFieldList  +  " from " + this.m_strPSiteTable + " order by name;";
			m_dataMgr.ConfigureDataAdapterDeleteCommand(m_dataMgr.m_Connection,
                m_dataMgr.m_DataAdapter,
                m_dataMgr.m_Transaction,
                "select psite_id from " + this.m_strPSiteTable,
				this.m_strPSiteTable);
		}
		
		private void m_dg_TextBox_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			switch (e.Button)
			{
				case System.Windows.Forms.MouseButtons.Right :
					this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=true;
					this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;
					this.m_mnuDataGridPopup.MenuItems[MENU_COUNTBYVALUE].Enabled=true;
					Point pt = new Point(e.X,e.Y);
					this.m_intPopupColumn=this.m_dg.CurrentCell.ColumnNumber;
					if (this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].ReadOnly==true)
					{
						this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=false;
					}
				switch (this.m_dg[0,this.m_intPopupColumn].GetType().FullName.ToString().Trim())
				{
					case "System.String":
						this.m_mnuDataGridPopup.MenuItems[MENU_AVG].Enabled=false;
						this.m_mnuDataGridPopup.MenuItems[MENU_SUM].Enabled=false;

						break;
					default:
						this.m_mnuDataGridPopup.MenuItems[MENU_AVG].Enabled=true;
						this.m_mnuDataGridPopup.MenuItems[MENU_SUM].Enabled=true;
						break;
				}
					if (this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled=true;

					if (this.m_dv.RowFilter.Trim().Length > 0)
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=false;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=false;

					}
					else
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYVALUE].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;

					}
					if (this.m_dv.Sort.Trim().Length > 0)
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=true;
						
					}
					else
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_IDXASC].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_IDXDESC].Enabled=true;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEIDX].Enabled=false;

					}



					break;
			}

			
																	   
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.CleanUp();
            ((frmDialog)ParentForm).ParentControl.Enabled = true;
			this.ParentForm.Close();
		}
		private void CleanUp()
		{
			try
			{

				if (m_dataMgr.m_DataAdapter != null)
                    m_dataMgr.m_DataAdapter.Dispose();

				if (this.m_dv != null)
					this.m_dv.Dispose();

				if (this.m_dg != null)
					this.m_dg.Dispose();

                m_dataMgr.m_DataSet.Clear();
                m_dataMgr.m_DataSet.Dispose();

                if (m_dataMgr.m_DataAdapter != null)
                {
                    m_dataMgr.m_DataAdapter.Dispose();
                }

                if (m_dataMgr.m_Command != null)
                {
                    m_dataMgr.m_Command.Dispose();
                }

                if (m_dataMgr.m_Connection != null)
				{
                    m_dataMgr.m_Connection.Close();
                    m_dataMgr.m_Connection.Dispose();
                    m_dataMgr.m_Connection = null;
				}

                GC.Collect();
                GC.WaitForPendingFinalizers();

                this.m_dataMgr = null;
				this.m_DataSource = null;

			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message);
			}

		}
		private void vScrollBar_ValueChanged(Object sender, EventArgs e)
		{
			int myValue = ((VScrollBar)sender).Value;
			this.RepositionControls();
			this.m_intScrollValue = myValue;


			

		}
		private void frmTemp_MouseWheel(Object sender, MouseEventArgs e)
		{
			
			if (e.Delta == 120)
			{
				if (this.m_vScrollBar.Value > this.m_vScrollBar.Minimum) 
				{
					if (this.m_vScrollBar.Value - this.m_vScrollBar.LargeChange < this.m_vScrollBar.Minimum) 
					{
						this.m_vScrollBar.Value = this.m_vScrollBar.Minimum;
					}
					else 
					{
						this.m_vScrollBar.Value = this.m_vScrollBar.Value - 
							this.m_vScrollBar.LargeChange;
					}
				}
			}
			else 
			{
				if (this.m_vScrollBar.Value < this.m_vScrollBar.Maximum) 
				{
					if (this.m_vScrollBar.Value + this.m_vScrollBar.LargeChange > 
						this.m_vScrollBar.Maximum) 
					{
						this.m_vScrollBar.Value = this.m_vScrollBar.Maximum;
					}
					else 
					{
						this.m_vScrollBar.Value = this.m_vScrollBar.Value +
							this.m_vScrollBar.LargeChange;
					}
				}

			}
			
		}
		
		public void RepositionControls()
		{
            int intScroll;
			double dblVal = Convert.ToDouble(this.m_vScrollBar.Value);
			double dblMax = Convert.ToDouble(this.m_vScrollBar.Maximum);

			if (this.m_vScrollBar.Value == 0)
			{
			    this.m_dblNewPerc = 0;
			}
			else
			{
				this.m_dblNewPerc = (dblVal / dblMax);
			}

			if (this.m_dblNewPerc == 0 && this.m_dblOldPerc  > 0)
			{
				intScroll =  (int)(this.m_intMaxSize * this.m_dblOldPerc);

			}
			else if (this.m_dblNewPerc > 0 && this.m_dblOldPerc == 0)
			{
				intScroll = (int)(this.m_intMaxSize * this.m_dblNewPerc);
				intScroll = -1 * intScroll;
			}
			else if (this.m_dblNewPerc > 0 && this.m_dblOldPerc > 0)
			{
				
				intScroll = (int)(this.m_intMaxSize * this.m_dblNewPerc) - (int)(this.m_intMaxSize * this.m_dblOldPerc);
				intScroll = -1 * intScroll;
			
			}
			else
			{
				intScroll = 0;
			}
			for (int z=0;z<=this.m_frmDialog.Controls.Count-1;z++)
			{
				this.m_frmDialog.Controls[z].Top = this.m_frmDialog.Controls[z].Top + intScroll;
			}
			this.m_dblOldPerc = this.m_dblNewPerc;
		}
		

		
		/// <summary>
		/// get the unique id for a processing_site record
		/// </summary>
		/// <returns>the unique id</returns>
		private int getUniqueId()
		{
			string strUniqueId="";
			int intId=0;
			int intId2=0;
            strUniqueId = m_dataMgr.getSingleStringValueFromSQLQuery(m_dataMgr.m_Connection,
                m_dataMgr.m_Transaction, "select max(psite_id) as maxid from " + this.m_strPSiteTable, this.m_strPSiteTable);

			if (strUniqueId != null && strUniqueId.Trim().Length > 0)
				intId = Convert.ToInt32(strUniqueId);

			if (m_dataMgr.m_DataSet.Tables["processing_site"].Compute("Max(psite_id)", null) != System.DBNull.Value)
			{
				intId2 = Convert.ToInt32(m_dataMgr.m_DataSet.Tables["processing_site"].Compute("Max(psite_id)", null));
			}
			if (intId2 >= intId) 
			{
				intId=intId2 + 1;
			}

			if (intId==0) return 1;
			return intId;
		}

		private void btnNew_Click(object sender, System.EventArgs e)
		{
			
			int x=5;
			int y=0;
			int z=0;
			string strField;
			int intLargestFieldSize=200;
            int intFieldSize=0;
			/******************************************************************
			 **instatiate a new dialog form using template objects
			 **for title, labels, combo menus, and text boxes
			 ******************************************************************/
			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
			frmTemp.Text = "GIS: Processing Site (New)";
			frmTemp.MdiParent = ((frmMain)this.ParentForm.ParentForm);
			
			frmTemp.AutoScroll = false;
			this.m_vScrollBar = new VScrollBar();
			this.m_vScrollBar.Dock = System.Windows.Forms.DockStyle.Right;
			this.m_vScrollBar.Visible=false;
			this.m_vScrollBar.Minimum=0;
			this.m_vScrollBar.Maximum=this.Height;
			this.m_vScrollBar.Value=0;
			this.m_vScrollBar.LargeChange = 20;
			this.m_vScrollBar.SmallChange = 10;

			this.m_vScrollBar.ValueChanged += new EventHandler(
				vScrollBar_ValueChanged);
			frmTemp.MouseWheel += new MouseEventHandler(frmTemp_MouseWheel);	
			frmTemp.Controls.Add(this.m_vScrollBar);
			this.m_vScrollBar.Visible=true;
			
			
			FIA_Biosum_Manager.TemplateGroupBox p_groupBox1 = new TemplateGroupBox(frmTemp);
			FIA_Biosum_Manager.TemplateTitle p_lblTitle = new TemplateTitle(p_groupBox1,0,0,"Wood Processing Site (New)");
			FIA_Biosum_Manager.TemplateOkCancelButtons p_OkCancel = new TemplateOkCancelButtons(frmTemp,p_groupBox1);
			p_groupBox1.Dock = System.Windows.Forms.DockStyle.None;
		    FIA_Biosum_Manager.TemplateInputLabel p_lbl = new TemplateInputLabel(p_groupBox1,"temp","temp");
			FIA_Biosum_Manager.TemplateTextBox p_txt = new TemplateTextBox(p_groupBox1,"txttemp","");
			FIA_Biosum_Manager.TemplateComboBox p_cmb = new TemplateComboBox(p_groupBox1,"cmbtemp","");

			p_lbl.Visible=false;
			p_txt.Visible=false;
			p_cmb.Visible=false;

			//y controls the placement of the top of each object on the form
			y = p_lblTitle.Top + p_lblTitle.Height + 5;
			/*******************************************************************
			 **using the table schema information define the labels and 
			 **input that will be placed on the form
			 *******************************************************************/
			for (z=0; z<=this.m_dtTableSchema.Rows.Count-1;z++)
			{
				strField=this.m_dtTableSchema.Rows[z]["columnname"].ToString().Trim();
				switch (strField.ToUpper())
				{
					case "PSITE_ID":
						p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),strField);
						p_lbl.Left = x;
						p_lbl.Top = y;
						p_txt = new TemplateTextBox(p_groupBox1, strField, Convert.ToString(this.getUniqueId()));
						p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
						p_txt.Top = p_lbl.Top;
						p_txt.Width = 200;
						p_txt.Visible=true;
						p_txt.ReadOnly=true;
						y += p_lbl.Height;
						p_lbl.Visible=true;
						break;
					case "TRANCD":
						break;
					case "BIOCD":
						break;
					case "TRANCD_DEF":
						p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),"Transportation Code");
						p_lbl.Left = x;
						p_lbl.Top = y;
						p_cmb = new TemplateComboBox(p_groupBox1,strField.Trim(),"");
						p_cmb.Items.Add("Regular - PSite With Road Only Access");
						p_cmb.Items.Add("Railhead - Road To Rail Wood Transfer Point");
						p_cmb.Items.Add("Rail Collector - PSite With Both Road And Rail Access");
						p_cmb.Left = p_lbl.Left + p_lbl.Width + 5;
						p_cmb.Top = p_lbl.Top;
					    intFieldSize = (int)this.CreateGraphics().MeasureString("Rail Collector - PSite With Both Road And Rail Access", p_cmb.Font).Width + 50;
						p_cmb.Width = intFieldSize;
						p_cmb.Visible=true;
						//p_cmb.ReadOnly=true;
						y += p_lbl.Height;
						p_lbl.Visible=true;
						break;
					case "BIOCD_DEF":
						p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),"Processing Abilities");
						p_lbl.Left = x;
						p_lbl.Top = y;
						p_cmb = new TemplateComboBox(p_groupBox1,strField.Trim(),"");
						p_cmb.Items.Add("Merchantable - Logs Only");
						p_cmb.Items.Add("Chip - Chips Only");
						p_cmb.Items.Add("Both - Logs And Chips");
						p_cmb.Left = p_lbl.Left + p_lbl.Width + 5;
						p_cmb.Top = p_lbl.Top;
						intFieldSize = (int)this.CreateGraphics().MeasureString("Merchantable - Logs Only", p_cmb.Font).Width + 50;
						p_cmb.Width = intFieldSize;
						p_cmb.Visible=true;
						//p_cmb.ReadOnly=true;
						y += p_lbl.Height;
						p_lbl.Visible=true;
						break;
					case "EXISTS_YN":
						p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),"Processing Site Exist");
						p_lbl.Left = x;
						p_lbl.Top = y;
						p_cmb = new TemplateComboBox(p_groupBox1,strField.Trim(),"");
						p_cmb.Items.Add("Yes");
						p_cmb.Items.Add("No");
                        p_cmb.Text = "No";
						p_cmb.Left = p_lbl.Left + p_lbl.Width + 5;
						p_cmb.Top = p_lbl.Top;
						intFieldSize = (int)this.CreateGraphics().MeasureString("Yes", p_cmb.Font).Width + 50;
						p_cmb.Width = intFieldSize;
						p_cmb.Visible=true;
						//p_cmb.ReadOnly=true;
						y += p_lbl.Height;
						p_lbl.Visible=true;
						break;
                    case "KEEPLYR_YN":
                        // GitHub issue #56: Remove Keep Temp GIS Layer comboBox; Do nothing for now
                    //    p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),"Keep Temp GIS Layer");
                    //    p_lbl.Left = x;
                    //    p_lbl.Top = y;
                    //    p_cmb = new TemplateComboBox(p_groupBox1,strField.Trim(),"");
                    //    p_cmb.Items.Add("Yes");
                    //    p_cmb.Items.Add("No");
                    //    p_cmb.Text = "No";
                    //    p_cmb.Left = p_lbl.Left + p_lbl.Width + 5;
                    //    p_cmb.Top = p_lbl.Top;
                    //    intFieldSize = (int)this.CreateGraphics().MeasureString("Yes", p_cmb.Font).Width + 50;
                    //    p_cmb.Width = intFieldSize;
                    //    p_cmb.Visible=true;
                    //    //p_cmb.ReadOnly=true;
                    //    y += p_lbl.Height;
                    //    p_lbl.Visible=true;
                        break;
					case "NAME":
						p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),"Processing Site Name");
						p_lbl.Left = x;
						p_lbl.Top = y;
						p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), "");
						p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
						p_txt.Top = p_lbl.Top;
						p_txt.Width = 400;
						p_txt.MaxLength = Convert.ToInt32(this.m_dtTableSchema.Rows[z]["columnsize"]);
						p_txt.Visible=true;
						y += p_lbl.Height;
						p_lbl.Visible=true;
						break;
					default:
						p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),strField);
						p_lbl.Left = x;
						p_lbl.Top = y;
						switch (this.m_dtTableSchema.Rows[z]["datatype"].ToString().Trim())
						{
							case "System.String" :
								p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), "");
								p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
								p_txt.Top = p_lbl.Top;
								p_txt.Width = 200;
								p_txt.Visible=true;
								y += p_lbl.Height;
								p_lbl.Visible=true;
								p_txt.MaxLength = Convert.ToInt32(this.m_dtTableSchema.Rows[z]["columnsize"]);
							
								break;
							case "System.Double":
								p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), "");
								p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
								p_txt.Top = p_lbl.Top;
								p_txt.Width = 200;
								p_txt.Visible=true;
								y += p_lbl.Height;
								p_lbl.Visible=true;
								break;
							case "System.Boolean":
								p_cmb = new TemplateComboBox(p_groupBox1,strField.Trim(),"");
								p_cmb.Items.Add("True");
								p_cmb.Items.Add("False");
								p_cmb.Text = "False";
								p_cmb.Left = p_lbl.Left + p_lbl.Width + 5;
								p_cmb.Top = p_lbl.Top;
								intFieldSize = (int)this.CreateGraphics().MeasureString("Yes", p_cmb.Font).Width + 50;
								p_cmb.Width = intFieldSize;
								p_cmb.Visible=true;
								y += p_lbl.Height;
								break;
							case "System.DateTime":
								p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), "");
								p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
								p_txt.Top = p_lbl.Top;
								p_txt.Width = 200;
								p_txt.Visible=true;
								y += p_lbl.Height;
								p_lbl.Visible=true;
							
								break;
							case "System.Decimal":
								p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), "");
								p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
								p_txt.Top = p_lbl.Top;
								p_txt.Width = 200;
								p_txt.Visible=true;
								y += p_lbl.Height;
								p_lbl.Visible=true;						
								break;
							case "System.Int16":
								p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), "");
								p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
								p_txt.Top = p_lbl.Top;
								p_txt.Width = 200;
								p_txt.Visible=true;
								y += p_lbl.Height;
								p_lbl.Visible=true;						
								break;
							case "System.Int32":
								p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), "");
								p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
								p_txt.Top = p_lbl.Top;
								p_txt.Width = 200;
								p_txt.Visible=true;
								y += p_lbl.Height;
								p_lbl.Visible=true;						
								break;
							case "System.Int64":
								p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), "");
								p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
								p_txt.Top = p_lbl.Top;
								p_txt.Width = 200;
								p_txt.Visible=true;
								y += p_lbl.Height;
								p_lbl.Visible=true;						
								break;
							case "System.SByte":
								p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), "");
								p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
								p_txt.Top = p_lbl.Top;
								p_txt.Width = 100;
								p_txt.Visible=true;
								p_txt.MaxLength = 1;
								y += p_lbl.Height;
								p_lbl.Visible=true;
								
								break;
							case "System.Byte":
								p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), "");
								p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
								p_txt.Top = p_lbl.Top;
								p_txt.Width = 100;
								p_txt.Visible=true;
								p_txt.MaxLength = 1;
								y += p_lbl.Height;
								p_lbl.Visible=true;
						
								break;
							case "System.Single":
								p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), "");
								p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
								p_txt.Top = p_lbl.Top;
								p_txt.Width = 100;
								p_txt.Visible=true;
								y += p_lbl.Height;
								p_lbl.Visible=true;						
								break;
							default:
								MessageBox.Show("Datatype Not Found:" + this.m_dtTableSchema.Rows[z]["datatype"].ToString().Trim());
								break;
						}
						break;
				
					
				}
				if (intFieldSize > intLargestFieldSize)
				{
					intLargestFieldSize = intFieldSize;
				}
				
			}
			
			//size the groupbox to encompass all the objects
			p_groupBox1.Width  = p_txt.Left + intLargestFieldSize + 20;
			if (p_groupBox1.Left + p_groupBox1.Width < p_txt.Left + intLargestFieldSize + 20)
			{
				for (x=1;;x++)
				{
					p_groupBox1.Width = x;
					if (p_groupBox1.Width  < 
						p_txt.Left + intLargestFieldSize)
					{
						break;
					}
				}

			}

			

		
			p_OkCancel.btnOK.Top = p_lbl.Top + p_lbl.Height + 5;
			p_OkCancel.btnCancel.Top = p_OkCancel.btnOK.Top;
			p_OkCancel.btnOK.Left = (int)(p_groupBox1.Width * .50) - p_OkCancel.btnOK.Width;
			p_OkCancel.btnCancel.Left = p_OkCancel.btnOK.Left + p_OkCancel.btnOK.Width;
			p_groupBox1.Height = p_OkCancel.btnOK.Top + p_OkCancel.btnOK.Height + 10;

			//groupbox height
			if (p_groupBox1.Top + p_groupBox1.Height < p_OkCancel.btnOK.Top + p_OkCancel.btnOK.Height + 15)
			{
				for (x=1;;x++)
				{
					p_groupBox1.Height = x;
					if (p_groupBox1.Height + 5 >
						p_OkCancel.btnOK.Top + (p_OkCancel.btnOK.Height * 1.5))
					{
						break;
					}
				}

			}

             
			frmTemp.Height=0;
			frmTemp.Width=0;
			if (p_groupBox1.Top + p_groupBox1.Height > frmTemp.ClientSize.Height + 2)
			{
				for (x=1;;x++)
				{
					frmTemp.Height = x;
					if (p_groupBox1.Top + 
						p_groupBox1.Height + 20 < 
						frmTemp.ClientSize.Height)
					{
						break;
					}
				}

			}
			if (p_groupBox1.Left + p_groupBox1.Width > frmTemp.ClientSize.Width + 20)
			{
				for (x=1;;x++)
				{
					frmTemp.Width = x;
					if (p_groupBox1.Left + 
						p_groupBox1.Width + this.m_vScrollBar.Width + 5 < 
						frmTemp.ClientSize.Width)
					{
						break;
					}
				}

			}
		
			frmTemp.DisposeOfFormWhenClosing = false;
			p_groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			p_groupBox1.Dock = System.Windows.Forms.DockStyle.None;
			frmTemp.Left = 0;
			frmTemp.Top = 0;
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            frmTemp.MinimizeBox = false;
            frmTemp.MaximizeBox = false;

 	        this.m_vScrollBar.Maximum = p_groupBox1.Height; //p_OkCancel.btnOK.Top + (int)(p_OkCancel.btnOK.Height * 1.5);
            this.m_intMaxSize = p_groupBox1.Height;
            this.m_frmDialog = frmTemp;
			//loop until record passes validation or user cancels
			for (;;)
			{
				System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
				if (result==System.Windows.Forms.DialogResult.OK)
				{
					this.m_intError=0;
                    
					this.val_data(p_groupBox1);
					if (this.m_intError==0)
					{
						this.m_dv.AllowNew = true;
						System.Data.DataRow p_row =	this.m_ado.m_DataSet.Tables["processing_site"].NewRow();
						for (z=0;z<=p_groupBox1.Controls.Count-1;z++)
						{
							switch (p_groupBox1.Controls[z].Name.Trim().ToUpper())
							{
								case "PSITE_ID":
									p_row["psite_id"] = Convert.ToInt32(p_groupBox1.Controls[z].Text);
									break;
								case "NAME":
									p_row["name"] = p_groupBox1.Controls[z].Text;
									break;
                                // GitHub issue #56: Remove Keep Temp GIS Layer comboBox
                                //case "KEEPLYR_YN":
                                //    p_row["keeplyr_yn"] = p_groupBox1.Controls[z].Text.Substring(0,1);
                                //    break;
								case "EXISTS_YN":
									p_row["exists_yn"] = p_groupBox1.Controls[z].Text.Substring(0,1);
									break;
								case "BIOCD_DEF":
									if (p_groupBox1.Controls[z].Text.Trim().Length > 0)
									{
										if (p_groupBox1.Controls[z].Text.IndexOf("Merchantable -",0) >=0)
										{
											p_row["biocd"] = 1;
											p_row["biocd_def"] = "Merchantable";
										}
										else if (p_groupBox1.Controls[z].Text.IndexOf("Chip -",0) >=0)
										{
											p_row["biocd"] = 2;
											p_row["biocd_def"] = "Chip";
										}
										else if (p_groupBox1.Controls[z].Text.IndexOf("Both -",0) >=0)
										{
											p_row["biocd"] = 3;
											p_row["biocd_def"] = "Both";
										}
									}
									break;
								case "TRANCD_DEF":
									if (p_groupBox1.Controls[z].Text.Trim().Length > 0)
									{
										if (p_groupBox1.Controls[z].Text.IndexOf("Regular",0) >=0)
										{
											p_row["trancd"] = 1;
											p_row["trancd_def"] = "Regular";
										}
										else if (p_groupBox1.Controls[z].Text.IndexOf("Railhead",0) >=0)
										{
											p_row["trancd"] = 2;
											p_row["trancd_def"] = "Railhead";
										}
										else if (p_groupBox1.Controls[z].Text.IndexOf("Collector",0) >=0)
										{
											p_row["trancd"] = 3;
											p_row["trancd_def"] = "Both";
										}
									}
									break;
								default:
									for (x=0;x<=this.m_dtTableSchema.Rows.Count-1;x++)
									{
										strField=this.m_dtTableSchema.Rows[x]["columnname"].ToString().Trim();
										if (strField.ToUpper() == p_groupBox1.Controls[z].Name.Trim().ToUpper())
										{
											if (p_groupBox1.Controls[z].Text == null || p_groupBox1.Controls[z].Text.Trim().Length == 0)
											{
												p_row[strField] = System.DBNull.Value;
											}
											else
											{
												switch (this.m_dtTableSchema.Rows[x]["datatype"].ToString().Trim())
												{
													case "System.String" :
														p_row[strField] = Convert.ToString(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Double":
														p_row[strField] = Convert.ToDouble(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Boolean":
														p_row[strField] = Convert.ToBoolean(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.DateTime":
														p_row[strField] = Convert.ToDateTime(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Decimal":
														p_row[strField] = Convert.ToDecimal(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Int16":
														p_row[strField] = Convert.ToInt16(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Int32":
														p_row[strField] = Convert.ToInt32(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Int64":
														p_row[strField] = Convert.ToInt64(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.SByte":
														p_row[strField] = Convert.ToSByte(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Byte":
														p_row[strField] = Convert.ToByte(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Single":
														p_row[strField] = Convert.ToSingle(p_groupBox1.Controls[z].Text.Trim());
														break;
													default:
														MessageBox.Show("Datatype Not Found:" + this.m_dtTableSchema.Rows[z]["datatype"].ToString().Trim());
														break;
												}
											}

										}

									}
									break;
							}
						}

						m_dataMgr.m_DataSet.Tables["processing_site"].Rows.Add(p_row);
						p_row=null;
						this.m_dv.AllowNew = false;
						if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
						break;
					}
				}
				else break;
					 
			}
			frmTemp.Close();
			frmTemp.Dispose();
			this.m_intMaxSize=0;

			this.m_dblNewPerc=0;
			this.m_dblOldPerc=0;
			this.m_intScrollValue=0;
			frmTemp=null;
			this.m_frmDialog=null;

		}
		/// <summary>
		/// validate the input objects located on the p_groupBox object
		/// </summary>
		/// <param name="p_groupBox">groupbox control that owns the input objects</param>
		private void val_data(System.Windows.Forms.GroupBox p_groupBox)
		{
			int x;
			int z=0;

			try
			{
				for (z=0;z<=p_groupBox.Controls.Count-1;z++)
				{
					/**************************************************************
					 **the name of the control corresponds to the name of the 
					 **column on the grid and the tables field name
					 **************************************************************/
					switch (p_groupBox.Controls[z].Name.Trim().ToUpper())
					{
						case "NAME":
							//error for not entering the name of the processing site facility
							if (p_groupBox.Controls[z].Text.Trim().Length==0)
							{
								MessageBox.Show("!!Enter A Name For The Processing Site!!","FIA Biosum",
									System.Windows.Forms.MessageBoxButtons.OK,
									System.Windows.Forms.MessageBoxIcon.Exclamation);
								p_groupBox.Controls[z].Focus();
								m_intError=-1;
								return;
							}
							break;
                        //GitHub issue #56: Remove Keep Temp GIS Layer comboBox
                        //case "KEEPLYR_YN":
                        //    /****************************************************
                        //     **if user did not select a value then use 'No' as
                        //     **the default
                        //     ****************************************************/
                        //    if (p_groupBox.Controls[z].Text.Trim().ToUpper()!="YES" && 
                        //        p_groupBox.Controls[z].Text.Trim().ToUpper()!="NO")
                        //    {
                        //        p_groupBox.Controls[z].Text = "No";
                        //    }
                        //    break;
						case "EXISTS_YN":
							//error if user did not specify whether the processing site already exists
							if (p_groupBox.Controls[z].Text.Trim().ToUpper()!="YES" &&
								p_groupBox.Controls[z].Text.Trim().ToUpper()!="NO")
							{
								MessageBox.Show("!!Enter Whether The Processing Site Currently Exists!!","FIA Biosum",
									System.Windows.Forms.MessageBoxButtons.OK,
									System.Windows.Forms.MessageBoxIcon.Exclamation);
								p_groupBox.Controls[z].Focus();
								m_intError=-1;
								return;
							}
							break;
						case "BIOCD_DEF":
						     /*******************************************************************
							  **error if user did not specify the materials that
							  **the processing site can process
							  *******************************************************************/
							if (p_groupBox.Controls[z].Text.IndexOf("Merchantable",0) >=0 ||
								p_groupBox.Controls[z].Text.IndexOf("Chip",0) >=0 ||
								p_groupBox.Controls[z].Text.IndexOf("Both",0) >=0)
							{
							}
							else
							{
								MessageBox.Show("!!Enter The Wood Product Material(s) The Processing Site Is Able To Process!!","FIA Biosum",
									System.Windows.Forms.MessageBoxButtons.OK,
									System.Windows.Forms.MessageBoxIcon.Exclamation);
								p_groupBox.Controls[z].Focus();
								m_intError=-1;
								return;
							}
							break;
						case "TRANCD_DEF":
							/*************************************************************
							 **error if user did not specify whether the processing
							 **site can be access by road,rail, or both
							 *************************************************************/
							if (p_groupBox.Controls[z].Text.IndexOf("Regular",0) >=0 ||
								p_groupBox.Controls[z].Text.IndexOf("Railhead",0) >=0 ||
								p_groupBox.Controls[z].Text.IndexOf("Collector",0) >=0)
							{
							}
							else
							{
								MessageBox.Show("!!Enter The Transportation Access To The Processing Site!!","FIA Biosum",
									System.Windows.Forms.MessageBoxButtons.OK,
									System.Windows.Forms.MessageBoxIcon.Exclamation);
								p_groupBox.Controls[z].Focus();
								m_intError=-1;
								return;
							}
							break;

						default:
							string str;
                            double dbl;
							int intVar;
							//bool b;
							decimal dec;
							sbyte sb;
							byte byteVar;
							long l;
							string strField;
							
							//fields added by the user
							/*******************************************************
							 **check the input by converting the user input
							 **to the columns datatype.  An error will occur
							 **if the conversion to the datatype fails
							 *******************************************************/
							for (x=0;x<=this.m_dtTableSchema.Rows.Count-1;x++)
							{
								strField=this.m_dtTableSchema.Rows[x]["columnname"].ToString().Trim();
								if (strField.ToUpper() == p_groupBox.Controls[z].Name.Trim().ToUpper())
								{
									//check if the input value is null or length of zero
									if (p_groupBox.Controls[z].Text == null || p_groupBox.Controls[z].Text.Trim().Length == 0)
									{
									}
									else
									{
										//test to make sure it can convert to the fields datatype
										switch (this.m_dtTableSchema.Rows[x]["datatype"].ToString().Trim())
										{
											case "System.String" :
												str = Convert.ToString(p_groupBox.Controls[z].Text.Trim());
												break;
											case "System.Double":
												dbl = Convert.ToDouble(p_groupBox.Controls[z].Text.Trim());
												break;
											case "System.Boolean":
												if (p_groupBox.Controls[z].Text.Trim().ToUpper()!="TRUE" &&
													p_groupBox.Controls[z].Text.Trim().ToUpper()!="FALSE")
												{
													MessageBox.Show("!!Enter A Value For " + p_groupBox.Controls[z].Name + "!!","FIA Biosum",
														System.Windows.Forms.MessageBoxButtons.OK,
														System.Windows.Forms.MessageBoxIcon.Exclamation);
													p_groupBox.Controls[z].Focus();
													m_intError=-1;
													return;
												}												
												break;
											case "System.DateTime":
												Convert.ToDateTime(p_groupBox.Controls[z].Text.Trim());
												break;
											case "System.Decimal":
												dec = Convert.ToDecimal(p_groupBox.Controls[z].Text.Trim());
												break;
											case "System.Int16":
												intVar = Convert.ToInt16(p_groupBox.Controls[z].Text.Trim());
												break;
											case "System.Int32":
												intVar = Convert.ToInt32(p_groupBox.Controls[z].Text.Trim());
												break;
											case "System.Int64":
												l = Convert.ToInt64(p_groupBox.Controls[z].Text.Trim());
												break;
											case "System.SByte":
											    sb = Convert.ToSByte(p_groupBox.Controls[z].Text.Trim());
												break;
											case "System.Byte":
												byteVar = Convert.ToByte(p_groupBox.Controls[z].Text.Trim());
												break;
											case "System.Single":
											    dbl = Convert.ToSingle(p_groupBox.Controls[z].Text.Trim());
												break;
											default:
												MessageBox.Show("Datatype Not Found:" + this.m_dtTableSchema.Rows[z]["datatype"].ToString().Trim());
												break;
										}
									}

								}

							}
							break;
					}
					
				}
			
			}
			//display error message if datatype conversion fails
			catch 
			{
				MessageBox.Show("!!Invalid Value For " + p_groupBox.Controls[z].Name + "!!","FIA Biosum",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				p_groupBox.Controls[z].Focus();
				m_intError=-1;
			}
		}

		private void btnEdit_Click(object sender, System.EventArgs e)
		{
            if (m_dv.Count == 0) return;
			int x=5;
			int y=0;
			int z=0;
			//int w=0;
			string strField;
			int intLargestFieldSize=200;
			int intFieldSize=0;
			string strValue="";
			int intGridCol = 0;           //grid column
			
			/******************************************************************
			 **instatiate a new dialog form using template objects
			 **for title, labels, combo menus, and text boxes
			 ******************************************************************/
			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
			frmTemp.Text = "GIS: Processing Site (Edit)";

			frmTemp.AutoScroll = false;
			this.m_vScrollBar = new VScrollBar();
			this.m_vScrollBar.Dock = System.Windows.Forms.DockStyle.Right;
			this.m_vScrollBar.Visible=false;
			this.m_vScrollBar.Minimum=0;
			this.m_vScrollBar.Maximum=this.Height;
			this.m_vScrollBar.Value=0;
			this.m_vScrollBar.LargeChange = 20;
			this.m_vScrollBar.SmallChange = 10;

			this.m_vScrollBar.ValueChanged += new EventHandler(
				vScrollBar_ValueChanged);
			frmTemp.MouseWheel += new MouseEventHandler(frmTemp_MouseWheel);	
			frmTemp.Controls.Add(this.m_vScrollBar);
			this.m_vScrollBar.Visible=true;



			FIA_Biosum_Manager.TemplateGroupBox p_groupBox1 = new TemplateGroupBox(frmTemp);
			FIA_Biosum_Manager.TemplateTitle p_lblTitle = new TemplateTitle(p_groupBox1,0,0,"Wood Processing Site (Edit)");
			FIA_Biosum_Manager.TemplateOkCancelButtons p_OkCancel = new TemplateOkCancelButtons(frmTemp,p_groupBox1);
			p_groupBox1.Dock = System.Windows.Forms.DockStyle.None;
			FIA_Biosum_Manager.TemplateInputLabel p_lbl = new TemplateInputLabel(p_groupBox1,"temp","temp");
			FIA_Biosum_Manager.TemplateTextBox p_txt = new TemplateTextBox(p_groupBox1,"txttemp","");
			FIA_Biosum_Manager.TemplateComboBox p_cmb = new TemplateComboBox(p_groupBox1,"cmbtemp","");

			p_lbl.Visible=false;
			p_txt.Visible=false;
			p_cmb.Visible=false;

			//y controls the placement of the top of each object on the form
			y = p_lblTitle.Top + p_lblTitle.Height + 5;
			
			/*******************************************************************
			 **using the table schema information define the labels and 
			 **input that will be placed on the form
			 *******************************************************************/
			for (z=0; z<=this.m_dtTableSchema.Rows.Count-1;z++)
			{
				//get the name of the field
				strField=this.m_dtTableSchema.Rows[z]["columnname"].ToString().Trim();

				//get the column location on the grid that corresponds to strField
				intGridCol = this.getGridColumn(strField);

				/**********************************************************
				 **find the field and create the label and the input
				 **********************************************************/
				switch (strField.ToUpper())
				{
					case "PSITE_ID":
						p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),strField);
						p_lbl.Left = x;
						p_lbl.Top = y;
						p_txt = new TemplateTextBox(p_groupBox1, strField, this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim());
						p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
						p_txt.Top = p_lbl.Top;
						p_txt.Width = 200;
						p_txt.Visible=true;
						p_txt.ReadOnly=true;
						y += p_lbl.Height;
						p_lbl.Visible=true;
						break;
					case "TRANCD":
						break;
					case "BIOCD":
						break;
					case "TRANCD_DEF":
						//items in the combo menu box depend on the trancd value
						strValue="";
						if (this.m_dg[this.m_intCurrRow-1,this.getGridColumn("TRANCD")].ToString().Trim() == "1")
						{
							strValue = "Regular - PSite With Road Only Access";
						}
						else if (this.m_dg[this.m_intCurrRow-1,this.getGridColumn("TRANCD")].ToString().Trim() == "2")
						{
							strValue ="Railhead - Road To Rail Wood Transfer Point";
						}
						else if (this.m_dg[this.m_intCurrRow-1,this.getGridColumn("TRANCD")].ToString().Trim() == "3")
						{
							strValue = "Rail Collector - PSite With Both Road And Rail Access";
						}
						p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),"Transportation Code");
						p_lbl.Left = x;
						p_lbl.Top = y;
						p_cmb = new TemplateComboBox(p_groupBox1,strField.Trim(),strValue);
						p_cmb.Items.Add("Regular - PSite With Road Only Access");
						p_cmb.Items.Add("Railhead - Road To Rail Wood Transfer Point");
						p_cmb.Items.Add("Rail Collector - PSite With Both Road And Rail Access");
						p_cmb.Left = p_lbl.Left + p_lbl.Width + 5;
						p_cmb.Top = p_lbl.Top;
						intFieldSize = (int)this.CreateGraphics().MeasureString("Rail Collector - PSite With Both Road And Rail Access", p_cmb.Font).Width + 50;
						p_cmb.Width = intFieldSize;
						p_cmb.Visible=true;
						y += p_lbl.Height;
						p_lbl.Visible=true;
						break;
					case "BIOCD_DEF":
						//items in the combo menu box depend on the biocd value
						strValue="";
						if (this.m_dg[this.m_intCurrRow-1,this.getGridColumn("BIOCD")].ToString().Trim() == "1")
						{
							strValue = "Merchantable - Logs Only";
						}
						else if (this.m_dg[this.m_intCurrRow-1,this.getGridColumn("BIOCD")].ToString().Trim() == "2")
						{
							strValue = "Chip - Chips Only";
						}
						else if (this.m_dg[this.m_intCurrRow-1,this.getGridColumn("BIOCD")].ToString().Trim() == "3")
						{
							strValue = "Both - Logs And Chips";
						}
						p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),"Processing Abilities");
						p_lbl.Left = x;
						p_lbl.Top = y;
						p_cmb = new TemplateComboBox(p_groupBox1,strField.Trim(), strValue);
						p_cmb.Items.Add("Merchantable - Logs Only");
						p_cmb.Items.Add("Chip - Chips Only");
						p_cmb.Items.Add("Both - Logs And Chips");
						p_cmb.Left = p_lbl.Left + p_lbl.Width + 5;
						p_cmb.Top = p_lbl.Top;
						intFieldSize = (int)this.CreateGraphics().MeasureString("Merchantable - Logs Only", p_cmb.Font).Width + 50;
						p_cmb.Width = intFieldSize;
						p_cmb.Visible=true;
						y += p_lbl.Height;
						p_lbl.Visible=true;
						break;
					case "EXISTS_YN":
						p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),"Processing Site Exist");
						p_lbl.Left = x;
						p_lbl.Top = y;
						if (this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim() == "Y")
						{
							strValue = "Yes";
						}
						else if (this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim() == "N")
						{
							strValue = "No";
						}
						p_cmb = new TemplateComboBox(p_groupBox1,strField.Trim(),strValue);
						p_cmb.Items.Add("Yes");
						p_cmb.Items.Add("No");
						p_cmb.Left = p_lbl.Left + p_lbl.Width + 5;
						p_cmb.Top = p_lbl.Top;
						intFieldSize = (int)this.CreateGraphics().MeasureString("Yes", p_cmb.Font).Width + 50;
						p_cmb.Width = intFieldSize;
						p_cmb.Visible=true;
						y += p_lbl.Height;
						p_lbl.Visible=true;
						break;
                    case "KEEPLYR_YN":
                        //GitHub issue #56: Remove Keep Temp GIS Layer comboBox; Do nothing for now
                    //    p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),"Keep Temp GIS Layer");
                    //    p_lbl.Left = x;
                    //    p_lbl.Top = y;
                    //    if (this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim() == "Y")
                    //    {
                    //        strValue = "Yes";
                    //    }
                    //    else if (this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim() == "N")
                    //    {
                    //        strValue = "No";
                    //    }
                    //    p_cmb = new TemplateComboBox(p_groupBox1,strField.Trim(),strValue);
                    //    p_cmb.Items.Add("Yes");
                    //    p_cmb.Items.Add("No");
                    //    p_cmb.Left = p_lbl.Left + p_lbl.Width + 5;
                    //    p_cmb.Top = p_lbl.Top;
                    //    intFieldSize = (int)this.CreateGraphics().MeasureString("Yes", p_cmb.Font).Width + 50;
                    //    p_cmb.Width = intFieldSize;
                    //    p_cmb.Visible=true;
                    //    y += p_lbl.Height;
                    //    p_lbl.Visible=true;
                        break;
					case "NAME":
						p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),"Processing Site Name");
						p_lbl.Left = x;
						p_lbl.Top = y;
						p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim());
						p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
						p_txt.Top = p_lbl.Top;
						p_txt.Width = 400;
						p_txt.MaxLength = Convert.ToInt32(this.m_dtTableSchema.Rows[z]["columnsize"]);
						p_txt.Visible=true;
						y += p_lbl.Height;
						p_lbl.Visible=true;
						break;
					default:
						//process fields added by the user
						p_lbl = new TemplateInputLabel(p_groupBox1,"label" + z.ToString().Trim(),strField);
						p_lbl.Left = x;
						p_lbl.Top = y;
						//find the datatype of the field
					switch (this.m_dtTableSchema.Rows[z]["datatype"].ToString().Trim())
					{
						case "System.String" :
							p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim());
							p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
							p_txt.Top = p_lbl.Top;
							p_txt.Width = 200;
							p_txt.Visible=true;
							y += p_lbl.Height;
							p_lbl.Visible=true;
							p_txt.MaxLength = Convert.ToInt32(this.m_dtTableSchema.Rows[z]["columnsize"]);
							
							break;
						case "System.Double":
							p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim());
							p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
							p_txt.Top = p_lbl.Top;
							p_txt.Width = 200;
							p_txt.Visible=true;
							y += p_lbl.Height;
							p_lbl.Visible=true;
							break;
						case "System.Boolean":
							if (this.m_dg[this.m_intCurrRow-1,intGridCol] == System.DBNull.Value)
							{
								p_cmb = new TemplateComboBox(p_groupBox1,strField.Trim(),"");
							}
							else if (this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim().Length == 0)
								p_cmb = new TemplateComboBox(p_groupBox1,strField.Trim(),"");
							else 
							{
								p_cmb = new TemplateComboBox(p_groupBox1,strField.Trim(),this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim());
							}
							
							p_cmb.Items.Add("True");
							p_cmb.Items.Add("False");
							p_cmb.Left = p_lbl.Left + p_lbl.Width + 5;
							p_cmb.Top = p_lbl.Top;
							intFieldSize = (int)this.CreateGraphics().MeasureString("Yes", p_cmb.Font).Width + 50;
							p_cmb.Width = intFieldSize;
							p_cmb.Visible=true;
							y += p_lbl.Height;

							break;
						case "System.DateTime":
							p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim());
							p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
							p_txt.Top = p_lbl.Top;
							p_txt.Width = 200;
							p_txt.Visible=true;
							y += p_lbl.Height;
							p_lbl.Visible=true;

							
							break;
						case "System.Decimal":
							p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim());
							p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
							p_txt.Top = p_lbl.Top;
							p_txt.Width = 200;
							p_txt.Visible=true;
							y += p_lbl.Height;
							p_lbl.Visible=true;						
							break;
						case "System.Int16":
							p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim());
							p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
							p_txt.Top = p_lbl.Top;
							p_txt.Width = 200;
							p_txt.Visible=true;
							y += p_lbl.Height;
							p_lbl.Visible=true;						
							break;
						case "System.Int32":
							p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim());
							p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
							p_txt.Top = p_lbl.Top;
							p_txt.Width = 200;
							p_txt.Visible=true;
							y += p_lbl.Height;
							p_lbl.Visible=true;						
							break;
						case "System.Int64":
							p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim());
							p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
							p_txt.Top = p_lbl.Top;
							p_txt.Width = 200;
							p_txt.Visible=true;
							y += p_lbl.Height;
							p_lbl.Visible=true;						
							break;
						case "System.SByte":
							p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim());
							p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
							p_txt.Top = p_lbl.Top;
							p_txt.Width = 10;
							p_txt.Visible=true;
							p_txt.MaxLength = 1;
							y += p_lbl.Height;
							p_lbl.Visible=true;
								
							break;
						case "System.Byte":
							p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim());
							p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
							p_txt.Top = p_lbl.Top;
							p_txt.Width = 100;
							p_txt.Visible=true;
							p_txt.MaxLength = 1;
							y += p_lbl.Height;
							p_lbl.Visible=true;
							
							break;
						case "System.Single":
							p_txt = new TemplateTextBox(p_groupBox1, strField.Trim(), this.m_dg[this.m_intCurrRow-1,intGridCol].ToString().Trim());
							p_txt.Left = p_lbl.Left + p_lbl.Width + 5;
							p_txt.Top = p_lbl.Top;
							p_txt.Width = 100;
							p_txt.Visible=true;
							y += p_lbl.Height;
							p_lbl.Visible=true;						
							break;
						default:
							MessageBox.Show("Datatype Not Found:" + this.m_dtTableSchema.Rows[z]["datatype"].ToString().Trim());
							break;
					}
						break;
				}
				if (intFieldSize > intLargestFieldSize)
				{
					intLargestFieldSize = intFieldSize;
				}
				
			}
			
			//size the groupbox to encompass all the objects
			p_groupBox1.Width  = 0 ;  //p_txt.Left + intLargestFieldSize + 20;
			if (p_groupBox1.Left + p_groupBox1.Width < p_txt.Left + intLargestFieldSize + 20)
			{
				for (x=1;;x++)
				{
					p_groupBox1.Width = x;
					if (p_groupBox1.Width >
						p_groupBox1.Left + p_txt.Left + intLargestFieldSize + this.m_vScrollBar.Width + 5)
					{
						break;
					}
				}

			}


			//size the width of the groupbox and position the ok,cancel buttons
			//p_groupBox1.Width  = p_txt.Left + intLargestFieldSize + 10;
			p_OkCancel.btnOK.Top = p_lbl.Top + p_lbl.Height + 5;
			p_OkCancel.btnCancel.Top = p_OkCancel.btnOK.Top;
			p_OkCancel.btnOK.Left = (int)(p_groupBox1.Width * .50) - p_OkCancel.btnOK.Width;
			p_OkCancel.btnCancel.Left = p_OkCancel.btnOK.Left + p_OkCancel.btnOK.Width;
			p_groupBox1.Height = p_OkCancel.btnOK.Top + p_OkCancel.btnOK.Height + 10;


			//groupbox height
			if (p_groupBox1.Top + p_groupBox1.Height < p_OkCancel.btnOK.Top + p_OkCancel.btnOK.Height + 5)
			{
				for (x=1;;x++)
				{
					p_groupBox1.Height = x;
					if (p_groupBox1.Height + 5 < 
						p_OkCancel.btnOK.Top + p_OkCancel.btnOK.Height + 5)
					{
						break;
					}
				}

			}






			frmTemp.Height=0;
			frmTemp.Width=0;
			//size the form to be bigger than the groupbox

			//size the height
			if (p_groupBox1.Top + p_groupBox1.Height > frmTemp.ClientSize.Height + 2)
			{
				for (x=1;;x++)
				{
					frmTemp.Height = x;
					if (p_groupBox1.Top + 
						p_groupBox1.Height < 
						frmTemp.ClientSize.Height)
					{
						break;
					}
				}

			}
			//size the width
			if (p_groupBox1.Left + p_groupBox1.Width > frmTemp.ClientSize.Width + 2)
			{
				for (x=1;;x++)
				{
					frmTemp.Width = x;
					if (p_groupBox1.Left + 
						p_groupBox1.Width + this.m_vScrollBar.Width + 5 < 
						frmTemp.ClientSize.Width)
					{
						break;
					}
				}

			}
			frmTemp.DisposeOfFormWhenClosing = false;
			p_groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			p_groupBox1.Dock = System.Windows.Forms.DockStyle.None;
			frmTemp.Left = 0;
			frmTemp.Top = 0;
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            frmTemp.MinimizeBox = false;
            frmTemp.MaximizeBox = false;
			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

			this.m_vScrollBar.Maximum = p_groupBox1.Height; //p_OkCancel.btnOK.Top + (int)(p_OkCancel.btnOK.Height * 1.5);
			this.m_intMaxSize = p_groupBox1.Height;
			this.m_frmDialog = frmTemp;
			//loop until record passes validation or user cancels
			for (;;)
			{
				System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
				if (result==System.Windows.Forms.DialogResult.OK)
				{
					this.m_intError=0;
					/*******************************************************
                     **validate data by passing the groupbox since all the
					 **input controls are owned by p_groupBox1.
					 *******************************************************/
					this.val_data(p_groupBox1);
					if (this.m_intError==0)
					{
						//loop through all the controls in p_groupBox1 and save them to the datagrid
						for (z=0;z<=p_groupBox1.Controls.Count-1;z++)
						{
							//the name of the control corresponds to the name of the column on the grid
							switch (p_groupBox1.Controls[z].Name.Trim().ToUpper())
							{
								case "NAME":
									this.m_dg[this.m_intCurrRow-1,this.getGridColumn("NAME")] = p_groupBox1.Controls[z].Text;
									break;
                                //GitHub issue #56: Remove Keep Temp GIS Layer comboBox
                                //case "KEEPLYR_YN":
                                //    this.m_dg[this.m_intCurrRow-1,this.getGridColumn("KEEPLYR_YN")] = p_groupBox1.Controls[z].Text.Substring(0,1);
                                //    break;
								case "EXISTS_YN":
									this.m_dg[this.m_intCurrRow-1,this.getGridColumn("EXISTS_YN")] = p_groupBox1.Controls[z].Text.Substring(0,1);
									break;
								case "BIOCD_DEF":
									if (p_groupBox1.Controls[z].Text.Trim().Length > 0)
									{
										if (p_groupBox1.Controls[z].Text.IndexOf("Merchantable -",0) >=0)
										{
											this.m_dg[this.m_intCurrRow-1,this.getGridColumn("BIOCD")] = 1;
											this.m_dg[this.m_intCurrRow-1,this.getGridColumn("BIOCD_DEF")] = "Merchantable";
										}
										else if (p_groupBox1.Controls[z].Text.IndexOf("Chip -",0) >=0)
										{
											this.m_dg[this.m_intCurrRow-1,this.getGridColumn("BIOCD")] = 2;
											this.m_dg[this.m_intCurrRow-1,this.getGridColumn("BIOCD_DEF")] = "Chip";
										}
										else if (p_groupBox1.Controls[z].Text.IndexOf("Both -",0) >=0)
										{
											this.m_dg[this.m_intCurrRow-1,this.getGridColumn("BIOCD")] = 3;
											this.m_dg[this.m_intCurrRow-1,this.getGridColumn("BIOCD_DEF")] = "Both";
										}
									}
									break;
								case "TRANCD_DEF":
									if (p_groupBox1.Controls[z].Text.Trim().Length > 0)
									{
										if (p_groupBox1.Controls[z].Text.IndexOf("Regular",0) >=0)
										{
											this.m_dg[this.m_intCurrRow-1,this.getGridColumn("TRANCD")] = 1;
											this.m_dg[this.m_intCurrRow-1,this.getGridColumn("TRANCD_DEF")] = "Regular";
										}
										else if (p_groupBox1.Controls[z].Text.IndexOf("Railhead",0) >=0)
										{
										    this.m_dg[this.m_intCurrRow-1,this.getGridColumn("TRANCD")] = 2;
											this.m_dg[this.m_intCurrRow-1,this.getGridColumn("TRANCD_DEF")] = "Railhead";
										}
										else if (p_groupBox1.Controls[z].Text.IndexOf("Collector",0) >=0)
										{
										    this.m_dg[this.m_intCurrRow-1,this.getGridColumn("TRANCD")] = 3;
											this.m_dg[this.m_intCurrRow-1,this.getGridColumn("TRANCD_DEF")] = "Both";
										}
									}
									break;
								default:
									for (x=0;x<=this.m_dtTableSchema.Rows.Count-1;x++)
									{
										strField=this.m_dtTableSchema.Rows[x]["columnname"].ToString().Trim();
										if (strField.ToUpper() == p_groupBox1.Controls[z].Name.Trim().ToUpper())
										{
											if (p_groupBox1.Controls[z].Text == null || p_groupBox1.Controls[z].Text.Trim().Length == 0)
											{
											    this.m_dg[this.m_intCurrRow-1,this.getGridColumn(strField)] = System.DBNull.Value;
											}
											else
											{
												switch (this.m_dtTableSchema.Rows[x]["datatype"].ToString().Trim())
												{
													case "System.String" :
														this.m_dg[this.m_intCurrRow-1,this.getGridColumn(strField)] = Convert.ToString(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Double":
														this.m_dg[this.m_intCurrRow-1,this.getGridColumn(strField)] = Convert.ToDouble(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Boolean":
														this.m_dg[this.m_intCurrRow-1,this.getGridColumn(strField)] = Convert.ToBoolean(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.DateTime":
													    this.m_dg[this.m_intCurrRow-1,this.getGridColumn(strField)] = Convert.ToDateTime(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Decimal":
														this.m_dg[this.m_intCurrRow-1,this.getGridColumn(strField)] = Convert.ToDecimal(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Int16":
														this.m_dg[this.m_intCurrRow-1,this.getGridColumn(strField)] = Convert.ToInt16(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Int32":
														this.m_dg[this.m_intCurrRow-1,this.getGridColumn(strField)] = Convert.ToInt32(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Int64":
													    this.m_dg[this.m_intCurrRow-1,this.getGridColumn(strField)] = Convert.ToInt64(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.SByte":
														this.m_dg[this.m_intCurrRow-1,this.getGridColumn(strField)] = Convert.ToSByte(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Byte":
														this.m_dg[this.m_intCurrRow-1,this.getGridColumn(strField)] = Convert.ToByte(p_groupBox1.Controls[z].Text.Trim());
														break;
													case "System.Single":
														this.m_dg[this.m_intCurrRow-1,this.getGridColumn(strField)] = Convert.ToSingle(p_groupBox1.Controls[z].Text.Trim());
														break;
													default:
														MessageBox.Show("Datatype Not Found:" + this.m_dtTableSchema.Rows[z]["datatype"].ToString().Trim());
														break;
												}
											}

										}

									}
									break;
							}
						}
						if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
						break;
					}
				}
				else break;
					 
			}
			frmTemp.Close();
			frmTemp.Dispose();
			this.m_intMaxSize=0;

			this.m_dblNewPerc=0;
			this.m_dblOldPerc=0;
			this.m_intScrollValue=0;
			this.m_frmDialog=null;

			frmTemp=null;

		}

		/// <summary>
		/// find the table field in the grid by matching the field name with the column header
		/// </summary>
		/// <param name="strColumn">Column to search for</param>
		/// <returns>grid column number</returns>
		private int getGridColumn(string strColumn)
		{
			int x;
			int intGridCol=-1;
			for (x=0;x<=this.m_dg.TableStyles[0].GridColumnStyles.Count-1;x++)
			{
				if (strColumn.Trim().ToUpper() == this.m_dg.TableStyles[0].GridColumnStyles[x].HeaderText.Trim().ToUpper())
				{
					intGridCol=x;
					break;
				}
			}
			return intGridCol;
		}

		private void uc_gis_psite_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
				this.btnHelp.Top = this.btnClose.Top;
				this.m_dg.Width = this.groupBox1.Width  - (this.m_dg.Left * 2);
				this.m_dg.Height = this.groupBox1.Height - this.m_dg.Top - this.btnClose.Height - this.btnNew.Height - 50;
				this.btnNew.Top =this.m_dg.Height + this.m_dg.Top + 5;
				this.btnEdit.Top = this.btnNew.Top;
				this.btnSave.Top = this.btnNew.Top;
				this.btnDelete.Top = this.btnNew.Top;
				this.btnCancel.Top = this.btnNew.Top;
				this.btnDelete.Left = (int)(this.m_dg.Width * .50) - (int)(this.btnDelete.Width * .50);
				this.btnEdit.Left = this.btnDelete.Left - this.btnEdit.Width;
				this.btnNew.Left = this.btnEdit.Left - this.btnNew.Width;
				this.btnSave.Left = this.btnDelete.Left + this.btnSave.Width;
				this.btnCancel.Left = this.btnSave.Left + this.btnCancel.Width;

			}
			catch
			{
			}
		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
			this.savevalues();
		}
		private void savevalues()
		{
			int intCurrRow;
		

			this.m_intError=0;

			/******************************************************
			 **save the current row, move the current row to a
			 **different row to enable getchanges() method, then
			 **move back to current row
			 ******************************************************/
			intCurrRow = this.m_dg.CurrentRowIndex;
			if (intCurrRow==0)
			{
				this.m_dg.CurrentRowIndex++;
			}
			else
			{
				this.m_dg.CurrentRowIndex=0;
			}
			 
               			
			System.Data.DataTable p_dtChanges;

			try
			{
				
				p_dtChanges = m_dataMgr.m_DataSet.Tables["processing_site"].GetChanges();
								
				//check if any inserted rows
				//p_Rows = p_dtChanges.Select(null,null, DataViewRowState.Added);
				if (p_dtChanges.HasErrors)
				{
                    m_dataMgr.m_DataSet.Tables["processing_site"].RejectChanges();
					this.m_intError=-1;
				}
				else
				{
                    m_dataMgr.m_DataAdapter.Update(m_dataMgr.m_DataSet.Tables["processing_site"]);
                    m_dataMgr.m_Transaction.Commit();
                    m_dataMgr.m_DataSet.Tables["processing_site"].AcceptChanges();
                    this.InitializeTransactionCommands();
				}
				
			}
			catch (Exception caught)
			{
				this.m_intError=-1;
				MessageBox.Show(caught.Message);
                m_dataMgr.m_DataSet.Tables["processing_site"].RejectChanges();
                //rollback the transaction to the original records 
                m_dataMgr.m_Transaction.Rollback();				
			}
		
			p_dtChanges=null;
			
			this.m_dg.CurrentRowIndex = intCurrRow;		
			this.btnSave.Enabled=false;	
            
			return;
		}

		private void btnDelete_Click(object sender, System.EventArgs e)
		{
			this.DeleteRecords();
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.CleanUp();
			this.ParentForm.Close();
		}

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "DATABASE", "WOOD_PROCESSING_SITES" });
        }

	}
	public class PSite_DataGridColoredTextBoxColumn : DataGridTextBoxColumn
	{
		bool m_bEdit=false;
		FIA_Biosum_Manager.uc_gis_psite uc_gis_psite1;
		string m_strLastKey="";
		bool m_bNumericOnly=false;
		

		public PSite_DataGridColoredTextBoxColumn(bool bEdit,bool bNumericOnly,FIA_Biosum_Manager.uc_gis_psite p_uc)
		{
			this.m_bEdit = bEdit;
			this.m_bNumericOnly = bNumericOnly;
			this.uc_gis_psite1 = p_uc;
			this.TextBox.KeyDown += new KeyEventHandler(TextBox_KeyDown);
			this.TextBox.Leave += new EventHandler(TextBox_Leave);
			this.TextBox.Enter += new EventHandler(TextBox_Enter);
		}
		

		protected override void Paint(System.Drawing.Graphics g, System.Drawing.Rectangle bounds, System.Windows.Forms.CurrencyManager source, int rowNum, System.Drawing.Brush backBrush, System.Drawing.Brush foreBrush, bool alignToRight)
		{
		  	
			// color only the columns that can be edited by the user
			try
			{
				if (this.m_bEdit == true)
				{
					backBrush = new System.Drawing.Drawing2D.LinearGradientBrush(bounds,
						Color.FromArgb(255, 200, 200), 
						Color.FromArgb(128, 20, 20),
						System.Drawing.Drawing2D.LinearGradientMode.BackwardDiagonal);
					foreBrush = new SolidBrush(Color.White);
				}
			}
			catch { /* empty catch */ }
			finally
			{
				// make sure the base class gets called to do the drawing with
				// the possibly changed brushes
				base.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight);
			}
		}
		private void TextBox_TextChanged(object sender, EventArgs e)
		{
			//MessageBox.Show("textchange");
		}
		private void TextBox_KeyDown(object sender, KeyEventArgs e)
		{
			if (this.m_bEdit == true)
			{
				//if (this.m_strLastKey.Trim().Length == 0) 
				//{
				if (this.m_bNumericOnly==true)
				{
					if (Char.IsDigit((char)e.KeyValue))
					{
						this.m_strLastKey = Convert.ToString(e.KeyValue);
						if (this.uc_gis_psite1.btnSave.Enabled==false) this.uc_gis_psite1.btnSave.Enabled=true;
					}
					else
					{
						if (e.KeyCode == Keys.Back)
						{
							this.m_strLastKey = Convert.ToString(e.KeyValue);
							if (this.uc_gis_psite1.btnSave.Enabled==false) this.uc_gis_psite1.btnSave.Enabled=true;
						}
						else
						{
							e.Handled=true;	
							SendKeys.Send("{BACKSPACE}");
						}
					}
						
				}
				else
				{
					this.m_strLastKey = Convert.ToString(e.KeyValue);
					if (this.uc_gis_psite1.btnSave.Enabled==false) this.uc_gis_psite1.btnSave.Enabled=true;
				}
					
				//}
			}
			//if (this.m_bEdit==true) this.uc_gridview1.btnSave.Enabled=true;
		}
		private void TextBox_Enter(object sender, EventArgs e)
		{
			this.m_strLastKey="";
		}
		private void TextBox_Leave(object sender, EventArgs e)
		{
			if (this.m_bEdit == true)
			{
				if (this.m_strLastKey.Trim().Length > 0)
				{
					if (this.uc_gis_psite1.btnSave.Enabled==false) this.uc_gis_psite1.btnSave.Enabled=true;
				}
			}
		}
		     
	}
}

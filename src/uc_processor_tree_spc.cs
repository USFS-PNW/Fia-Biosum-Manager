using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Collections.Generic;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_tree_spc_conversion_and_groupings.
	/// </summary>
	public class uc_processor_tree_spc : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.GroupBox grpboxAudit;
        public System.Windows.Forms.ListView lstAudit;
		private System.Windows.Forms.Button btnHelp;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnAuditAdd;
		private System.Windows.Forms.Button btnAuditCheckAll;
		private System.Windows.Forms.Button btnAuditClearAll;
		private string m_strProjDir;
		private string m_strTempDbFile;
        private ODBCMgr m_odbcmgr = new ODBCMgr();
		private System.Data.DataView m_dv;
		public int m_intIndex=0;
		private int m_intCurrRow=0;
		public int m_intError=0;
        private bool handleCheck = true;
		private SQLite.ADO.DataMgr SQLite { get; set; } = new SQLite.ADO.DataMgr();


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
		private bool m_bDelete=false;
		private string m_strDeletedList="";
		private int m_intDeletedCount=0;
		private string m_strColumnFilterList="";
		private string m_strColumnSortList="";
		//private string m_strTreeSpCdList="";

		private System.Data.DataTable m_dtTableSchema;
		private System.Windows.Forms.ComboBox cmbAudit;
		private System.Windows.Forms.Button btnAudit;


		Queries m_oQueries = new Queries();
		RxTools m_oRxTools = new RxTools();
        FIA_Biosum_Manager.RxPackageItem_Collection m_oRxPackageItem_Collection = new RxPackageItem_Collection();
		FIA_Biosum_Manager.RxPackageItem m_oRxPackageItem=null;
		string m_strRxCycleList="";
		string[] m_strRxCycleArray=null;
        private Button btnView;

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultFvsXPSFile;

        private string m_strDebugFile = frmMain.g_oEnv.strTempDir + "\\biosum_processor_audit_debug.txt";
        private GroupBox grpBoxTreeSpc;
        public Button btnDelete;
        private Button btnNew;
        private Button btnEdit;
        public Button btnSave;
        private Button btnCancel;
        private DataGrid m_dg;

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_processor_tree_spc(string p_strProjDir)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call

			this.m_strProjDir = p_strProjDir;
            this.m_oEnv = new env();

			this.m_oQueries = new Queries();
			m_oQueries.m_oFvs.LoadDatasource=true;
			m_oQueries.m_oFIAPlot.LoadDatasource=true;
			m_oQueries.LoadDatasources(true);

            m_oRxTools.LoadAllRxPackageItems(m_oRxPackageItem_Collection);			

			this.cmbAudit.Text = Convert.ToString(this.cmbAudit.Items[0]).Trim();

            if (System.IO.File.Exists(m_strDebugFile))
                System.IO.File.Delete(m_strDebugFile);

			// Used by Processor audit
			m_strTempDbFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");

			System.Threading.Thread.Sleep(2000);
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
			SQLite.m_DataSet = new DataSet("tree_species");
			SQLite.m_DataAdapter = new System.Data.SQLite.SQLiteDataAdapter();
			string strMasterConn = SQLite.GetConnectionString(m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Tree));
			SQLite.OpenConnection(strMasterConn);
			// Attach tree_species database so we can access tree_species table
			SQLite.m_strSQL = $@"ATTACH '{m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.TreeSpecies)}' as REF_MASTER";
			SQLite.SqlNonQuery(SQLite.m_Connection, SQLite.m_strSQL);
			// Attach fvs_tree_species database so we can access fvs_tree_species table
			SQLite.m_strSQL = $@"ATTACH '{m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.FvsTreeSpecies)}' as BIOSUM_REF";
			SQLite.SqlNonQuery(SQLite.m_Connection, SQLite.m_strSQL);

			this.InitializeSqliteTransactionCommands();
			//SQLite.m_strSQL = "SELECT s.fvs_variant, s.spcd, s.common_name, " +
			//                      "s.fvs_input_spcd, f.fvs_species, f.fvs_common_name, " +
			//                      "s.genus, s.species, comments, s.id " +
			//                      "FROM " + this.m_oQueries.m_oFvs.m_strTreeSpcTable + " s, " +
			//                      this.m_oQueries.m_oFvs.m_strFvsTreeSpcRefTable + " f " +
			//                      "WHERE EXISTS (SELECT DISTINCT(spcd) FROM " + this.m_oQueries.m_oFIAPlot.m_strTreeTable + " t WHERE s.spcd=t.spcd) " +
			//                      "AND s.fvs_input_spcd = f.spcd " +
			//                      "AND s.fvs_variant = f.fvs_variant " +
			//                      "ORDER BY s.fvs_variant, s.spcd; ";
			// 2025-03-19: Change this query so spcd missing from fvs_tree_species will still appear on the grid but with null fields
			SQLite.m_strSQL = $@"SELECT s.fvs_variant, s.spcd, s.common_name, s.fvs_input_spcd, f.fvs_species, f.fvs_common_name, s.genus, s.species, comments, s.id
				FROM {this.m_oQueries.m_oFvs.m_strTreeSpcTable} s LEFT JOIN {this.m_oQueries.m_oFvs.m_strFvsTreeSpcRefTable} f ON s.fvs_input_spcd = f.spcd AND s.fvs_variant = f.fvs_variant
				WHERE EXISTS (SELECT DISTINCT(spcd) FROM tree t WHERE s.spcd=t.spcd)
				ORDER BY s.fvs_variant, s.spcd";
			this.m_dtTableSchema = SQLite.getTableSchema(SQLite.m_Connection, SQLite.m_Transaction, SQLite.m_strSQL);

			if (SQLite.m_intError == 0)
			{
				SQLite.m_Command = SQLite.m_Connection.CreateCommand();
				SQLite.m_Command.CommandText = SQLite.m_strSQL;
				SQLite.m_DataAdapter.SelectCommand = SQLite.m_Command;
				SQLite.m_DataAdapter.SelectCommand.Transaction = SQLite.m_Transaction;
				try 
				{

					SQLite.m_DataAdapter.Fill(SQLite.m_DataSet,"tree_species");
					this.m_dv = new DataView(SQLite.m_DataSet.Tables["tree_species"]);
				
					this.m_dv.AllowNew = false;       //cannot append new records
					this.m_dv.AllowDelete = false;    //cannot delete records
					this.m_dv.AllowEdit = true;
					this.m_dg.CaptionText = "tree_species";
					m_dg.BackgroundColor=frmMain.g_oGridViewBackgroundColor;
					/***********************************************************************************
					 **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
					 ***********************************************************************************/
					TreeSpcAudit_DataGridColoredTextBoxColumn aColumnTextColumn ;


					/***************************************************************
					 **custom define the grid style
					 ***************************************************************/
					DataGridTableStyle tableStyle = new DataGridTableStyle();

					/***********************************************************************
					 **map the data grid table style to the scenario rx intensity dataset
					 ***********************************************************************/
					tableStyle.MappingName = "tree_species";
					tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
					tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
					tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
					tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
					
					
   
					/******************************************************************************
					 **since the dataset has things like field name and number of columns,
					 **we will use those to create new columnstyles for the columns in our grid
					 ******************************************************************************/
					//get the number of columns from the scenario_rx_intensity data set
					int numCols = SQLite.m_DataSet.Tables["tree_species"].Columns.Count;
                
                    
					/************************************************
					 **loop through all the columns in the dataset	
					 ************************************************/
					for(int i = 0; i < numCols; ++i)
					{
                        strColumnName = SQLite.m_DataSet.Tables["tree_species"].Columns[i].ColumnName;

					    /******************************************************************
					    **create a new instance of the DataGridColoredTextBoxColumn class
						******************************************************************/
						aColumnTextColumn = new TreeSpcAudit_DataGridColoredTextBoxColumn(false,false,this);
						/***********************************
						**all columns are read-only
						***********************************/
						aColumnTextColumn.ReadOnly=true;
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
					if (frmMain.g_oGridViewFont != null) this.m_dg.Font = frmMain.g_oGridViewFont;
					this.m_dg.TableStyles.Clear();
					this.m_dg.TableStyles.Add(tableStyle);
                
                    string strCaption = "                                 |                    FIA" +
                                        "                      |                           FVS Input" +
                                        "                           |                              FIA";
                    this.m_dg.CaptionText = strCaption;
                    this.m_dg.DataSource = this.m_dv;  

					if (SQLite.m_DataSet.Tables["tree_species"].Rows.Count > 0)
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
					SQLite.m_Connection.Close();
					SQLite.m_Connection = null;
					SQLite.m_DataSet.Clear();
					SQLite.m_DataSet= null;
					SQLite.m_DataAdapter.Dispose();
					SQLite.m_DataAdapter = null;
					SQLite = null;
					return;

				}

			
			

				if (SQLite.m_DataSet.Tables["tree_species"].Rows.Count == 0)
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
			if (SQLite.m_intError < 0)
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

					if (strCol.ToUpper()== "fvs_input_spcd")
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
					else if (strCol.ToUpper() == "SPECIES" || strCol.ToUpper() == "COMMON_NAME" || strCol.ToUpper() == "GENUS")
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
						frmTemp.Text = "Modify (" + strCol.Trim() + ")";
		                    
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
						
						if (strCol.ToUpper() == "GENUS")
						{
							p_textbox.MaxLength = 20;
						}
						else
							p_textbox.MaxLength = 50;

						
						p_textbox.Text = "";
						intLeft = p_textbox.Left;
						intTop = p_textbox.Top;
						frmTemp.txtBox = p_textbox;
						frmTemp.strCallingFormType="TS";
					}
					else if (strCol.ToUpper() == "COMMENTS")
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
						frmTemp.Text = "Modify (" + strCol.Trim() + ")";
		                    
						//define numeric text class properties
						p_textbox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
						p_textbox.Name = "txtModify";
						p_textbox.TabIndex = 0;
						p_textbox.Tag = "";
						p_textbox.Visible = true;
						p_textbox.Enabled = true;
						frmTemp.Controls.Add(p_textbox);  //add the text box to the form
						p_textbox.Height = 100;

						//p_textbox.Top = (int)(frmTemp.ClientSize.Height * .50) - (int)(p_textbox.Height * .50);
						p_textbox.Top = 10;
						p_textbox.Width = frmTemp.ClientSize.Width - 20;
						p_textbox.Left = 10;
						
						p_textbox.MaxLength = 200;
						
						p_textbox.Text = "";
						intLeft = p_textbox.Left;
						intTop = p_textbox.Top + p_textbox.Height;
						p_textbox.Multiline=true;
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
						btnTempOK.Top = frmTemp.ClientSize.Height -  (int)(btnTempOK.Height * 1.5); //intTop + btnTempOK.Height + 10;
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
						btnTempCancel.Top = btnTempOK.Top; //intTop + btnTempCancel.Height + 10;
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
											m_dg[x,this.m_intPopupColumn] = frmTemp.txtNumeric.Text;
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
							MessageBox.Show(strCol + " Maximum: " + SQLite.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
						}
						else
						{
							MessageBox.Show(strCol + " Maximum: " + SQLite.m_DataSet.Tables[0].Compute(strFilter, null));
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
							MessageBox.Show(strCol + " Minimum: " + SQLite.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
						}
						else
						{
							MessageBox.Show(strCol + " Minimum: " + SQLite.m_DataSet.Tables[0].Compute(strFilter, null));
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
							MessageBox.Show(strCol + " Average: " + SQLite.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
						}
						else
						{
							MessageBox.Show(strCol + " Average: " + SQLite.m_DataSet.Tables[0].Compute(strFilter, null));
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
							MessageBox.Show(strCol + " Sum: " + SQLite.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
						}
						else
						{
							MessageBox.Show(strCol + " Sum: " + SQLite.m_DataSet.Tables[0].Compute(strFilter, null));
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
							MessageBox.Show(strCol + " Count: " + SQLite.m_DataSet.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString() + " and " + strExp));
						}
						else
						{
							MessageBox.Show(strCol + " Count: " + SQLite.m_DataSet.Tables[0].Compute(strFilter, strExp));
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
							this.m_strDeletedList += "," +  p_dv[x]["id"].ToString().Trim();
						}
						else
						{
							this.m_strDeletedList = p_dv[x]["id"].ToString().Trim();
						}
						strList += "," + p_dv[x]["id"].ToString().Trim() + ",";
						this.m_intDeletedCount++;						
					}
				}
			if (this.m_intDeletedCount > 0)
			{
				string strMsg = this.m_intDeletedCount.ToString().Trim() + " record(s) will be deleted from the table. \n Do you want to delete the record(s) from the table? (Y/N)";
				DialogResult result = MessageBox.Show(strMsg,"Delete Selected Tree Species", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
				if (result == DialogResult.Yes)
				{
					this.m_dv.AllowDelete = true;

					if (this.m_intError==0 && SQLite.m_intError==0)
					{
						//delete any plots that are in the delete list
						for (x=0; x < p_dv.Count;x++)
						{
							strValue = "," + p_dv[x]["id"].ToString().Trim() + ",";
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

		private void m_dg_CurrentCellChanged(object sender, 
			System.EventArgs e)
		{
			if (this.m_intCurrRow > 0)
			{
				if (this.m_dg.CurrentRowIndex != this.m_intCurrRow - 1)
				{
					this.m_intCurrRow = this.m_dg.CurrentRowIndex + 1;
				}
			}
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
						if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==true)
							this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;
						if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==false)
							this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;

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

		private void m_dg_TextBox_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			switch (e.Button)
			{
				case System.Windows.Forms.MouseButtons.Right :
					this.m_mnuDataGridPopup.Show(this,new Point(e.X,e.Y));
					MessageBox.Show("here i am");
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


		public string strProjectDirectory
		{
			set
			{
				this.m_strProjDir = value;
			}
			get
			{
				return this.m_strProjDir;
			}
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
            this.grpBoxTreeSpc = new System.Windows.Forms.GroupBox();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnNew = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.m_dg = new System.Windows.Forms.DataGrid();
            this.grpboxAudit = new System.Windows.Forms.GroupBox();
            this.btnView = new System.Windows.Forms.Button();
            this.cmbAudit = new System.Windows.Forms.ComboBox();
            this.btnAudit = new System.Windows.Forms.Button();
            this.btnAuditClearAll = new System.Windows.Forms.Button();
            this.btnAuditCheckAll = new System.Windows.Forms.Button();
            this.btnAuditAdd = new System.Windows.Forms.Button();
            this.lstAudit = new System.Windows.Forms.ListView();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.grpBoxTreeSpc.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.m_dg)).BeginInit();
            this.grpboxAudit.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.grpBoxTreeSpc);
            this.groupBox1.Controls.Add(this.grpboxAudit);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(736, 616);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(600, 576);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 47;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(8, 576);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(96, 32);
            this.btnHelp.TabIndex = 46;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // grpBoxTreeSpc
            // 
            this.grpBoxTreeSpc.Controls.Add(this.btnDelete);
            this.grpBoxTreeSpc.Controls.Add(this.btnNew);
            this.grpBoxTreeSpc.Controls.Add(this.btnEdit);
            this.grpBoxTreeSpc.Controls.Add(this.btnSave);
            this.grpBoxTreeSpc.Controls.Add(this.btnCancel);
            this.grpBoxTreeSpc.Controls.Add(this.m_dg);
            this.grpBoxTreeSpc.Location = new System.Drawing.Point(24, 304);
            this.grpBoxTreeSpc.Name = "grpBoxTreeSpc";
            this.grpBoxTreeSpc.Size = new System.Drawing.Size(656, 256);
            this.grpBoxTreeSpc.TabIndex = 29;
            this.grpBoxTreeSpc.TabStop = false;
            this.grpBoxTreeSpc.Text = "Tree Species Table";
            // 
            // btnDelete
            // 
            this.btnDelete.Enabled = false;
            this.btnDelete.Location = new System.Drawing.Point(311, 220);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(64, 32);
            this.btnDelete.TabIndex = 51;
            this.btnDelete.Text = "Delete";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnNew
            // 
            this.btnNew.Location = new System.Drawing.Point(183, 220);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(64, 32);
            this.btnNew.TabIndex = 46;
            this.btnNew.Text = "New";
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(247, 220);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(64, 32);
            this.btnEdit.TabIndex = 47;
            this.btnEdit.Text = "Edit";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnSave
            // 
            this.btnSave.Enabled = false;
            this.btnSave.Location = new System.Drawing.Point(375, 220);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(64, 32);
            this.btnSave.TabIndex = 48;
            this.btnSave.Text = "Save";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(439, 220);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(64, 32);
            this.btnCancel.TabIndex = 49;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // m_dg
            // 
            this.m_dg.DataMember = "";
            this.m_dg.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.m_dg.Location = new System.Drawing.Point(8, 24);
            this.m_dg.Name = "m_dg";
            this.m_dg.Size = new System.Drawing.Size(640, 192);
            this.m_dg.TabIndex = 2;
            this.m_dg.MouseDown += new System.Windows.Forms.MouseEventHandler(this.m_dg_MouseDown);
            this.m_dg.MouseUp += new System.Windows.Forms.MouseEventHandler(this.m_dg_MouseUp);
            // 
            // grpboxAudit
            // 
            this.grpboxAudit.Controls.Add(this.btnView);
            this.grpboxAudit.Controls.Add(this.cmbAudit);
            this.grpboxAudit.Controls.Add(this.btnAudit);
            this.grpboxAudit.Controls.Add(this.btnAuditClearAll);
            this.grpboxAudit.Controls.Add(this.btnAuditCheckAll);
            this.grpboxAudit.Controls.Add(this.btnAuditAdd);
            this.grpboxAudit.Controls.Add(this.lstAudit);
            this.grpboxAudit.Location = new System.Drawing.Point(24, 48);
            this.grpboxAudit.Name = "grpboxAudit";
            this.grpboxAudit.Size = new System.Drawing.Size(688, 248);
            this.grpboxAudit.TabIndex = 28;
            this.grpboxAudit.TabStop = false;
            this.grpboxAudit.Text = "Audit Results";
            // 
            // btnView
            // 
            this.btnView.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnView.Location = new System.Drawing.Point(536, 136);
            this.btnView.Name = "btnView";
            this.btnView.Size = new System.Drawing.Size(136, 32);
            this.btnView.TabIndex = 35;
            this.btnView.Text = "View Affected Trees";
            this.btnView.Visible = false;
            this.btnView.Click += new System.EventHandler(this.btnView_Click);
            // 
            // cmbAudit
            // 
            this.cmbAudit.Items.AddRange(new object[] {
            "Assess Data Readiness: Check If Each Tree And FVS Variant Combination Is In The T" +
                "ree Species Table",
            "Join Tree Table, FVS Tree Output Table And Tree Species Table And List  FIA SpCd+" +
                "FVS SpCd+Variant Combinations Not Found In The Tree Species Table",
            "Join FVS Tree Output Table And Tree Species Table  And List Incomplete Dry And We" +
                "t Tree Ratios In The Tree Species Table"});
            this.cmbAudit.Location = new System.Drawing.Point(16, 184);
            this.cmbAudit.Name = "cmbAudit";
            this.cmbAudit.Size = new System.Drawing.Size(656, 21);
            this.cmbAudit.TabIndex = 34;
            this.cmbAudit.Text = "Assess Data Readiness: Check If Each Tree And FVS Variant Combination Is In The T" +
    "ree Species Table";
            this.cmbAudit.SelectedIndexChanged += new System.EventHandler(this.cmbAudit_SelectedIndexChanged);
            // 
            // btnAudit
            // 
            this.btnAudit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAudit.Location = new System.Drawing.Point(16, 215);
            this.btnAudit.Name = "btnAudit";
            this.btnAudit.Size = new System.Drawing.Size(656, 24);
            this.btnAudit.TabIndex = 33;
            this.btnAudit.Text = "Run Audit";
            this.btnAudit.Click += new System.EventHandler(this.btnAudit_Click);
            // 
            // btnAuditClearAll
            // 
            this.btnAuditClearAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAuditClearAll.Location = new System.Drawing.Point(88, 136);
            this.btnAuditClearAll.Name = "btnAuditClearAll";
            this.btnAuditClearAll.Size = new System.Drawing.Size(72, 32);
            this.btnAuditClearAll.TabIndex = 32;
            this.btnAuditClearAll.Text = "Clear All";
            this.btnAuditClearAll.Click += new System.EventHandler(this.btnAuditClearAll_Click);
            // 
            // btnAuditCheckAll
            // 
            this.btnAuditCheckAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAuditCheckAll.Location = new System.Drawing.Point(16, 136);
            this.btnAuditCheckAll.Name = "btnAuditCheckAll";
            this.btnAuditCheckAll.Size = new System.Drawing.Size(72, 32);
            this.btnAuditCheckAll.TabIndex = 31;
            this.btnAuditCheckAll.Text = "Check All";
            this.btnAuditCheckAll.Click += new System.EventHandler(this.btnAuditCheckAll_Click);
            // 
            // btnAuditAdd
            // 
            this.btnAuditAdd.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAuditAdd.Location = new System.Drawing.Point(160, 136);
            this.btnAuditAdd.Name = "btnAuditAdd";
            this.btnAuditAdd.Size = new System.Drawing.Size(312, 32);
            this.btnAuditAdd.TabIndex = 30;
            this.btnAuditAdd.Text = "Add Checked Items To Tree Species Table";
            this.btnAuditAdd.Click += new System.EventHandler(this.btnAuditAdd_Click);
            // 
            // lstAudit
            // 
            this.lstAudit.CheckBoxes = true;
            this.lstAudit.FullRowSelect = true;
            this.lstAudit.GridLines = true;
            this.lstAudit.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstAudit.HideSelection = false;
            this.lstAudit.Location = new System.Drawing.Point(16, 24);
            this.lstAudit.MultiSelect = false;
            this.lstAudit.Name = "lstAudit";
            this.lstAudit.Size = new System.Drawing.Size(656, 104);
            this.lstAudit.TabIndex = 27;
            this.lstAudit.UseCompatibleStateImageBehavior = false;
            this.lstAudit.View = System.Windows.Forms.View.Details;
            this.lstAudit.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lstAudit_ItemCheck);
            this.lstAudit.SelectedIndexChanged += new System.EventHandler(this.lstAudit_SelectedIndexChanged);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(730, 24);
            this.lblTitle.TabIndex = 1;
            this.lblTitle.Text = "Tree Species";
            // 
            // uc_processor_tree_spc
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_tree_spc";
            this.Size = new System.Drawing.Size(736, 616);
            this.Resize += new System.EventHandler(this.uc_tree_spc_conversion_Resize);
            this.groupBox1.ResumeLayout(false);
            this.grpBoxTreeSpc.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.m_dg)).EndInit();
            this.grpboxAudit.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.CleanUp();
            ((frmDialog)ParentForm).ParentControl.Enabled = true;
			this.ParentForm.Close();
		}

		private void uc_tree_spc_conversion_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
				this.btnHelp.Top = this.btnClose.Top;
				this.grpboxAudit.Width = this.groupBox1.Width  - (this.grpboxAudit.Left * 2);
				this.grpBoxTreeSpc.Width = this.grpboxAudit.Width ;
				this.lstAudit.Width = this.grpboxAudit.Width - (this.lstAudit.Left * 2);
				this.cmbAudit.Width = this.lstAudit.Width;
				this.btnAudit.Width = this.lstAudit.Width;


				this.m_dg.Width = this.grpBoxTreeSpc.Width - (this.m_dg.Left * 2);
				this.grpBoxTreeSpc.Height = this.btnClose.Top - this.grpBoxTreeSpc.Top - 5;
				this.btnNew.Top =this.grpBoxTreeSpc.Height - this.btnNew.Height - 2;
				this.btnEdit.Top = this.btnNew.Top;
				this.btnSave.Top = this.btnNew.Top;
				this.btnDelete.Top = this.btnNew.Top;
				this.btnCancel.Top = this.btnNew.Top;
				this.m_dg.Height = this.btnNew.Top - this.m_dg.Top - 2;
				this.btnDelete.Left = (int)(this.grpBoxTreeSpc.Width * .50) - (int)(this.btnDelete.Width * .50);
				this.btnEdit.Left = this.btnDelete.Left - this.btnEdit.Width;
				this.btnNew.Left = this.btnEdit.Left - this.btnNew.Width;
				this.btnSave.Left = this.btnDelete.Left + this.btnSave.Width;
				this.btnCancel.Left = this.btnSave.Left + this.btnCancel.Width;

			}
			catch
			{
			}
		}


		private void btnAuditCheckAll_Click(object sender, System.EventArgs e)
		{
			if (this.lstAudit.Items.Count==0) return;
			for (int x=0;x<=this.lstAudit.Items.Count-1;x++)
			{
                if (cmbAudit.Text.Trim() == "Assess Data Readiness: Check If Each FIA Tree Spc, FVS Variant, And FVS Tree Spc Combination Is In The Tree Spc Table")
                {
                    if (this.lstAudit.Items[x].SubItems[lstAudit.Columns.Count - 1].Text.Trim() == "N")
                    {
                        this.lstAudit.Items[x].Checked = true;
                    }
                }
                else this.lstAudit.Items[x].Checked = true;
			}
		}

		private void btnAuditClearAll_Click(object sender, System.EventArgs e)
		{
			if (this.lstAudit.Items.Count==0) return;
			if (this.lstAudit.CheckedItems.Count==0) return;
			for (int x=0;x<=this.lstAudit.Items.Count-1;x++)
			{
				this.lstAudit.Items[x].Checked=false;
			}
		}

		private void btnAuditAdd_Click(object sender, System.EventArgs e)
		{
            if (this.lstAudit.CheckedItems.Count == 0)
            {
                    MessageBox.Show("!!No Items Selected!!", "Tree Species",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
            }

            int x,y;
            string strCommonName = "";
            string strGenus = "";
            string strSpc = "";

            if (this.btnSave.Enabled == false) this.btnSave.Enabled = true;
            if (this.btnDelete.Enabled == false) this.btnDelete.Enabled = true;
            if (this.btnEdit.Enabled == false) this.btnEdit.Enabled = true;
               
                try
                {
                    this.m_dv.AllowNew = true;

                    int intId = this.getUniqueId();

                    for (x = 0; x <= this.lstAudit.Items.Count - 1; x++)
                    {
                        if (this.lstAudit.Items[x].Checked == true)
                        {
                            //check to make sure the spcd exists in fia_tree_species_ref before trying to add
                            SQLite.m_strSQL = "SELECT common_name FROM " + Tables.ProcessorScenarioRun.DefaultFiaTreeSpeciesRefTableName +
                                                  " WHERE spcd = " + this.lstAudit.Items[x].SubItems[1].Text.Trim();
							SQLite.SqlQueryReader(SQLite.m_Connection, SQLite.m_Transaction, SQLite.m_strSQL);
                            if (!SQLite.m_DataReader.HasRows)
                            {
                                string strMessage = "spcd " + this.lstAudit.Items[x].SubItems[1].Text.Trim() + " is missing from the " +
                                                    Tables.ProcessorScenarioRun.DefaultFiaTreeSpeciesRefTableName + " table. If this is a " +
                                                    "valid FIA species code, please contact the BioSum Administrator to have it added.";
                                System.Windows.Forms.MessageBox.Show(strMessage, "FIA Biosum");
                                return;
                            }
                            
                            //get the species information if a record with the same species already exists
                            SQLite.m_strSQL = "SELECT DISTINCT common_name,genus,species " +
                                "FROM " + this.m_oQueries.m_oFvs.m_strTreeSpcTable + " t " +
                                "WHERE SPCD = " + this.lstAudit.Items[x].SubItems[1].Text.Trim() + ";";

                            SQLite.SqlQueryReader(SQLite.m_Connection, SQLite.m_Transaction, SQLite.m_strSQL);

                            strCommonName = "";
                            strGenus = "";
                            strSpc = "";

                            if (SQLite.m_DataReader.HasRows)
                            {
							SQLite.m_DataReader.Read();
                                if (SQLite.m_DataReader["common_name"] != System.DBNull.Value)
                                {
                                    strCommonName = Convert.ToString(SQLite.m_DataReader["common_name"]).Trim();
                                    if (SQLite.m_DataReader["genus"] != System.DBNull.Value)
                                        strGenus = Convert.ToString(SQLite.m_DataReader["genus"]).Trim();
                                    if (SQLite.m_DataReader["species"] != System.DBNull.Value)
                                        strSpc = Convert.ToString(SQLite.m_DataReader["species"]).Trim();                                }
                            }
						SQLite.m_DataReader.Close();

                            System.Data.DataRow p_row = SQLite.m_DataSet.Tables["tree_species"].NewRow();
                            p_row["id"] = intId;
                            p_row["fvs_variant"] = this.lstAudit.Items[x].Text;
                            p_row["spcd"] = Convert.ToInt32(this.lstAudit.Items[x].SubItems[1].Text);
                            if (this.lstAudit.Columns.Count == 5)
                            {
                                p_row["fvs_species"] = this.lstAudit.Items[x].SubItems[3].Text;  //fvs tree species 2 character code
                                p_row["fvs_input_spcd"] = Convert.ToString(Convert.ToInt32(this.lstAudit.Items[x].SubItems[2].Text));
                            }
                            if (this.lstAudit.Columns.Count == 4)
                            {
                                p_row["fvs_species"] = this.lstAudit.Items[x].SubItems[3].Text;  //fvs tree species numeric code
                                p_row["fvs_input_spcd"] = this.lstAudit.Items[x].SubItems[2].Text;
                            }

                            p_row["common_name"] = strCommonName;
                            p_row["genus"] = strGenus;
                            p_row["species"] = strSpc;
                            p_row["comments"] = "Added during tree species audit " + DateTime.Now.ToString("MMddyyyy");

                            SQLite.m_DataSet.Tables["tree_species"].Rows.Add(p_row);
                            p_row = null;
                            this.lstAudit.Items[x].Remove();
                            intId++;
                            x--;
                        }


                        this.m_dv.AllowNew = false;
                    }
                }
                catch (Exception err)
                {
                    MessageBox.Show("!!Error!! \n" +
                        "Module - uc_processor_tree_spc:btnAuditAdd_Click() \n" +
                        "Err Msg - " + err.Message,
                        "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Exclamation);

                    this.m_intError = -1;
                }

		}

		private void m_dg_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			switch (e.Button)
			{
				case System.Windows.Forms.MouseButtons.Right :
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
					
					if (this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled=true;
					if (this.m_mnuDataGridPopup.MenuItems[MENU_SELECTALL].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_SELECTALL].Enabled=true;
					if (this.m_mnuDataGridPopup.MenuItems[MENU_MIN].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_MIN].Enabled=true;
					if (this.m_mnuDataGridPopup.MenuItems[MENU_MAX].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_MAX].Enabled=true;
					if (this.m_bDelete==true)
					{
						if (this.m_mnuDataGridPopup.MenuItems[MENU_DELETE].Enabled==false)
						{
						}
					}
					   
					break;
			}
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
							if (this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled==true)
								this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;
							if (this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled==false)
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;

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
		
		public void savevalues()
		{
			int intCurrRow;
		
            this.val_data();
			if (this.m_intError !=0) return;

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
				
				p_dtChanges = SQLite.m_DataSet.Tables["tree_species"].GetChanges();
								
				//check if any inserted rows
				if (p_dtChanges.HasErrors)
				{
					SQLite.m_DataSet.Tables["tree_species"].RejectChanges();
					this.m_intError=-1;
				}
				else
				{
					SQLite.m_DataAdapter.Update(SQLite.m_DataSet.Tables["tree_species"]);
					SQLite.m_Transaction.Commit();
					SQLite.m_DataSet.Tables["tree_species"].AcceptChanges();
					this.InitializeSqliteTransactionCommands();
				}																	
			}
			catch (Exception caught)
			{
				this.m_intError=-1;
				MessageBox.Show(caught.Message);
				SQLite.m_DataSet.Tables["tree_species"].RejectChanges();
				//rollback the transaction to the original records 
				SQLite.m_Transaction.Rollback();				
			}			
			p_dtChanges=null;
			
			this.m_dg.CurrentRowIndex = intCurrRow;		
			this.btnSave.Enabled=false;	            
			return;
		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
			this.savevalues();
		}
		private void InitializeSqliteTransactionCommands()
		{
            SQLite.m_strSQL = "select id, fvs_variant,spcd,fvs_input_spcd,common_name,genus,species,comments from " + this.m_oQueries.m_oFvs.m_strTreeSpcTable + " order by fvs_variant, spcd;";
			//initialize the transaction object with the connection
			SQLite.m_Transaction = SQLite.m_Connection.BeginTransaction();

			SQLite.ConfigureDataAdapterInsertCommand(SQLite.m_Connection,
				SQLite.m_DataAdapter,
				SQLite.m_Transaction,
				SQLite.m_strSQL,
				this.m_oQueries.m_oFvs.m_strTreeSpcTable);

			SQLite.m_strSQL = "select fvs_variant, spcd,fvs_input_spcd,common_name,genus,species,comments from " + m_oQueries.m_oFvs.m_strTreeSpcTable + " order by fvs_variant, spcd;";
			SQLite.ConfigureDataAdapterUpdateCommand(SQLite.m_Connection,
				SQLite.m_DataAdapter,
				SQLite.m_Transaction,
				SQLite.m_strSQL,"select id from " + m_oQueries.m_oFvs.m_strTreeSpcTable,
				m_oQueries.m_oFvs.m_strTreeSpcTable);

			SQLite.m_strSQL = "select fvs_variant, spcd, common_name,fvs_input_spcd,genus,species,comments from " + m_oQueries.m_oFvs.m_strTreeSpcTable + " order by fvs_variant, spcd;";
			SQLite.ConfigureDataAdapterDeleteCommand(SQLite.m_Connection,
				SQLite.m_DataAdapter,
				SQLite.m_Transaction,
				"select id from " + m_oQueries.m_oFvs.m_strTreeSpcTable,
				m_oQueries.m_oFvs.m_strTreeSpcTable);
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
		private void CleanUp()
		{
			try
			{
                if (this.m_dv != null)
                    this.m_dv.Dispose();

                if (this.m_dg != null)
                    this.m_dg.Dispose();

                SQLite.CloseAndDisposeConnection(SQLite.m_Connection, true);
				//if (SQLite.m_DataAdapter != null)
				//	SQLite.m_DataAdapter.Dispose();

				//if (this.m_dv != null)
				//	this.m_dv.Dispose();

				//if (this.m_dg != null)
				//	this.m_dg.Dispose();

				//SQLite.m_DataSet.Clear();
				//SQLite.m_DataSet.Dispose();

				//if (SQLite.m_Connection != null)
				//{
				//	SQLite.m_Connection.Close();
				//	SQLite.m_Connection.Dispose();
				//	SQLite.m_Connection = null;
				//}
				//SQLite.m_Connection = null;
				//this.m_DataSource = null;

			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message);
			}

		}

		private void btnNew_Click(object sender, System.EventArgs e)
		{
			EditForm("New");
		}

		private void EditForm(string p_strAction)
		{
			
			
			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
			frmTemp.MaximizeBox = false;
			frmTemp.BackColor = System.Drawing.SystemColors.Control;
			frmTemp.Text = "FVS: Processor Tree Species (" + p_strAction + ")";

			FIA_Biosum_Manager.uc_processor_tree_spc_edit  p_uc;
			if (p_strAction.Trim().ToUpper()=="NEW")
			{
				p_uc = new uc_processor_tree_spc_edit(SQLite,this.m_oQueries.m_oFvs.m_strTreeSpcTable,this.m_oQueries.m_oFvs.m_strFvsTreeSpcRefTable,"");
			}
			else
			{
				p_uc = new uc_processor_tree_spc_edit(SQLite,this.m_oQueries.m_oFvs.m_strTreeSpcTable,this.m_oQueries.m_oFvs.m_strFvsTreeSpcRefTable,this.m_dg[this.m_intCurrRow-1,this.getGridColumn("fvs_variant")].ToString().Trim());
			}
			
			frmTemp.Controls.Add(p_uc);
			frmTemp.ProcessorTreeSpcEditUserControl=p_uc ;

			frmTemp.Height=0;
			frmTemp.Width=0;
			if (p_uc.Top + p_uc.Height > frmTemp.ClientSize.Height + 2)
			{
				for (int x=1;;x++)
				{
					frmTemp.Height = x;
					if (p_uc.Top + 
						p_uc.Height < 
						frmTemp.ClientSize.Height)
					{
						break;
					}
				}

			}
			if (p_uc.Left + p_uc.Width > frmTemp.ClientSize.Width + 2)
			{
				for (int x=1;;x++)
				{
					frmTemp.Width = x;
					if (p_uc.Left + 
						p_uc.Width < 
						frmTemp.ClientSize.Width)
					{
						break;
					}
				}

			}
			frmTemp.DisposeOfFormWhenClosing = false;
			p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
			frmTemp.Left = 0;
			frmTemp.Top = 0;
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			try
			{
				if (p_strAction.Trim().ToUpper() == "NEW")
				{
					p_uc.strId = Convert.ToString(this.getUniqueId());
				}
				else
				{
					p_uc.strId = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("id")].ToString().Trim();
					p_uc.strSpCd = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("spcd")].ToString().Trim();
					p_uc.strVariant = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("fvs_variant")].ToString().Trim();
					p_uc.strCommonName = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("common_name")].ToString().Trim();
					p_uc.strTreeSpeciesGenus = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("genus")].ToString().Trim();
					p_uc.strTreeSpecies = this.m_dg[this.m_intCurrRow-1,this.getGridColumn("species")].ToString().Trim();
                    int intConvertToSpcd = -1;
                    bool bValidSpcd = int.TryParse(this.m_dg[this.m_intCurrRow - 1, this.getGridColumn("fvs_input_spcd")].ToString().Trim(), 
                        out intConvertToSpcd);
                    if (bValidSpcd == true)
                    {
                        p_uc.intConvertToSpCd = intConvertToSpcd;
                    }
                    else
                    {
                        // Try setting to same as FIA spcd if invalid fvs_input_spcd column
                        int intSpCd = -1;
                        if (int.TryParse(p_uc.strSpCd, out intSpCd) == true)
                        {
                            p_uc.intConvertToSpCd = intSpCd;
                        }
                    }
				}
				System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
				if (result==System.Windows.Forms.DialogResult.OK)
				{
					if (p_strAction.Trim().ToUpper()=="NEW")
					{
						this.m_dv.AllowNew = true;
						System.Data.DataRow p_row =	SQLite.m_DataSet.Tables["tree_species"].NewRow();
						p_row["id"] = Convert.ToInt32(p_uc.strId);
						p_row["fvs_variant"] = p_uc.strVariant;
						p_row["spcd"] = Convert.ToInt32(p_uc.strSpCd);
						p_row["fvs_species"] = p_uc.strFvsSpeciesCode;
						p_row["fvs_common_name"] = p_uc.strFvsCommonName;
						p_row["common_name"] = p_uc.strCommonName;
						p_row["genus"] = p_uc.strTreeSpeciesGenus;
						p_row["species"] = p_uc.strTreeSpecies;
                        if (p_uc.intConvertToSpCd < 1)
						{
							p_row["fvs_input_spcd"] = System.DBNull.Value;
						}
						else
						{
                            p_row["fvs_input_spcd"] = p_uc.intConvertToSpCd;
						}
						SQLite.m_DataSet.Tables["tree_species"].Rows.Add(p_row);
						p_row=null;
						this.m_dv.AllowNew = false;

					}
					else
					{
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("fvs_variant")] = p_uc.strVariant;
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("spcd")] = p_uc.strSpCd;
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("fvs_species")] = p_uc.strFvsSpeciesCode;
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("common_name")] = p_uc.strCommonName;
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("genus")] = p_uc.strTreeSpeciesGenus;
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("species")] = p_uc.strTreeSpecies;
                        if (p_uc.intConvertToSpCd < 1)
						{
							this.m_dg[this.m_intCurrRow-1,this.getGridColumn("fvs_input_spcd")] = System.DBNull.Value;
						}
						else
						{
							this.m_dg[this.m_intCurrRow-1,this.getGridColumn("fvs_input_spcd")] = p_uc.intConvertToSpCd;
						}
						
						this.m_dg[this.m_intCurrRow-1,this.getGridColumn("fvs_common_name")] = p_uc.strFvsCommonName;
						this.m_dg.SetDataBinding(this.m_dv,"");
						this.m_dg.Update();

					}


					if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
					if (this.btnDelete.Enabled==false) this.btnDelete.Enabled=true;
					if (this.btnEdit.Enabled==false) this.btnEdit.Enabled=true;
				}
				frmTemp.Close();
				frmTemp.Dispose();
				frmTemp=null;
			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_processor_tree_spc:EditForm() \n" + 
					"Err Msg - " + e.Message,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
			}



		}
		private int getUniqueId()
		{
            string strUniqueId;
			int intId=0;
			int intId2=0;
			strUniqueId = SQLite.getSingleStringValueFromSQLQuery(SQLite.m_Connection,
				SQLite.m_Transaction,
				"select max(id) as maxid from " + this.m_oQueries.m_oFvs.m_strTreeSpcTable,
				this.m_oQueries.m_oFvs.m_strTreeSpcTable);

			if (strUniqueId != null && strUniqueId.Trim().Length > 0)
				intId = Convert.ToInt32(strUniqueId) + 1;

			if (SQLite.m_DataSet.Tables["tree_species"].Compute("Max(id)", null) != System.DBNull.Value)
			{
				intId2 = Convert.ToInt32(SQLite.m_DataSet.Tables["tree_species"].Compute("Max(id)", null));
			}
			if (intId2 >= intId) 
			{
				intId=intId2 + 1;
			}

			if (intId==0) return 1;
			return intId;

		}

		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			this.EditForm("Edit");
		}
		private void val_data()
		{

			
			this.m_intError=0;
			string strCurVariant="";
			string strCurSpCd="";
			string strCurConvertedSpCd;
			try
			{
				System.Data.DataTable p_dt = SQLite.m_DataSet.Tables["tree_species"];
				for (int x=0;x<=p_dt.Rows.Count-1;x++)
				{
					if (p_dt.Rows[x].RowState != System.Data.DataRowState.Deleted)
					{
						strCurVariant=Convert.ToString(p_dt.Rows[x]["fvs_variant"]);
						strCurSpCd = Convert.ToString(p_dt.Rows[x]["spcd"]);

						//check some basic spc code conversions
						if (p_dt.Rows[x]["fvs_input_spcd"] != 
							System.DBNull.Value &&
							p_dt.Rows[x]["fvs_input_spcd"].ToString().Trim().Length > 0)
						{
							strCurConvertedSpCd = p_dt.Rows[x]["fvs_input_spcd"].ToString().Trim();
							if (strCurConvertedSpCd.Trim() != strCurSpCd.Trim())
							{
								//cannot convert softwood to hardwood
								if (Convert.ToInt32(strCurConvertedSpCd)!= 999)
								{
								}
							}
						}
						else
						{
							strCurConvertedSpCd="";
						}

						//make sure no duplicate variant + spcd + fvs spcd combinations
						for (int y=x+1;y<=p_dt.Rows.Count-1;y++)
						{
							if (p_dt.Rows[y].RowState != System.Data.DataRowState.Deleted)
							{
								//if (strCurFvsSpCd == Convert.ToString(p_dt.Rows[y]["fvs_species"]).Trim())
								//{
									if (strCurVariant == Convert.ToString(p_dt.Rows[y]["fvs_variant"]))
									{
										if (strCurSpCd.Trim() == Convert.ToString(p_dt.Rows[y]["spcd"]).Trim())
										{
											MessageBox.Show("!!Duplicate variant+fia spcd+ fvs spcd combination values found for variant " + strCurVariant + "and, fia spcd " + strCurSpCd + ". Delete one of the records!!","FIA Biosum",
												System.Windows.Forms.MessageBoxButtons.OK,
												System.Windows.Forms.MessageBoxIcon.Exclamation);
											this.m_intError=-1;
											return;

										}
									}
								//}
							}
						}
					}
				}
			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_processor_tree_spc:val_data() \n" + 
					"Err Msg - " + e.Message,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
			}
		}

		private void btnAudit_Click(object sender, System.EventArgs e)
		{
			frmMain.g_sbpInfo.Text = "Running Audit...Stand By";
			if (this.cmbAudit.Text.Trim() == "Assess Data Readiness: Check If Each FIA Tree Spc And FVS Variant Combination Is In The Tree Spc Table")
			{
				this.AuditSpCdCvt();
			}
			//else if (this.cmbAudit.Text.Trim()== "Assess Data Readiness: Check If Each FIA Site Tree Spc, FVS Variant, And FVS Tree Spc Combination Is In The Tree Spc Table")
			//{

			//}
			else if (this.cmbAudit.Text.Trim()== "Assess Data Readiness: Check If Each FIA Tree Spc, FVS Variant, And FVS Tree Spc Combination Is In The Tree Spc Table")
			{
				this.AuditFvsOutSpCdCvt();
			}
            //else if (this.cmbAudit.Text.Trim() == "Assess Data Readiness: Check If Oven Dry Weight And Green Weight Conversion Ratios Exist In The Tree Spc Table")
				
            //{
            //    this.AuditOvenDryGreenWtRatio();
            //}
            //LCB - Hiding; No longer needed with species groups changes
            //else if (this.cmbAudit.Text.Trim() == "Assess Data Readiness: Check If a 2-Character FVS Tree Species Value is Assigned To Each Tree Species Table Record")
            //{
            //    this.AuditAssignFvsSpcAlphaChar();
            //}
			frmMain.g_sbpInfo.Text = "Ready";
			
			
		}

		/// <summary>
		/// Check to make sure that every biosum_cond_id represented in the tree table has a variant assigned to it in plot table
		/// </summary>
		private void AuditSpCdCvt()
		{
			try
			{
				this.lstAudit.Clear();
				btnAuditAdd.Enabled = false;
                btnAuditCheckAll.Enabled = false;
                btnAuditClearAll.Enabled = false;
                lstAudit.CheckBoxes = false;
                this.btnView.Hide();
				//first get unique tree species
				string[,] strValues = new string[1000,2];

                frmMain.g_oFrmMain.ActivateStandByAnimation(this.ParentForm.WindowState,
                this.ParentForm.Left, this.ParentForm.Height, this.ParentForm.Width, this.ParentForm.Top);

				int count=0;
				string strSpCd="";
				string strVar="";
				string strBuild="";
				string strConcat="";
				string strMsg="";
				this.lstAudit.Columns.Add("fvs_variant", 80, HorizontalAlignment.Left);
				this.lstAudit.Columns.Add("spcd", 80, HorizontalAlignment.Left);

				SQLite.m_strSQL = "SELECT t.spcd,p.fvs_variant " + 
									  "FROM " + this.m_oQueries.m_oFIAPlot.m_strTreeTable + " t," + 
												this.m_oQueries.m_oFIAPlot.m_strPlotTable + " p " + 
									  "INNER JOIN " + this.m_oQueries.m_oFIAPlot.m_strCondTable + " c "  + 
									  "ON p.biosum_plot_id=c.biosum_plot_id " +
									  "WHERE t.biosum_cond_id=c.biosum_cond_id AND " + 
					                        "p.fvs_variant IS NOT NULL AND LENGTH(TRIM(p.fvs_variant)) > 0 " + 
									  "ORDER by p.fvs_variant,t.spcd;";


				SQLite.SqlQueryReader(SQLite.m_Connection, SQLite.m_Transaction, SQLite.m_strSQL);
                if (SQLite.m_intError == 0)
                {
                    if (SQLite.m_DataReader.HasRows)
                    {
                        while (SQLite.m_DataReader.Read())
                        {

                            if (SQLite.m_DataReader["spcd"] != System.DBNull.Value)
                            {

                                strVar = "";
                                strSpCd = SQLite.m_DataReader["spcd"].ToString().Trim();
                                if (SQLite.m_DataReader["fvs_variant"] != System.DBNull.Value)
                                    strVar = SQLite.m_DataReader["fvs_variant"].ToString().Trim();

                                strConcat = strSpCd + strVar;
                                if (strBuild.IndexOf(strConcat, 0, strBuild.Length) == -1)
                                {
                                    strValues[count, 0] = strSpCd;
                                    strValues[count, 1] = strVar;
                                    strBuild = strBuild + "'" + strSpCd + strVar + "',";
                                    count++;
                                }
                            }
                        }

                        if (count > 0)
                        {
                            //int y=0;
                            for (int x = 0; x <= count - 1; x++)
                            {


                                System.Data.DataRow[] p_rows = SQLite.m_DataSet.Tables["tree_species"].Select("spcd = " + strValues[x, 0].ToString().Trim() + " and trim(fvs_variant) = '" + strValues[x, 1].ToString().Trim() + "'");
                                if (p_rows != null)
                                {
                                    if (p_rows.Length == 0)
                                    {
                                        this.lstAudit.Items.Add(strValues[x, 1].ToString().Trim());
                                        this.lstAudit.Items[this.lstAudit.Items.Count - 1].SubItems.Add(strValues[x, 0].ToString().Trim());
                                    }
                                }
                                else
                                {
                                    this.lstAudit.Items.Add(strValues[x, 1].ToString().Trim());
                                    this.lstAudit.Items[this.lstAudit.Items.Count - 1].SubItems.Add(strValues[x, 0].ToString().Trim());
                                }
                            }
                        }
                    }
                    SQLite.m_DataReader.Close();

                    if (this.lstAudit.Items.Count == 0)
                    {
                        strMsg = "Audit Passed. \r\n\r\n  Every FIA Tree SpCd + FVS Variant Combination In Your Project Is Found In The Tree Species Table";
                    }
                    else
                    {
                        strMsg = "Audit Failed.\r\n\r\n" + Convert.ToString(this.lstAudit.Items.Count) + " FIA SpCd + FVS Variant Combinations Are NOT Found In The Tree Species Table. ";
                        strMsg += "The SpCd/Variant assignments are listed at the top of the form.\r\n";
                        strMsg += "To update the tree species table in your project with these SpCd/Variant combinations\r\n";
                        strMsg += "check the items to add and click the <Add Checked Items To Tree Species Table> button.";

                        btnAuditAdd.Enabled = true;
                        btnAuditCheckAll.Enabled = true;
                        btnAuditClearAll.Enabled = true;
                        lstAudit.CheckBoxes = true;
                    }
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();
                    MessageBox.Show(strMsg,
                        "Tree Species",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
                else
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();
			}
			catch (Exception e)
			{
                frmMain.g_oFrmMain.DeactivateStandByAnimation();
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_processor_tree_spc:AuditSpCdCvt() \n" + 
					"Err Msg - " + e.Message,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
			}
		}
			
		private void AuditFvsOutSpCdCvt()
		{
			int x;
            int COUNT = 0;

			this.lstAudit.Clear();
			this.btnAuditAdd.Enabled=false;
            this.btnAuditClearAll.Enabled = false;
            this.btnAuditCheckAll.Enabled = false;
            this.lstAudit.CheckBoxes = false;
			//first get unique tree species
			string[,] strValues = new string[1000,2];

			this.lstAudit.Columns.Add("fvs_variant", 80, HorizontalAlignment.Left);
			this.lstAudit.Columns.Add("spcd", 80, HorizontalAlignment.Left);
			this.lstAudit.Columns.Add("fvs_species_numeric_code",150,HorizontalAlignment.Left);
			this.lstAudit.Columns.Add("fvs_species_two_letter_code",150,HorizontalAlignment.Left);
            this.lstAudit.Columns.Add("exist_in_tree_species_table_yn", 180, HorizontalAlignment.Left);
           
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
			string strDbConn = oDataMgr.GetConnectionString(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
				Tables.FVS.DefaultFVSTreeListDbFile);
			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strDbConn))
			{
				conn.Open();
				string strSQL = $@"SELECT COUNT(*) FROM {Tables.FVS.DefaultFVSCutTreeTableName}";
				long lngTreeRecords = oDataMgr.getRecordCount(conn, strSQL, Tables.FVS.DefaultFVSCutTreeTableName);
				if (lngTreeRecords < 1)
				{
					string strMessage = $@"No tree records were found in {Tables.FVS.DefaultFVSCutTreeTableName}. There is nothing to audit!";
					MessageBox.Show(strMessage, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return;
				}
			}

			strDbConn = oDataMgr.GetConnectionString(m_strTempDbFile);
			oDataMgr.OpenConnection(strDbConn);
			// Attach master database so we can access tree and cond tables
			oDataMgr.m_strSQL = $@"ATTACH '{m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Tree)}' as MASTER";
			oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
			// Attach tree_species database so we can access tree_species table
			oDataMgr.m_strSQL = $@"ATTACH '{m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.TreeSpecies)}' as REF_MASTER";
			oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
			// Attach fvs_tree_species database so we can access fvs_tree_species table
			oDataMgr.m_strSQL = $@"ATTACH '{m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.FvsTreeSpecies)}' as BIOSUM_REF";
			oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);

			if (oDataMgr.TableExist(oDataMgr.m_Connection, "treetemp"))
				oDataMgr.SqlNonQuery(oDataMgr.m_Connection, "DROP TABLE treetemp");

			if (oDataMgr.TableExist(oDataMgr.m_Connection, "fvsouttreetemp"))
				oDataMgr.SqlNonQuery(oDataMgr.m_Connection, "DROP TABLE fvsouttreetemp");

			if (oDataMgr.TableExist(oDataMgr.m_Connection, "fvsouttreetemp2"))
				oDataMgr.SqlNonQuery(oDataMgr.m_Connection, "DROP TABLE fvsouttreetemp2");

			if (oDataMgr.TableExist(oDataMgr.m_Connection, "spcd_variant_temp_work_table"))
				oDataMgr.SqlNonQuery(oDataMgr.m_Connection, "DROP TABLE spcd_variant_temp_work_table");

			//let the user know if there are no records in the tree species  table
			if (oDataMgr.getRecordCount(oDataMgr.m_Connection, "select count(*) from " + this.m_oQueries.m_oFvs.m_strTreeSpcTable + ";", m_oQueries.m_oFvs.m_strTreeSpcTable) == 0)
			{
				MessageBox.Show("0 Records In The FVS Tree Species Table", "FIA Biosum",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}
            
            frmMain.g_oFrmMain.ActivateStandByAnimation(this.ParentForm.WindowState,
                this.ParentForm.Left, this.ParentForm.Height, this.ParentForm.Width, this.ParentForm.Top);

			// Attach FVSOUT_TREE_LIST.db
			oDataMgr.m_strSQL = $@"ATTACH '{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSTreeListDbFile}' as FVSOUT_TREE";
			oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);

			List<string> strSqlCommandList = Queries.Processor.AuditFvsOut_SelectIntoUnionOfFVSTreeTablesUsingListArray(
                oDataMgr,
                oDataMgr.m_Connection,
                "fvsouttreetemp2",
                "fvs_tree_id,fvs_variant,fvs_species,biosum_cond_id");
			
            for (x = 0; x <= strSqlCommandList.Count - 1; x++)
            {
                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, strSqlCommandList[x]);
            }
			
			//let the user know if there are no records in the fvs out processor in tree table
			if (oDataMgr.getRecordCount(oDataMgr.m_Connection, "select count(*) from fvsouttreetemp;","fvstree") == 0)
			{
                frmMain.g_oFrmMain.DeactivateStandByAnimation();
				MessageBox.Show("0 Records In The FVS-Out Processor-In Tree Tables","FIA Biosum",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				oDataMgr.CloseConnection(oDataMgr.m_Connection);
				oDataMgr.m_Connection.Dispose();
				return;
			}

 			oDataMgr.m_strSQL = "CREATE TABLE treetemp AS " +
							  "SELECT t.spcd,t.fvs_tree_id,t.biosum_cond_id " +
							  "FROM " + this.m_oQueries.m_oFIAPlot.m_strTreeTable + " t " +
							  "WHERE t.fvs_tree_id IS NOT NULL AND t.biosum_cond_id IS NOT NULL";
			oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);

			oDataMgr.m_strSQL = "CREATE TABLE spcd_variant_temp_work_table AS " +
							  "SELECT DISTINCT t.spcd as treetable_spcd, " + 
							  "f.fvs_species as fvsouttable_spcd," +
							  "f.fvs_variant as fvsouttable_fvs_variant, " +
							  "' ' AS fvs_species_two_letter_code," +
							  "'N' AS exist_in_tree_species_table_yn " +
							  "FROM treetemp t,fvsouttreetemp f " +
							 "WHERE t.fvs_tree_id = CAST(f.fvs_tree_id AS INTEGER) " +
							 "AND t.biosum_cond_id = f.biosum_cond_id";
			oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);

			//add the fvs two letter code 
			oDataMgr.m_strSQL = "UPDATE spcd_variant_temp_work_table " +
				"SET fvs_species_two_letter_code = " +
				"(SELECT fvs_species FROM fvs_tree_species fvs " +
				"WHERE CAST(fvsouttable_spcd AS INTEGER) = fvs.spcd " +
				"AND TRIM(fvsouttable_fvs_variant) = TRIM(fvs.fvs_variant))";
			oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);

            oDataMgr.m_strSQL = "DELETE FROM spcd_variant_temp_work_table as w " +
                           "WHERE EXISTS (SELECT s.spcd,s.fvs_variant,f.fvs_species,s.fvs_input_spcd " +
                           "FROM " + this.m_oQueries.m_oFvs.m_strTreeSpcTable + " s, " +
                           this.m_oQueries.m_oFvs.m_strFvsTreeSpcRefTable + " f " +
                           "WHERE  w.treetable_spcd = s.spcd " +
                           "AND TRIM(w.fvsouttable_fvs_variant)=TRIM(s.fvs_variant) " +
                           "AND (CAST(w.fvsouttable_spcd AS INTEGER)= s.spcd OR CAST(w.fvsouttable_spcd AS INTEGER) = s.fvs_input_spcd) " +
                           "AND (CAST(w.fvsouttable_spcd AS INTEGER)= f.spcd AND TRIM(w.fvsouttable_fvs_variant) = TRIM(f.fvs_variant)) " +
                           "AND f.spcd = s.fvs_input_spcd)";

			oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);

			oDataMgr.m_strSQL = "UPDATE spcd_variant_temp_work_table as w INNER JOIN " +
                                this.m_oQueries.m_oFvs.m_strTreeSpcTable + " t " +
                            "ON w.treetable_spcd = t.spcd AND " +
                               "W.fvsouttable_fvs_variant = t.FVS_VARIANT " +
                            "SET w.exist_in_tree_species_table_yn = 'Y'";
			oDataMgr.m_strSQL = "UPDATE spcd_variant_temp_work_table " +
				"SET exist_in_tree_species_table_yn = " +
				"(SELECT 'Y' FROM tree_species t where treetable_spcd = t.spcd " +
				"AND fvsouttable_fvs_variant = t.FVS_VARIANT)";
			oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);

			try
			{
				oDataMgr.m_strSQL = "SELECT * FROM spcd_variant_temp_work_table";
				oDataMgr.SqlQueryReader(oDataMgr.m_Connection, oDataMgr.m_strSQL);
					if (oDataMgr.m_intError == 0)
					{
						if (oDataMgr.m_DataReader.HasRows)
						{
							while (oDataMgr.m_DataReader.Read())
							{

								if (oDataMgr.m_DataReader["treetable_spcd"] != System.DBNull.Value)
								{
									this.lstAudit.Items.Add(oDataMgr.m_DataReader["fvsouttable_fvs_variant"].ToString().Trim());
									this.lstAudit.Items[this.lstAudit.Items.Count - 1].SubItems.Add(oDataMgr.m_DataReader["treetable_spcd"].ToString().Trim());
									this.lstAudit.Items[this.lstAudit.Items.Count - 1].SubItems.Add(oDataMgr.m_DataReader["fvsouttable_spcd"].ToString().Trim());
									this.lstAudit.Items[this.lstAudit.Items.Count - 1].SubItems.Add(oDataMgr.m_DataReader["fvs_species_two_letter_code"].ToString().Trim());
									this.lstAudit.Items[this.lstAudit.Items.Count - 1].SubItems.Add(oDataMgr.m_DataReader["exist_in_tree_species_table_yn"].ToString().Trim());

									if (oDataMgr.m_DataReader["exist_in_tree_species_table_yn"].ToString().Trim() == "Y") COUNT++;

								}
							}
						}

						oDataMgr.m_DataReader.Close();
						string strMsg = "";
						if (this.lstAudit.Items.Count == 0)
						{
							this.btnView.Hide();
							strMsg = "Audit Passed. \r\n\r\n Every FIA SpCd + FVS Variant + FVS Species Combination is represented in the tree species table";
						}
						else
						{
							strMsg = "Audit Failed.\r\n\r\n" + Convert.ToString(this.lstAudit.Items.Count) + " FIA SpCd + FVS Variant + FVS Species Combinations NOT In The Tree Species Table ";
							this.btnAuditAdd.Text = "Add Checked Items To Tree Species Table";
							//check if all items exist in the tree species table
							if (COUNT == lstAudit.Items.Count)
							{

							}
							else
							{
								lstAudit.CheckBoxes = true;
								btnAuditCheckAll.Enabled = true;
								btnAuditClearAll.Enabled = true;
								btnAuditAdd.Enabled = true;
							}
							btnView.Show();
						}


                    frmMain.g_oFrmMain.DeactivateStandByAnimation();
                    MessageBox.Show(strMsg,
                        "Tree Species",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);

                }
                else
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();

				
			}
			catch (Exception e)
			{
                frmMain.g_oFrmMain.DeactivateStandByAnimation();
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_processor_tree_spc:AuditFvsOutSpCdCvt() \n" + 
					"Err Msg - " + e.Message,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
                if (oDataMgr.m_Connection != null && oDataMgr.m_Connection.State != ConnectionState.Closed)
                {
					oDataMgr.m_Connection.Close();
					oDataMgr.m_Connection.Dispose();
				}
			}
			oDataMgr.m_Connection.Close();
			oDataMgr.m_Connection.Dispose();
		}
        // 09-OCT-2017: No longer used with Tree Species redesign; issue #58
        //private void AuditOvenDryGreenWtRatio()
        //{
        //    this.lstAudit.Clear();
			
        //    btnAuditAdd.Enabled = false;
        //    btnAuditCheckAll.Enabled = false;
        //    btnAuditClearAll.Enabled = false;
        //    btnAuditAdd.Enabled = false;
        //    lstAudit.CheckBoxes = false;
        //    this.btnView.Hide();
        //    //first get unique tree species
        //    string[,] strValues = new string[1000,2];

        //    this.lstAudit.Columns.Add("Id",80,HorizontalAlignment.Left);
        //    this.lstAudit.Columns.Add("fvs_variant", 80, HorizontalAlignment.Left);
        //    this.lstAudit.Columns.Add("spcd", 80, HorizontalAlignment.Left);
        //    this.lstAudit.Columns.Add("fvs_species_numeric_code",150,HorizontalAlignment.Left);
        //    this.lstAudit.Columns.Add("fvs_species_two_letter_code",150,HorizontalAlignment.Left);

        //    ado_data_access oAdo = new ado_data_access();
        //    oAdo.OpenConnection(oAdo.getMDBConnString(this.m_strTempMDBFile,"",""));


			
			

        //    //let the user know if there are no records in the tree species conversion table
        //    if (oAdo.getRecordCount(oAdo.m_OleDbConnection,"select count(*) from " + this.m_oQueries.m_oFvs.m_strTreeSpcTable + ";",this.m_oQueries.m_oFvs.m_strTreeSpcTable) == 0)
        //    {
        //        MessageBox.Show("0 Records In The FVS Tree Species Table","FIA Biosum",
        //            System.Windows.Forms.MessageBoxButtons.OK,
        //            System.Windows.Forms.MessageBoxIcon.Exclamation);
        //        return;
        //    }

        //    if (oAdo.TableExist(oAdo.m_OleDbConnection,"unique_fvs_tree"))
        //        oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE unique_fvs_tree");

        //    if (oAdo.TableExist(oAdo.m_OleDbConnection,"fvsouttreetemp"))
        //        oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE fvsouttreetemp");

        //    if (oAdo.TableExist(oAdo.m_OleDbConnection, "fvsouttreetemp2"))
        //        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE fvsouttreetemp2");

        //    if (oAdo.TableExist(oAdo.m_OleDbConnection,"treespeciestemp"))
        //        oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE treespeciestemp");

        //    if (oAdo.TableExist(oAdo.m_OleDbConnection,"fvsspeciestemp"))
        //        oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE fvsspeciestemp");

        //    if (oAdo.TableExist(oAdo.m_OleDbConnection,"missing_values"))
        //        oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE missing_values");

        //    frmMain.g_oFrmMain.ActivateStandByAnimation(this.ParentForm.WindowState,
        //        this.ParentForm.Left, this.ParentForm.Height, this.ParentForm.Width, this.ParentForm.Top);

        //    List<string> strSqlCommandList;

        //    strSqlCommandList = Queries.Processor.AuditFvsOut_SelectIntoUnionOfFVSTreeTablesUsingListArray(
        //        oAdo,
        //        oAdo.m_OleDbConnection,
        //        "fvsouttreetemp2",
        //        m_oRxPackageItem_Collection,
        //        m_strFVSVariantsArray,
        //        "fvs_tree_id,fvs_variant,fvs_species");

        //    for (int x = 0; x <= strSqlCommandList.Count - 1; x++)
        //    {
        //        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSqlCommandList[x]);
        //    }

        //    oAdo.m_strSQL = "SELECT DISTINCT * INTO fvsouttreetemp FROM fvsouttreetemp2";
        //    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
			

        //    //let the user know if there are no records in the fvs out processor in tree table
        //    if (oAdo.getRecordCount(oAdo.m_OleDbConnection,"select count(*) from fvsouttreetemp;","fvstree") == 0)
        //    {
        //        frmMain.g_oFrmMain.DeactivateStandByAnimation();
        //        MessageBox.Show("0 Records In The FVS-Out Processor-In Tree Tables","FIA Biosum",
        //            System.Windows.Forms.MessageBoxButtons.OK,
        //            System.Windows.Forms.MessageBoxIcon.Exclamation);
        //        return;
        //    }

        //    //let the user know if there are no records in the fvs out processor in tree table
        //    string strTreeListTable = "";
        //    string strTreeListFile="";
		
			

			


			

        //    //join the FIADB biosum tree table with the FVSOut fvstree cutlist table 
        //    oAdo.m_strSQL = "select DISTINCT t.spcd,f.fvs_variant,f.fvs_species " + 
        //                     "INTO unique_fvs_tree " + 
        //                     "FROM fvsouttreetemp  f," + 
        //                          m_oQueries.m_oFIAPlot.m_strTreeTable + " t " + 
        //                     "WHERE f.fvs_tree_id = t.fvs_tree_id;";
        //    oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);

			
        //    //populate a table with records from the tree species that have no values for oven dry weight
        //    //or dry to green conversion columns
        //    oAdo.m_strSQL = "SELECT s.id,s.spcd,s.fvs_variant,s.fvs_species " + 
        //                     "INTO treespeciestemp FROM " + this.m_oQueries.m_oFvs.m_strTreeSpcTable + " s " +
        //                     "WHERE (s.OD_WGT IS NULL OR " + 
        //                            "s.OD_WGT =0 OR s.dry_to_green IS NULL OR " + 
        //                            "s.dry_to_green=0);";
        //    oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);


        //    oAdo.m_strSQL = "SELECT fvs.spcd,fvs.fvs_variant,fvs.fvs_species as fvs_species_two_letter_code " + 
        //                     "INTO fvsspeciestemp " + 
        //                     "FROM " + this.m_oQueries.m_oFvs.m_strFvsTreeSpcRefTable + " fvs ";
        //    oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);

			
        //    //populate a table with records from the tree species table that are found in the fvsout and tree table 
        //    //by species code,fvs variant, and fvs species 2 letter code. These records are missing oven dry and 
        //    //dry to green conversion values
        //    oAdo.m_strSQL = "SELECT DISTINCT s.id, s.spcd, s.fvs_variant,f.fvs_species," + 
        //                                     "s.fvs_species as fvs_species_two_letter_code " + 
        //                     "INTO missing_values " + 
        //                     "FROM treespeciestemp s,unique_fvs_tree f,fvsspeciestemp fvs " + 
        //                     "WHERE f.spcd = s.spcd  AND " + 
        //                     "TRIM(UCASE(f.fvs_variant)) = TRIM(UCASE(s.fvs_variant)) AND " + 
        //                     "TRIM(UCASE(fvs.fvs_species_two_letter_code)) = TRIM(UCASE(s.fvs_species))";
        //    oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);

			
			
        //    oAdo.m_strSQL = "SELECT * FROM missing_values";



	

        //    try
        //    {
        //        oAdo.SqlQueryReader(oAdo.m_OleDbConnection,oAdo.m_strSQL);
        //        if (oAdo.m_intError == 0)
        //        {
        //            if (oAdo.m_OleDbDataReader.HasRows)
        //            {
        //                while (oAdo.m_OleDbDataReader.Read())
        //                {

        //                    if (oAdo.m_OleDbDataReader["spcd"] != System.DBNull.Value)
        //                    {
        //                        this.lstAudit.Items.Add(oAdo.m_OleDbDataReader["id"].ToString().Trim());
        //                        this.lstAudit.Items[this.lstAudit.Items.Count - 1].SubItems.Add(oAdo.m_OleDbDataReader["fvs_variant"].ToString().Trim());
        //                        this.lstAudit.Items[this.lstAudit.Items.Count - 1].SubItems.Add(oAdo.m_OleDbDataReader["spcd"].ToString().Trim());
        //                        this.lstAudit.Items[this.lstAudit.Items.Count - 1].SubItems.Add(oAdo.m_OleDbDataReader["fvs_species"].ToString().Trim());
        //                        this.lstAudit.Items[this.lstAudit.Items.Count - 1].SubItems.Add(oAdo.m_OleDbDataReader["fvs_species_two_letter_code"].ToString().Trim());
        //                    }
        //                }
        //            }

        //            oAdo.m_OleDbDataReader.Close();
        //            string strMsg = "";
        //            if (lstAudit.Items.Count == 0) strMsg = "Audit Passed. \r\n\r\n Every Tree Species Has Wood Weight Ratio Data";
        //            else
        //            {
        //                strMsg = "Audit Failed. \r\n\r\n" + Convert.ToString(this.lstAudit.Items.Count) + " Tree Species Record(s) Are Missing Wood Weight Ratio Data";

        //            }
        //            frmMain.g_oFrmMain.DeactivateStandByAnimation();
        //            MessageBox.Show(strMsg,
        //                           "Tree Species",
        //                           System.Windows.Forms.MessageBoxButtons.OK,
        //                           System.Windows.Forms.MessageBoxIcon.Information);
        //        }
        //        else
        //            frmMain.g_oFrmMain.DeactivateStandByAnimation();
        //    }
        //    catch (Exception e)
        //    {
        //        frmMain.g_oFrmMain.DeactivateStandByAnimation();
        //        MessageBox.Show("!!Error!! \n" + 
        //            "Module - uc_processor_tree_spc:AuditOvenDryGreenWtRatio() \n" + 
        //            "Err Msg - " + e.Message,
        //            "FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
        //            System.Windows.Forms.MessageBoxIcon.Exclamation);

        //        this.m_intError=-1;
        //    }
        //    if (oAdo.TableExist(oAdo.m_OleDbConnection,"unique_fvs_tree"))
        //        oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE unique_fvs_tree");

        //    if (oAdo.TableExist(oAdo.m_OleDbConnection,"fvsouttreetemp"))
        //        oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE fvsouttreetemp");

        //    if (oAdo.TableExist(oAdo.m_OleDbConnection,"treespeciestemp"))
        //        oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE treespeciestemp");

        //    if (oAdo.TableExist(oAdo.m_OleDbConnection,"fvsspeciestemp"))
        //        oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE fvsspeciestemp");

        //    if (oAdo.TableExist(oAdo.m_OleDbConnection,"missing_values"))
        //        oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE missing_values");

        //}

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
		public void SetMenuOptions(string strType)
		{
			this.cmbAudit.Items.Clear();
			if (strType.Trim().ToUpper() == "FVS")
			{
               this.cmbAudit.Items.Add("Assess Data Readiness: Check If Each FIA Tree Spc And FVS Variant Combination Is In The Tree Spc Table");
			   this.cmbAudit.SelectedIndex = 0;
			}
			else
			{
               //this.cmbAudit.Items.Add("Assess Data Readiness: Check If a 2-Character FVS Tree Species Value is Assigned To Each Tree Species Table Record");
			   this.cmbAudit.Items.Add("Assess Data Readiness: Check If Each FIA Tree Spc, FVS Variant, And FVS Tree Spc Combination Is In The Tree Spc Table");
			   //this.cmbAudit.Items.Add("Assess Data Readiness: Check If Oven Dry Weight And Green Weight Conversion Ratios Exist In The Tree Spc Table");
			   this.cmbAudit.SelectedIndex = 0;
			}

		}

		private void lstAudit_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.lstAudit.SelectedItems.Count==0) return;
            if (cmbAudit.Text.Trim()=="Assess Data Readiness: Check If Each FIA Tree Spc, FVS Variant, And FVS Tree Spc Combination Is In The Tree Spc Table")
			{
				CurrencyManager p_cm;
				p_cm = (CurrencyManager)this.BindingContext[this.m_dg.DataSource,this.m_dg.DataMember];
				string strVariant=this.lstAudit.SelectedItems[0].SubItems[0].Text.Trim();
				string strSpCd = this.lstAudit.SelectedItems[0].SubItems[1].Text.Trim();
				for (int x=0;x<=p_cm.Count-1;x++)
				{
					string strVariantCellValue = this.m_dg[x,1].ToString().Trim();
					string strSpCdCellValue = m_dg[x,2].ToString().Trim();
					if (strVariantCellValue == strVariant && strSpCd == strSpCdCellValue)
					{
						m_dg.CurrentRowIndex = x;
						break;
					}

				}
                btnView.Text = "View Affected Trees for Variant:" + strVariant + " SpCd:" + strSpCd;
			}
		}

        private void cmbAudit_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            if (lstAudit.SelectedItems.Count == 0)
            {
                MessageBox.Show("No Rows Are Selected", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            int x;
            string[] strFVSVariantArray = new string[1];
           
            string strFiaSpCd = lstAudit.SelectedItems[0].SubItems[1].Text.Trim();
            string strFvsVariant = lstAudit.SelectedItems[0].SubItems[0].Text.Trim();
            string strFvsSpCd = lstAudit.SelectedItems[0].SubItems[2].Text.Trim();
            strFVSVariantArray[0] = strFvsVariant;

            

            List<string> strSqlCommandList;

            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(this.m_strTempDbFile, "", ""));

            if (oAdo.TableExist(oAdo.m_OleDbConnection, "fvsouttreetemp2"))
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE fvsouttreetemp2");

            strSqlCommandList = Queries.Processor.AuditFvsOut_SelectIntoUnionOfFVSTreeTablesUsingListArray(
                SQLite,
                SQLite.m_Connection,
                "fvsouttreetemp2",
                "fvs_tree_id,fvs_species,biosum_cond_id");

            for (x = 0; x <= strSqlCommandList.Count - 1; x++)
            {
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSqlCommandList[x]);
            }

            oAdo.m_strSQL = "SELECT DISTINCT " + 
                                    "b.fvs_tree_id," + 
                                    "b.fvs_species AS FVS_SpCd," + 
                                    "a.SpCd AS FIA_SpCd," + 
                                    "a.* " + 
                            "FROM " + m_oQueries.m_oFIAPlot.m_strTreeTable + " a," + 
                                    "fvsouttreetemp2 b " + 
                            "WHERE a.fvs_tree_id=b.fvs_tree_id AND " +
                                  "a.biosum_cond_id=b.biosum_cond_id AND " + 
                                  "a.spcd=" + strFiaSpCd + " AND " + 
                                  "TRIM(b.fvs_species)='" + strFvsSpCd + "'";

            frmMain.g_sbpInfo.Text = "Loading Grid...Stand by";

            frmGridView frmGridView1 = new frmGridView();
            frmGridView1.LoadDataSet(
               oAdo.m_OleDbConnection,
               oAdo.m_OleDbConnection.ConnectionString,
               oAdo.m_strSQL, "AuditFiaSpCdAndFvsSpCd");
            frmGridView1.TileGridViews();
            frmMain.g_sbpInfo.Text = "Ready";
            frmGridView1.ShowDialog();
            if (oAdo.TableExist(oAdo.m_OleDbConnection, "fvsouttreetemp2"))
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE fvsouttreetemp2");
            
        }

        private void lstAudit_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (handleCheck)
            {
                if (cmbAudit.Text.Trim() == "Assess Data Readiness: Check If Each FIA Tree Spc, FVS Variant, And FVS Tree Spc Combination Is In The Tree Spc Table")
                {


                    ListViewItem oItem = lstAudit.Items[e.Index] as ListViewItem;
                    if (oItem != null)
                    {
                        if (e.CurrentValue == CheckState.Unchecked && oItem.SubItems[lstAudit.Columns.Count - 1].Text.Trim() == "Y")
                        {
                            MessageBox.Show("Cannot check this record: Combination of fvs variant and tree species already exists in the tree species table", "BIOSUM");
                            e.NewValue = e.CurrentValue;
                        }
                    }


                }
            }
            else handleCheck = true;

        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            // This form is shared by FVS and PROCESSOR; In FVS the audit list only has one item
            if (cmbAudit.Items.Count == 1)
            {
                m_oHelp.XPSFile = Help.DefaultFvsXPSFile;
                m_oHelp.ShowHelp(new string[] { "FVS", "TREE_SPECIES" });
            }
            else
            {
                m_oHelp.XPSFile = Help.DefaultProcessorXPSFile;
                m_oHelp.ShowHelp(new string[] { "PROCESSOR", "TREE_SPECIES" });
            }
        }

	}
	public class TreeSpcAudit_DataGridColoredTextBoxColumn : DataGridTextBoxColumn
	{
		bool m_bEdit=false;
		FIA_Biosum_Manager.uc_processor_tree_spc uc_processor_tree_spc1;
		string m_strLastKey="";
		bool m_bNumericOnly=false;
		

		public TreeSpcAudit_DataGridColoredTextBoxColumn(bool bEdit,bool bNumericOnly,FIA_Biosum_Manager.uc_processor_tree_spc p_uc)
		{
			this.m_bEdit = bEdit;
			this.m_bNumericOnly = bNumericOnly;
			this.uc_processor_tree_spc1 = p_uc;
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
                try
                {
                    // make sure the base class gets called to do the drawing with
                    // the possibly changed brushes
                    base.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight);
                }
                catch
                {
                }
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
				
					if (this.m_bNumericOnly==true)
					{
						if (Char.IsDigit((char)e.KeyValue) || (e.KeyCode== Keys.OemPeriod && this.Format.IndexOf(".",0) >=0 && this.TextBox.Text.IndexOf(".",0) < 0))
						{
							this.m_strLastKey = Convert.ToString(e.KeyValue);
							if (this.uc_processor_tree_spc1.btnSave.Enabled==false) this.uc_processor_tree_spc1.btnSave.Enabled=true;
						}
						else
						{
							if (e.KeyCode == Keys.Back)
							{
								this.m_strLastKey = Convert.ToString(e.KeyValue);
								if (this.uc_processor_tree_spc1.btnSave.Enabled==false) this.uc_processor_tree_spc1.btnSave.Enabled=true;
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
						if (this.uc_processor_tree_spc1.btnSave.Enabled==false) this.uc_processor_tree_spc1.btnSave.Enabled=true;
					}
					
				
			}


			
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
					if (this.uc_processor_tree_spc1.btnSave.Enabled==false) this.uc_processor_tree_spc1.btnSave.Enabled=true;
				}
			}
		}		
		     
	}
}


using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;


namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_gridview.
	/// </summary>
	///
	public class uc_gridview : System.Windows.Forms.UserControl

	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.StatusBar statusBar1;
		private System.Windows.Forms.StatusBarPanel sbMsg;
		private System.Windows.Forms.StatusBarPanel sbQueryRecordCount;
		private System.Windows.Forms.StatusBarPanel sbDisplayedRecordCount;
		private System.Windows.Forms.ToolBar toolBar1;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ToolBarButton btnStructure;
		private System.Windows.Forms.ToolBarButton btnMaxSize;
		private System.Windows.Forms.ToolBarButton btnPrint;
		private System.Windows.Forms.Button btnClose;
		private System.ComponentModel.IContainer components;
	    public  System.Data.DataSet m_ds;
        private System.Data.SQLite.SQLiteDataAdapter m_da;
        private System.Data.SQLite.SQLiteConnection m_conn;
        public int m_intError=0;
		public System.Windows.Forms.DataGrid m_dg;
		public System.Data.DataView m_dv;
		private System.Windows.Forms.ToolBar toolBar2;
		private System.Windows.Forms.ToolBarButton btnFirst;
		private System.Windows.Forms.ToolBarButton btnPrev;
		private System.Windows.Forms.ToolBarButton btnNext;
		private System.Windows.Forms.ToolBarButton btnLast;
		public string m_strError="";
		public int m_intID=0;
		public int m_intIndex=0;
		public System.Windows.Forms.Button btnSQL;
		public int m_intCurrRow=0;
		private int m_intMouseUpCurrRow=-1;
		public System.Windows.Forms.TextBox txtDropDown;
		private string m_strSQL;

		private System.Windows.Forms.ToolBarButton btnID;
		public System.Windows.Forms.ToolBarButton btnSave;
		private int m_intBiosumIdColumn=0;
		private string[] m_strColumnsToEdit;
		private int m_intColumnsToEditCount;
		private string[] m_strRecordKeyColumns;
        private int[] m_intRecordKeyColumns;
        
		private bool m_bClearDataSet=true;
		public System.Windows.Forms.ContextMenu m_mnuDataGridPopup;
		public System.Windows.Forms.ContextMenu m_mnuDataGridTextBoxPopup;
		private const int MENU_FILTERBYVALUE=0;
		private const int MENU_FILTERBYENTEREDVALUE=1;
		private const int MENU_REMOVEFILTER = 2;
		private const int MENU_UNIQUEVALUES = 3;
		private const int MENU_MODIFY = 5;
		private const int MENU_SELECTALL = 7;
		private const int MENU_IDXASC = 8;
		private const int MENU_IDXDESC = 9;
		private const int MENU_REMOVEIDX = 10;
		private const int MENU_MAX = 12;
		private const int MENU_MIN = 13;
		private const int MENU_AVG = 14;
		private const int MENU_SUM = 15;
		private const int MENU_COUNTBYVALUE=16;

		private int m_intPopupColumn=0;
		private string m_strDeletedPlotList="";
		private int m_intDeletedPlotCount=0;
		private string m_strColumnFilterList="";
		private string m_strColumnSortList="";
		private FIA_Biosum_Manager.frmTherm m_frmTherm;
		private System.Data.DataTable m_dtTableSchema;
		private int m_intTableSchemaColumnRow;
        private bool m_bAdapterDisposed = true;
		private FIA_Biosum_Manager.frmGridView _frmGridView=null;

		public uc_gridview()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
            this.sbQueryRecordCount.Text = "";
            this.sbDisplayedRecordCount.Text = "";
            this.sbMsg.Text = "";
			// TODO: Add any initialization after the InitializeComponent call

		}
		public uc_gridview(string strConn, string strSQL,string strDataSetName)
		{
			InitializeComponent();
            LoadGridView(strConn, strSQL, strDataSetName);
		}
		
		public uc_gridview(System.Data.SQLite.SQLiteConnection p_conn,
			string strSQL,string strDataSetName)
		{
			InitializeComponent();
			this.InitializePopup();
			string strColumnName="";
			this.m_dg.MouseWheel+=new MouseEventHandler(m_dg_MouseWheel);
			this.m_ds = new DataSet();
			this.m_da = new System.Data.SQLite.SQLiteDataAdapter();

			this.m_dg.Left = 5;
			this.toolBar1.Left = 5;
			this.btnClose.Top = this.groupBox1.Top + 10;

			this.m_dg.Width = this.groupBox1.Width - 10;
			this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - this.groupBox1.Left - 10;
			this.toolBar2.Top = this.statusBar1.Top - this.toolBar2.Height  - 2;
			this.toolBar2.Left = (int)(this.groupBox1.Width * .50) - (int)(this.toolBar2.Width * .50);
			this.m_dg.Height =  this.toolBar2.Top - this.m_dg.Top;

			this.m_da.SelectCommand = new System.Data.SQLite.SQLiteCommand(strSQL, p_conn);
			
			try 
			{
				this.m_da.Fill(this.m_ds, strDataSetName);
					
				this.m_dv = new DataView(this.m_ds.Tables[strDataSetName]);
				
				this.m_dv.AllowNew = false;       //cannot append new records
				this.m_dv.AllowDelete = false;    //cannot delete records
				this.m_dv.AllowEdit = false;
				this.m_dg.CaptionText = strDataSetName;
				this.m_dg.BackgroundColor=frmMain.g_oGridViewBackgroundColor;
				/***********************************************************************************
				 **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
				 ***********************************************************************************/
				gridview_DataGridColoredTextBoxColumn aColumnTextColumn ;

				/***************************************************************
				 **custom define the grid style
				 ***************************************************************/
				DataGridTableStyle tableStyle = new DataGridTableStyle();

				/***********************************************************************
				 **map the data grid table style to the scenario rx intensity dataset
				 ***********************************************************************/
				tableStyle.MappingName = strDataSetName;
				tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
				tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
				tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
				tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
   
				/******************************************************************************
				 **since the dataset has things like field name and number of columns,
				 **we will use those to create new columnstyles for the columns in our grid
				 ******************************************************************************/
				//get the number of columns from the scenario_rx_intensity data set
				int numCols = this.m_ds.Tables[strDataSetName].Columns.Count;
                
				/************************************************
				 **loop through all the columns in the dataset	
				 ************************************************/
				for(int i = 0; i < numCols; ++i)
				{
					strColumnName = this.m_ds.Tables[strDataSetName].Columns[i].ColumnName;
					
					/******************************************************************
					 **create a new instance of the DataGridColoredTextBoxColumn class
					 ******************************************************************/
					aColumnTextColumn = new gridview_DataGridColoredTextBoxColumn(false,this);

					/***********************************
					 **all columns are read-only except
					 **the edit columns
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

				this.m_dg.DataSource = this.m_dv;  

				this.m_dg.Expand(-1);
			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message,"Table",MessageBoxButtons.OK,MessageBoxIcon.Error);
				this.m_intError=-1;
				this.m_ds.Clear();
				this.m_ds = null;
                if (p_conn != null)
                {
					p_conn.Close();
					this.m_da.Dispose();
					this.m_da = null;
                }
				return;
			}

			this.sbMsg.Text= strDataSetName + " (Read Only)";
			if (this.m_ds.Tables[strDataSetName].Rows.Count == 0)
			{
				this.sbQueryRecordCount.Text= "0/0";
				this.sbDisplayedRecordCount.Text="0";
			}
			else
			{
				this.m_intCurrRow = 1;
				this.sbQueryRecordCount.Text= "1/"+ this.m_ds.Tables[strDataSetName].Rows.Count.ToString().Trim();

				this.sbDisplayedRecordCount.Text = Convert.ToString(this.m_dg.BindingContext[this.m_dg.DataSource,this.m_dg.DataMember].Count);
			}
			
			this.sbQueryRecordCount.Width  =
				(int)this.CreateGraphics().MeasureString("9999999/9999999",this.statusBar1.Font).Width;
			this.sbMsg.Width = (int)(this.groupBox1.Width * .50) -  (int)(this.sbQueryRecordCount.Width * .50);
			this.sbDisplayedRecordCount.Width =(int)(this.groupBox1.Width * .50) - (int)(this.sbQueryRecordCount.Width * .50);
			this.sbDisplayedRecordCount.Alignment = System.Windows.Forms.HorizontalAlignment.Right;
			this.sbQueryRecordCount.Alignment = System.Windows.Forms.HorizontalAlignment.Center;

			//event handler to keep track of current row and cell movement
			this.m_dg.CurrentCellChanged += new
				System.EventHandler(this.m_dg_CurrentCellChanged);

			this.m_strSQL = strSQL;
			if (this.m_strSQL.Trim().Length > 0) this.btnSQL.Enabled=true;

			/******************************************************************
			 **if biosum_cond_id or biosum_plot_id is part of the dataset then
			 **enable the id toolbar button
			 ******************************************************************/
			int x=0;
			for (x=0; x<=this.m_ds.Tables[strDataSetName].Columns.Count-1;x++)
			{
				if (this.m_ds.Tables[strDataSetName].Columns[x].ColumnName.Trim().ToUpper() == "BIOSUM_COND_ID")
				{
					this.m_intBiosumIdColumn = x;
					break;
				}
				if (this.m_ds.Tables[strDataSetName].Columns[x].ColumnName.Trim().ToUpper() == "BIOSUM_PLOT_ID")
				{
					this.m_intBiosumIdColumn = x;
					break;
				}
			}
			if (x <= this.m_ds.Tables[strDataSetName].Columns.Count-1)
			{
				this.toolBar1.Buttons[3].Enabled=true;
			}
			else
			{
				this.toolBar1.Buttons[3].Enabled=false;
			}
		}

		/*************************************************
		 ** Edit Version
		 *************************************************/
		public uc_gridview(string strConn, 
			               string strSQL,
			               string strTableName,
			               string[] strColumnsToEdit,
			               int intColumnsToEditCount,
			               string[] strRecordKeyColumns)
		{
			int x,y;
			InitializeComponent();
			this.InitializePopup();
			this.m_dg.MouseWheel+=new MouseEventHandler(m_dg_MouseWheel);
			string strColumnName="";
            this.m_ds = new DataSet();
            this.m_conn = new System.Data.SQLite.SQLiteConnection();
            this.m_da = new System.Data.SQLite.SQLiteDataAdapter();

            this.m_dg.Left = 5;
			this.toolBar1.Left = 5;
			this.btnClose.Top = this.groupBox1.Top + 10;

			this.m_dg.Width = this.groupBox1.Width - 10;
			this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - this.groupBox1.Left - 10;
			this.toolBar2.Top = this.statusBar1.Top - this.toolBar2.Height  - 2;
			this.toolBar2.Left = (int)(this.groupBox1.Width * .50) - (int)(this.toolBar2.Width * .50);
			this.m_dg.Height =  this.toolBar2.Top - this.m_dg.Top;

			this.m_intColumnsToEditCount = intColumnsToEditCount;
			this.m_strColumnsToEdit = new string[intColumnsToEditCount];
			this.m_strColumnsToEdit = strColumnsToEdit;
			this.m_strRecordKeyColumns = new string[strRecordKeyColumns.Length];
			this.m_strRecordKeyColumns = strRecordKeyColumns;
            this.m_intRecordKeyColumns = new int[strRecordKeyColumns.Length];

            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            dataMgr.OpenConnection(strConn, ref this.m_conn);
            if (dataMgr.m_intError != 0)
            {
                this.m_intError = dataMgr.m_intError;
                dataMgr = null;
                return;
            }
            //get the table schema of the result of the sql
            this.m_dtTableSchema = dataMgr.getTableSchema(this.m_conn, strSQL);
            this.m_da.SelectCommand = new System.Data.SQLite.SQLiteCommand(strSQL, this.m_conn);

			try 
			{
                this.m_da.Fill(this.m_ds, strTableName);
                //for (int x=0; x<=this.m_ds.Tables.Count-1;x++) MessageBox.Show(this.m_ds.Tables[x].TableName);
                this.m_dv = new DataView(this.m_ds.Tables[strTableName]);
				
				this.m_dv.AllowNew = false;       //cannot append new records
				this.m_dv.AllowDelete = false;    //cannot delete records
				
				this.m_dg.CaptionText = strTableName;
				m_dg.BackgroundColor=frmMain.g_oGridViewBackgroundColor;
				
				/***********************************************************************************
				 **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
				 ***********************************************************************************/
				gridview_DataGridColoredTextBoxColumn aColumnTextColumn ;

				/***************************************************************
				 **custom define the grid style
				 ***************************************************************/
				DataGridTableStyle tableStyle = new DataGridTableStyle();
               
				/***********************************************************************
				 **map the data grid table style to the scenario rx intensity dataset
				 ***********************************************************************/
				tableStyle.MappingName = strTableName;
				tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
				tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
				tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
				tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
   
				/******************************************************************************
				 **since the dataset has things like field name and number of columns,
				 **we will use those to create new columnstyles for the columns in our grid
				 ******************************************************************************/
				//get the number of columns from the scenario_rx_intensity data set
				int numCols = this.m_ds.Tables[strTableName].Columns.Count;
                
                int intRecordKeyCount=0;    
				/************************************************
				 **loop through all the columns in the dataset	
				 ************************************************/
				for(int i = 0; i < numCols; ++i)
				{
					strColumnName = this.m_ds.Tables[strTableName].Columns[i].ColumnName;

                    if (intRecordKeyCount < m_strRecordKeyColumns.Length)
                    {
                        for (y = 0; y <= m_strRecordKeyColumns.Length - 1; y++)
                        {
                            if (strColumnName.Trim().ToUpper() == m_strRecordKeyColumns[y].Trim().ToUpper())
                            {
                                m_intRecordKeyColumns[intRecordKeyCount] = i;
                                intRecordKeyCount++;
                            }
                        }
                    }
					/*****************************************************************
					 **see if this column is one that is to be edited
					 *****************************************************************/
					for (x=0; x <= intColumnsToEditCount-1; x++)
					{
						if (strColumnName.Trim().ToUpper() == 
							strColumnsToEdit[x].Trim().ToUpper()) break;
					}
					/*****************************************************************
					 **if the column is not to be edited then set it to read only
					 *****************************************************************/
					if (x > intColumnsToEditCount-1)
					{
						/******************************************************************
						 **create a new instance of the DataGridColoredTextBoxColumn class
						 ******************************************************************/
						aColumnTextColumn = new gridview_DataGridColoredTextBoxColumn(false,this);

						/***********************************
						 **all columns are read-only except
						 **the edit columns
						 ***********************************/
						aColumnTextColumn.ReadOnly=true;
					}
					else
					{
						/******************************************************************
						 **create a new instance of the DataGridColoredTextBoxColumn class
						 ******************************************************************/
						aColumnTextColumn = new gridview_DataGridColoredTextBoxColumn(true,this);
						
						aColumnTextColumn.Format="#0.00";
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
				if (frmMain.g_oGridViewFont != null) this.m_dg.Font = frmMain.g_oGridViewFont;

				this.m_dg.TableStyles.Clear();
				this.m_dg.TableStyles.Add(tableStyle);
				
				this.m_dg.DataSource = this.m_dv;
			

				this.m_dg.Expand(-1);
			}
			catch (Exception e)
			{
				MessageBox.Show(e.Message,"Table",MessageBoxButtons.OK,MessageBoxIcon.Error);
				this.m_intError=-1;
				this.m_ds.Clear();
				this.m_ds = null;
                if (m_conn != null)
                {
                    m_conn.Close();
                }
				return;
			}
			
			this.sbMsg.Text= strTableName + " (Edit/No Append/No Delete)";
			if (this.m_ds.Tables[strTableName].Rows.Count == 0)
			{
				this.sbQueryRecordCount.Text= "0/0";
				this.sbDisplayedRecordCount.Text="0";
			}
			else
			{
				this.m_intCurrRow = 1;
				this.sbQueryRecordCount.Text= "1/"+ this.m_ds.Tables[strTableName].Rows.Count.ToString().Trim();
				
				this.sbDisplayedRecordCount.Text = Convert.ToString(this.m_dg.BindingContext[this.m_dg.DataSource,this.m_dg.DataMember].Count);
			}
			
			this.sbQueryRecordCount.Width  =
				(int)this.CreateGraphics().MeasureString("9999999/9999999",this.statusBar1.Font).Width;
			this.sbMsg.Width = (int)(this.groupBox1.Width * .50) -  (int)(this.sbQueryRecordCount.Width * .50);
			this.sbDisplayedRecordCount.Width =(int)(this.groupBox1.Width * .50) - (int)(this.sbQueryRecordCount.Width * .50);
			this.sbDisplayedRecordCount.Alignment = System.Windows.Forms.HorizontalAlignment.Right;
			this.sbQueryRecordCount.Alignment = System.Windows.Forms.HorizontalAlignment.Center;

            //event handler to keep track of current row and cell movement
            this.m_dg.CurrentCellChanged += new
                System.EventHandler(this.m_dg_CurrentCellChanged);

			this.m_strSQL = strSQL;
			if (this.m_strSQL.Trim().Length > 0) this.btnSQL.Enabled=true;

			/******************************************************************
			 **if biosum_cond_id or biosum_plot_id is part of the dataset then
			 **enable the id toolbar button
			 ******************************************************************/
			for (x=0; x<=this.m_ds.Tables[strTableName].Columns.Count-1;x++)
			{
				if (this.m_ds.Tables[strTableName].Columns[x].ColumnName.Trim().ToUpper() == "BIOSUM_COND_ID")
				{
					this.m_intBiosumIdColumn = x;
					break;
				}
				if (this.m_ds.Tables[strTableName].Columns[x].ColumnName.Trim().ToUpper() == "BIOSUM_PLOT_ID")
				{
					this.m_intBiosumIdColumn = x;
					break;
				}
			}
			if (x <= this.m_ds.Tables[strTableName].Columns.Count-1)
			{
				this.toolBar1.Buttons[3].Enabled=true;
			}
			else
			{
				this.toolBar1.Buttons[3].Enabled=false;
			}
			
		}

		public string DataSetName
		{
			get
			{
				return this.m_dg.CaptionText;
			}
		}
		public int CloseGrid
		{
			get
			{
				this.CloseGridView();
				return 0;
			}
		}
		public System.Data.DataSet getDataSet
		{
			get
			{
				return this.m_ds;
			}
		}
		public uc_gridview getGridViewObject
		{
			get
			{
				return (this);
			}
		}
		public int MultiPaneDisplayOrder
		{
			get
			{
				return 0;
			}
			set
			{
				MultiPaneDisplayOrder = value;
			}
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_gridview));
            this.m_dg = new System.Windows.Forms.DataGrid();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtDropDown = new System.Windows.Forms.TextBox();
            this.btnSQL = new System.Windows.Forms.Button();
            this.toolBar2 = new System.Windows.Forms.ToolBar();
            this.btnFirst = new System.Windows.Forms.ToolBarButton();
            this.btnPrev = new System.Windows.Forms.ToolBarButton();
            this.btnNext = new System.Windows.Forms.ToolBarButton();
            this.btnLast = new System.Windows.Forms.ToolBarButton();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.btnClose = new System.Windows.Forms.Button();
            this.toolBar1 = new System.Windows.Forms.ToolBar();
            this.btnStructure = new System.Windows.Forms.ToolBarButton();
            this.btnMaxSize = new System.Windows.Forms.ToolBarButton();
            this.btnPrint = new System.Windows.Forms.ToolBarButton();
            this.btnID = new System.Windows.Forms.ToolBarButton();
            this.btnSave = new System.Windows.Forms.ToolBarButton();
            this.statusBar1 = new System.Windows.Forms.StatusBar();
            this.sbMsg = new System.Windows.Forms.StatusBarPanel();
            this.sbQueryRecordCount = new System.Windows.Forms.StatusBarPanel();
            this.sbDisplayedRecordCount = new System.Windows.Forms.StatusBarPanel();
            ((System.ComponentModel.ISupportInitialize)(this.m_dg)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sbMsg)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sbQueryRecordCount)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sbDisplayedRecordCount)).BeginInit();
            this.SuspendLayout();
            // 
            // m_dg
            // 
            this.m_dg.DataMember = "";
            this.m_dg.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.m_dg.Location = new System.Drawing.Point(16, 56);
            this.m_dg.Name = "m_dg";
            this.m_dg.Size = new System.Drawing.Size(632, 320);
            this.m_dg.TabIndex = 0;
            this.m_dg.MouseUp += new System.Windows.Forms.MouseEventHandler(this.m_dg_MouseUp);
            this.m_dg.MouseDown += new System.Windows.Forms.MouseEventHandler(this.m_dg_MouseDown);
            this.m_dg.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.m_dg_KeyPress);
            this.m_dg.Navigate += new System.Windows.Forms.NavigateEventHandler(this.dataGrid1_Navigate);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtDropDown);
            this.groupBox1.Controls.Add(this.btnSQL);
            this.groupBox1.Controls.Add(this.toolBar2);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.toolBar1);
            this.groupBox1.Controls.Add(this.statusBar1);
            this.groupBox1.Controls.Add(this.m_dg);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(664, 456);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // txtDropDown
            // 
            this.txtDropDown.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.txtDropDown.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDropDown.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtDropDown.HideSelection = false;
            this.txtDropDown.Location = new System.Drawing.Point(624, 392);
            this.txtDropDown.Multiline = true;
            this.txtDropDown.Name = "txtDropDown";
            this.txtDropDown.ReadOnly = true;
            this.txtDropDown.Size = new System.Drawing.Size(24, 24);
            this.txtDropDown.TabIndex = 41;
            this.txtDropDown.Visible = false;
            // 
            // btnSQL
            // 
            this.btnSQL.BackColor = System.Drawing.SystemColors.Control;
            this.btnSQL.Enabled = false;
            this.btnSQL.Image = ((System.Drawing.Image)(resources.GetObject("btnSQL.Image")));
            this.btnSQL.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnSQL.Location = new System.Drawing.Point(224, 32);
            this.btnSQL.Name = "btnSQL";
            this.btnSQL.Size = new System.Drawing.Size(88, 24);
            this.btnSQL.TabIndex = 40;
            this.btnSQL.Text = "SQL";
            this.btnSQL.UseVisualStyleBackColor = false;
            this.btnSQL.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnSQL_MouseDown);
            this.btnSQL.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btnSQL_MouseUp);
            // 
            // toolBar2
            // 
            this.toolBar2.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
            this.btnFirst,
            this.btnPrev,
            this.btnNext,
            this.btnLast});
            this.toolBar2.Divider = false;
            this.toolBar2.Dock = System.Windows.Forms.DockStyle.None;
            this.toolBar2.DropDownArrows = true;
            this.toolBar2.ImageList = this.imageList1;
            this.toolBar2.Location = new System.Drawing.Point(296, 384);
            this.toolBar2.Name = "toolBar2";
            this.toolBar2.ShowToolTips = true;
            this.toolBar2.Size = new System.Drawing.Size(96, 26);
            this.toolBar2.TabIndex = 39;
            this.toolBar2.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.toolBar2_ButtonClick);
            // 
            // btnFirst
            // 
            this.btnFirst.ImageIndex = 0;
            this.btnFirst.Name = "btnFirst";
            this.btnFirst.ToolTipText = "First Record";
            // 
            // btnPrev
            // 
            this.btnPrev.ImageIndex = 1;
            this.btnPrev.Name = "btnPrev";
            this.btnPrev.ToolTipText = "Previous Record";
            // 
            // btnNext
            // 
            this.btnNext.ImageIndex = 2;
            this.btnNext.Name = "btnNext";
            this.btnNext.ToolTipText = "Next Record";
            // 
            // btnLast
            // 
            this.btnLast.ImageIndex = 3;
            this.btnLast.Name = "btnLast";
            this.btnLast.ToolTipText = "Last Record";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "");
            this.imageList1.Images.SetKeyName(1, "");
            this.imageList1.Images.SetKeyName(2, "");
            this.imageList1.Images.SetKeyName(3, "");
            this.imageList1.Images.SetKeyName(4, "");
            this.imageList1.Images.SetKeyName(5, "");
            this.imageList1.Images.SetKeyName(6, "");
            this.imageList1.Images.SetKeyName(7, "");
            this.imageList1.Images.SetKeyName(8, "");
            this.imageList1.Images.SetKeyName(9, "");
            // 
            // btnClose
            // 
            this.btnClose.ImageIndex = 7;
            this.btnClose.ImageList = this.imageList1;
            this.btnClose.Location = new System.Drawing.Point(624, 16);
            this.btnClose.Name = "btnClose";
            this.btnClose.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnClose.Size = new System.Drawing.Size(32, 32);
            this.btnClose.TabIndex = 38;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // toolBar1
            // 
            this.toolBar1.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
            this.btnStructure,
            this.btnMaxSize,
            this.btnPrint,
            this.btnID,
            this.btnSave});
            this.toolBar1.Divider = false;
            this.toolBar1.Dock = System.Windows.Forms.DockStyle.None;
            this.toolBar1.DropDownArrows = true;
            this.toolBar1.ImageList = this.imageList1;
            this.toolBar1.Location = new System.Drawing.Point(24, 30);
            this.toolBar1.Name = "toolBar1";
            this.toolBar1.ShowToolTips = true;
            this.toolBar1.Size = new System.Drawing.Size(128, 26);
            this.toolBar1.TabIndex = 37;
            this.toolBar1.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.toolBar1_ButtonClick);
            // 
            // btnStructure
            // 
            this.btnStructure.ImageIndex = 4;
            this.btnStructure.Name = "btnStructure";
            this.btnStructure.ToolTipText = "Structure";
            // 
            // btnMaxSize
            // 
            this.btnMaxSize.ImageIndex = 5;
            this.btnMaxSize.Name = "btnMaxSize";
            this.btnMaxSize.ToolTipText = "Maximum Size";
            // 
            // btnPrint
            // 
            this.btnPrint.ImageIndex = 6;
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.ToolTipText = "Print Report";
            // 
            // btnID
            // 
            this.btnID.ImageIndex = 8;
            this.btnID.Name = "btnID";
            this.btnID.ToolTipText = "biosum_cond_id or biosum_plot_id";
            // 
            // btnSave
            // 
            this.btnSave.Enabled = false;
            this.btnSave.ImageIndex = 9;
            this.btnSave.Name = "btnSave";
            this.btnSave.ToolTipText = "Save ";
            // 
            // statusBar1
            // 
            this.statusBar1.Location = new System.Drawing.Point(3, 429);
            this.statusBar1.Name = "statusBar1";
            this.statusBar1.Panels.AddRange(new System.Windows.Forms.StatusBarPanel[] {
            this.sbMsg,
            this.sbQueryRecordCount,
            this.sbDisplayedRecordCount});
            this.statusBar1.ShowPanels = true;
            this.statusBar1.Size = new System.Drawing.Size(658, 24);
            this.statusBar1.SizingGrip = false;
            this.statusBar1.TabIndex = 36;
            // 
            // sbMsg
            // 
            this.sbMsg.Name = "sbMsg";
            this.sbMsg.Text = "sbMsg";
            // 
            // sbQueryRecordCount
            // 
            this.sbQueryRecordCount.Name = "sbQueryRecordCount";
            this.sbQueryRecordCount.Text = "sbQueryRecordCount";
            // 
            // sbDisplayedRecordCount
            // 
            this.sbDisplayedRecordCount.Name = "sbDisplayedRecordCount";
            this.sbDisplayedRecordCount.Text = "sbDisplayedRecordCount";
            // 
            // uc_gridview
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_gridview";
            this.Size = new System.Drawing.Size(664, 456);
            this.Resize += new System.EventHandler(this.uc_gridview_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.m_dg)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sbMsg)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sbQueryRecordCount)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sbDisplayedRecordCount)).EndInit();
            this.ResumeLayout(false);

		}
		#endregion

		private void dataGrid1_Navigate(object sender, System.Windows.Forms.NavigateEventArgs ne)
		{
		
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.m_dg.Width = this.groupBox1.Width - 10 ;

				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - this.groupBox1.Left - 10;
				this.toolBar2.Top = this.statusBar1.Top - this.toolBar2.Height  - 2;
				this.toolBar2.Left = (int)(this.groupBox1.Width * .50) - (int)(this.toolBar2.Width * .50);
				this.m_dg.Height =  this.toolBar2.Top - this.m_dg.Top;
				this.sbQueryRecordCount.Width  =
					(int)this.CreateGraphics().MeasureString("9999999/9999999",this.statusBar1.Font).Width;
				this.sbMsg.Width = (int)(this.groupBox1.Width * .50) -  (int)(this.sbQueryRecordCount.Width * .50);
				this.sbDisplayedRecordCount.Width =(int)(this.groupBox1.Width * .50) - (int)(this.sbQueryRecordCount.Width * .50);
			}
			catch
			{
			}

		}
		private void m_dg_CurrentCellChanged(object sender, 
			System.EventArgs e)
		{
			
			if (m_dg.CurrentRowIndex >=0) m_dg.Select(m_dg.CurrentRowIndex);
			if (this.m_intCurrRow > 0)
			{
				if (this.m_dg.CurrentRowIndex != this.m_intCurrRow - 1)
				{
					this.m_intCurrRow = this.m_dg.CurrentRowIndex + 1;
					this.sbQueryRecordCount.Text= this.m_intCurrRow.ToString().Trim() + "/"+ this.m_ds.Tables[DataSetName].Rows.Count.ToString().Trim();
				}
			}
		}

		private void toolBar2_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
            if (m_ds == null) return;
			if (m_dg.CurrentRowIndex >= 0) m_dg.UnSelect(m_dg.CurrentRowIndex);
			switch (this.toolBar2.Buttons.IndexOf(e.Button))
			{
				case 0:
					this.m_dg.CurrentRowIndex = 0;
					break;
				case 1:
					if (this.m_dg.CurrentRowIndex > 0) 
							this.m_dg.CurrentRowIndex = this.m_dg.CurrentRowIndex - 1;

					break;
				case 2:
					if (this.m_dg.CurrentRowIndex < this.m_ds.Tables[DataSetName].Rows.Count-1) 
						this.m_dg.CurrentRowIndex = this.m_dg.CurrentRowIndex + 1;

					break;
				case 3:
					this.m_dg.CurrentRowIndex = this.m_ds.Tables[DataSetName].Rows.Count - 1;
					break;
			}
			if (m_dg.CurrentRowIndex >=0) m_dg.Select(m_dg.CurrentRowIndex);
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.CloseGridView();
		}
		public void CloseGridView()
		{
			if (this.Visible == true) this.Visible=false;
			if (this.m_bClearDataSet == true)
			{
				if (m_da != null) 
				{
					this.m_da.Dispose();
					while (this.m_bAdapterDisposed==false)
						System.Threading.Thread.Sleep(1000);
				}
                if (m_da != null)
                {
                    this.m_da.Dispose();
                    // For now, let's just dispose of it and hope it works properly
                }
				if (m_da != null) this.m_da = null;
                if (m_da != null) this.m_da = null;
                if (m_ds.Tables[m_dg.CaptionText] != null)  m_ds.Tables[m_dg.CaptionText].Clear();
				if (m_ds.Tables[m_dg.CaptionText] !=null) m_ds.Tables[m_dg.CaptionText].Dispose();
				if (m_ds != null) this.m_ds.Clear();
				if (m_conn != null)
				{
					this.m_conn.Close();
					while (m_conn.State != System.Data.ConnectionState.Closed)
						System.Threading.Thread.Sleep(1000);
				}
				if (m_conn != null) m_conn.Dispose();
				if (m_conn != null) this.m_conn = null;
                if (m_conn != null)
                {
                    this.m_conn.Close();
                }
                if (m_conn != null) this.m_conn.Dispose();

			}
			if (this.DataSetName.Trim().Length > 0) this.ReferenceGridViewForm.RemoveGridViewMenuItem(this.DataSetName);
			if (this.DataSetName.Trim().Length > 0) this.ReferenceGridViewForm.RemoveGridViewCollectionItem(this.DataSetName);
			this.Dispose();
		}
		private void m_da_Disposed(object sender, EventArgs e)
		{
				this.m_bAdapterDisposed=true;
		}

		private void btnSQL_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			string strTemp = this.m_strSQL;
			int lines = (int) (strTemp.Length / 30);  //40 characters per line
			if (lines == 0) lines = 1;
			this.txtDropDown.Text = strTemp;
			this.txtDropDown.Top = this.btnSQL.Top + this.btnSQL.Height - 3;
			int textWidth = (int)this.CreateGraphics().MeasureString(strTemp, this.txtDropDown.Font).Width;
			int textHeight = (int)this.CreateGraphics().MeasureString(strTemp, this.txtDropDown.Font).Height;
			this.txtDropDown.Left = this.btnSQL.Left;
			this.txtDropDown.Width = (int) ((textWidth / lines) * 1.5);
			this.txtDropDown.Height = (int)Math.Round((textHeight * lines) * 1.5,0) ;
			this.txtDropDown.BringToFront();
			this.txtDropDown.Visible = true;

		}

		private void btnSQL_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			this.txtDropDown.Visible=false;
		}

		private void uc_gridview_Resize(object sender, System.EventArgs e)
		{

		}

		private void toolBar1_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
            if (m_dg == null || this.m_ds==null) return;
			switch (this.toolBar1.Buttons.IndexOf(e.Button))
			{
				case 0:    //view and print structure
					TableStructure tempTableStructure = new TableStructure(this.m_ds.Tables[this.m_dg.CaptionText],"VIEW");
					break;
				case 1:     //max size the grid if it is in a multi-pane view
                    if (this.ReferenceGridViewForm != null)
                    {
                        if (((frmGridView)this.ParentForm).toolBar1.Buttons[0].Pushed == true) return;
                        ((frmGridView)this.ParentForm).GridViewMaxSize(this.m_dg.CaptionText);
                    }
					break;
				case 2:
					frmDialog frmTemp = new frmDialog();
					frmTemp.Visible=false;
					uc_print_report_wizard uc_print_report_wizard1 = new uc_print_report_wizard(this.m_dg,this.m_dg.CaptionText);
					frmTemp.Controls.Add(uc_print_report_wizard1);
					frmTemp.MaximizeBox = false;
					frmTemp.MinimizeBox = false;
					frmTemp.Width = uc_print_report_wizard1.m_DialogWd;
					frmTemp.Height = uc_print_report_wizard1.m_DialogHt;
					frmTemp.Text = "Print Report Wizard";
					uc_print_report_wizard1.Dock = System.Windows.Forms.DockStyle.Fill;
					frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
					uc_print_report_wizard1.Visible=true;
					frmTemp.ShowDialog();


					
					break;
				case 3:    //biosum_cond_id or biosum_plot_id record identification
					if (this.m_dg.CurrentRowIndex != -1)
					{
						frmDialog frmBiosumId = new frmDialog();
						frmBiosumId.Visible=false;

						uc_biosum_id uc_biosum_id1 = new uc_biosum_id(this.m_dg[this.m_dg.CurrentRowIndex,this.m_intBiosumIdColumn].ToString().Trim());
						frmBiosumId.Controls.Add(uc_biosum_id1);
						frmBiosumId.MaximizeBox = false;
						frmBiosumId.MinimizeBox = false;
						frmBiosumId.Width = uc_biosum_id1.m_intWd;
						frmBiosumId.Height = uc_biosum_id1.m_intHt;
						uc_biosum_id1.Visible=true;
						frmBiosumId.Text = "FIA Biosum Id (" + this.m_dg[this.m_dg.CurrentRowIndex,this.m_intBiosumIdColumn].ToString().Trim() + ")";
						uc_biosum_id1.Dock =  System.Windows.Forms.DockStyle.Fill;
						frmBiosumId.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
						uc_biosum_id1.btnClose.Left = uc_biosum_id1.Width - uc_biosum_id1.btnClose.Width -5;
						uc_biosum_id1.btnClose.Top = uc_biosum_id1.Height - uc_biosum_id1.btnClose.Height -5;
						frmBiosumId.ShowDialog();
					}

					break;
				case 4:  //save gridview contents
					this.savevalues();
                    break;
				
                
			}

		}

		public void savevalues()
        {
			int x;
			int y;
			string strSQL = "";
			string strSQLWhere = "";
			int intCurrRow;
			bool bFirstTime = false;
			SQLite.ADO.DataMgr p_dataMgr = new SQLite.ADO.DataMgr();

			this.m_intError = 0;

			/******************************************************
			 **save the current row, move the current row to a
			 **different row to enable getchanges() method, then
			 **move back to current row
			 ******************************************************/
			intCurrRow = this.m_dg.CurrentRowIndex;
			if (intCurrRow == 0)
			{
				this.m_dg.CurrentRowIndex++;
			}
			else
			{
				this.m_dg.CurrentRowIndex = 0;
			}

			System.Data.DataTable p_dtChanges;

			/******************************************************
			 **copy the rows that changed to the table p_dtChanges
			 ******************************************************/
			p_dtChanges = this.m_ds.Tables[this.m_dg.CaptionText].GetChanges(System.Data.DataRowState.Modified);
			if (p_dtChanges != null && p_dtChanges.Rows.Count > 0)
            {
				frmMain.g_oFrmMain.ActivateStandByAnimation(
					this.ParentForm.WindowState,
					this.ParentForm.Left,
					this.ParentForm.Height,
					this.ParentForm.Width,
					this.ParentForm.Top);

                try
                {
					/******************************************************
					 **save all the rows that changed
					 ******************************************************/

					for (x = 0; x <= p_dtChanges.Rows.Count - 1; x++)
					{
						/*****************************************************
						 **build the sql string
						 *****************************************************/
						bFirstTime = true;
						for (y = 0; y <= this.m_strColumnsToEdit.Length - 1; y++)
						{
							if (p_dtChanges.Rows[x][this.m_strColumnsToEdit[y]] != System.DBNull.Value)
							{
								if (bFirstTime == true)
								{
									bFirstTime = false;
									/***********************************************************
									 **check to see if the column data type is numeric or not
									 ***********************************************************/
									if (p_dataMgr.getIsTheFieldAStringDataType(p_dtChanges.Columns[this.m_strColumnsToEdit[y]].DataType.FullName.ToString()) == "Y")
									{
										/**********************************************************
										 **it is a string datatype so enclose the variable value
										 **with single quotation marks
										 **********************************************************/
										strSQL = "update " + this.m_dg.CaptionText + " set " +
											this.m_strColumnsToEdit[y] + "='" + p_dtChanges.Rows[x][this.m_strColumnsToEdit[y]].ToString() + "'";
									}
									else
									{
										/**********************************************
										 **not a string datatype
										 **********************************************/
										strSQL = "update " + this.m_dg.CaptionText + " set " +
											this.m_strColumnsToEdit[y] + "=" + p_dtChanges.Rows[x][this.m_strColumnsToEdit[y]].ToString();
									}
								}
								else
								{
									/***********************************************************
									 **check to see if the column data type is numeric or not
									 ***********************************************************/
									if (p_dataMgr.getIsTheFieldAStringDataType(p_dtChanges.Columns[this.m_strColumnsToEdit[y]].DataType.FullName.ToString()) == "Y")
									{
										/**********************************************************
										 **it is a string datatype so enclose the variable value
										 **with single quotation marks
										 **********************************************************/
										strSQL += "," +
											this.m_strColumnsToEdit[y] + "= '" + p_dtChanges.Rows[x][this.m_strColumnsToEdit[y]].ToString() + "'";
									}
									else
									{
										/**********************************************
										 **not a string datatype
										 **********************************************/
										strSQL += "," +
											this.m_strColumnsToEdit[y] + "=" + p_dtChanges.Rows[x][this.m_strColumnsToEdit[y]].ToString();
									}
								}
							}
						}
						/****************************************
						 **make sure that we have sql
						 ****************************************/
						if (strSQL.Trim().Length > 0)
						{
							/***********************************************************
							 **build the where clause to update the correct record(s)
							 ***********************************************************/
							for (y = 0; y <= this.m_strRecordKeyColumns.Length - 1; y++)
							{
								if (y == 0)
								{
									/***********************************************************
									 **check to see if the column data type is numeric or not
									 ***********************************************************/
									if (p_dataMgr.getIsTheFieldAStringDataType(p_dtChanges.Columns[this.m_strRecordKeyColumns[y]].DataType.FullName.ToString()) == "Y")
									{
										/**********************************************************
										 **it is a string datatype so enclose the variable value
										 **with single quotation marks
										 **********************************************************/
										strSQLWhere = " where " + this.m_strRecordKeyColumns[y].Trim() + " = '" + p_dtChanges.Rows[x][this.m_strRecordKeyColumns[y]].ToString() + "'";
									}
									else
									{
										/**********************************************
										 **not a string datatype
										 **********************************************/
										strSQLWhere = " where " + this.m_strRecordKeyColumns[y].Trim() + " = " + p_dtChanges.Rows[x][this.m_strRecordKeyColumns[y]].ToString();
									}
								}
								else
								{
									/***********************************************************
									 **check to see if the column data type is numeric or not
									 ***********************************************************/
									if (p_dataMgr.getIsTheFieldAStringDataType(p_dtChanges.Columns[this.m_strRecordKeyColumns[y]].DataType.FullName.ToString()) == "Y")
									{
										/**********************************************************
										 **it is a string datatype so enclose the variable value
										 **with single quotation marks
										 **********************************************************/
										strSQLWhere += " and " + this.m_strRecordKeyColumns[y].Trim() + " = '" + p_dtChanges.Rows[x][this.m_strRecordKeyColumns[y]].ToString() + "'";
									}
									else
									{
										/**********************************************
										 **not a string datatype
										 **********************************************/
										strSQLWhere += " and " + this.m_strRecordKeyColumns[y].Trim() + " = " + p_dtChanges.Rows[x][this.m_strRecordKeyColumns[y]].ToString();
									}
								}
							}

							strSQL += strSQLWhere + ";";

							p_dataMgr.SqlNonQuery(this.m_conn, strSQL);
							if (p_dataMgr.m_intError == -1)
							{
								p_dtChanges.Clear();
								p_dtChanges = null;
								this.m_intError = -1;
								return;
							}
						}
					}
				}
				catch (Exception caught)
				{
					frmMain.g_oFrmMain.DeactivateStandByAnimation();
					MessageBox.Show(caught.Message);
					p_dtChanges.Clear();
					p_dtChanges = null;
					this.m_intError = -1;
				}
				if (this.m_intError == 0)
				{
					frmMain.g_oFrmMain.DeactivateStandByAnimation();
					p_dtChanges.AcceptChanges(); //this.m_ds.Tables[this.m_dg.CaptionText].AcceptChanges();
				}
				p_dtChanges.Clear();
				p_dataMgr = null;
			}
			p_dtChanges = null;

			this.m_dg.CurrentRowIndex = intCurrRow;
			this.btnSave.Enabled = false;
			if (m_intError == 0)
			{
				if (this.ReferenceGridViewForm.ReferenceProcessorScenarioForm != null)
					this.ReferenceGridViewForm.ReferenceProcessorScenarioForm.m_bSave = true;
			}

			return;
		}
		private void m_dg_MouseWheel(object sender,MouseEventArgs e)
		{
			this.m_dg.Select();
		}
		private void InitializePopup()
		{
			this.m_mnuDataGridPopup = new ContextMenu();
			this.m_mnuDataGridTextBoxPopup = new ContextMenu();

			this.m_mnuDataGridPopup.MenuItems.Add("Filter By Value",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Filter By Entered Value", new EventHandler(Filter_Clicked));
            this.m_mnuDataGridPopup.MenuItems.Add("Remove Filter",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems.Add("Unique Values",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;
			this.m_mnuDataGridPopup.MenuItems.Add("-");
            this.m_mnuDataGridPopup.MenuItems.Add("Modify",new EventHandler(PopUp_Clicked));
			this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=false;
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
					this.sbDisplayedRecordCount.Text = this.m_dv.Count.ToString();

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
                    if (((frmGridView)this.ParentForm).HarvestCostColumns)
                    {

                        //------------text box------------//
                        //define form properties
                        frmTemp.Width = 200;
                        frmTemp.Height = 200;
                        frmTemp.WindowState = System.Windows.Forms.FormWindowState.Normal;
                        frmTemp.MaximizeBox = false;
                        frmTemp.MinimizeBox = false;
                        frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                        frmTemp.Text = "Modify (" + strCol.Trim() + ")";

                        //define numeric text class properties
                        FIA_Biosum_Manager.TemplateTextBox oTextBox = new TemplateTextBox((System.Windows.Forms.Control)frmTemp, "txtModify", "$0.00", true);
                        oTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
                        oTextBox.TabIndex = 0;
                        oTextBox.Tag = "";
                        oTextBox.Visible = true;
                        oTextBox.Enabled = true;
                        oTextBox.Height = 100;
                        oTextBox.Top = (int)(frmTemp.ClientSize.Height * .50) - (int)(oTextBox.Height * .50);
                        oTextBox.Width = frmTemp.ClientSize.Width - 20;
                        oTextBox.Left = 10;
                        intLeft = oTextBox.Left;
                        intTop = oTextBox.Top;
                        oTextBox.m_oValidateNumericValues.Money = true;
                        oTextBox.m_oValidateNumericValues.RoundDecimalLength = 2;
                        oTextBox.m_oValidateNumericValues.NullsAllowed = false;
                        oTextBox.m_oValidateNumericValues.MaxValue = 999.99;
                        oTextBox.m_oValidateNumericValues.MinValue = 0.00;
                        oTextBox.m_oValidateNumericValues.TestForMax = true;
                        oTextBox.m_oValidateNumericValues.TestForMin = true;

                        
                        frmTemp.txtBox=oTextBox;
                        
                        oTextBox.Focus();

                        frmTemp.strCallingFormType = "TM";


                    }
                     
					else
					{
						this.m_intTableSchemaColumnRow = this.getTableSchemaColumnDefinition(this.m_dv.Table.Columns[this.m_dg.CurrentCell.ColumnNumber].ColumnName.ToString());
						if (this.m_intTableSchemaColumnRow != -1)
						{
                            switch (this.m_dtTableSchema.Rows[this.m_intTableSchemaColumnRow]["DATATYPE"].ToString().Trim())
							{
								case "System.String":
									//------------text box------------//
                                    System.Windows.Forms.TextBox p_txtString = new TextBox();
		                    
									//define form properties
									frmTemp.Width = 200;
									frmTemp.Height = 200;
									frmTemp.WindowState = System.Windows.Forms.FormWindowState.Normal;
									frmTemp.MaximizeBox = false;
									frmTemp.MinimizeBox = false;
									frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
									frmTemp.Text = "Modify (" + strCol.Trim() + ")";
									//define numeric text class properties
									p_txtString.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
									p_txtString.Name = "txtModify";
									p_txtString.TabIndex = 0;
									p_txtString.Tag = "";
									p_txtString.Visible = true;
									p_txtString.Enabled = true;
									frmTemp.Controls.Add(p_txtString);  //add the text box to the form
									p_txtString.Height = 100;

									p_txtString.Top = (int)(frmTemp.ClientSize.Height * .50) - (int)(p_txtString.Height * .50);
									p_txtString.Width = frmTemp.ClientSize.Width - 20;
									p_txtString.Left = 10;

									p_txtString.ReadOnly=false;
									p_txtString.Text = "";
									if (this.m_dtTableSchema.Rows[this.m_intTableSchemaColumnRow]["COLUMNNAME"].ToString().Trim().ToUpper() == "FVS_VARIANT")
									{
										p_txtString.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
									}
									if (Convert.ToInt32(this.m_dtTableSchema.Rows[this.m_intTableSchemaColumnRow]["COLUMNSIZE"]) > 0)
									{
										p_txtString.MaxLength = Convert.ToInt32(this.m_dtTableSchema.Rows[this.m_intTableSchemaColumnRow]["COLUMNSIZE"]);
									}

									intLeft = p_txtString.Left;
									intTop = p_txtString.Top;
									frmTemp.txtBox = p_txtString;
									frmTemp.txtBox.Focus();
									frmTemp.txtBox.SelectionStart = 1;
									frmTemp.strCallingFormType="TS";
									break;

								case "System.Integer":
									break;
								case "System.Double":
							
									//------------text box------------//
									//instatiate numeric text class
									FIA_Biosum_Manager.txtNumeric p_txtDouble = new FIA_Biosum_Manager.txtNumeric(12,20);
		                    
									//define form properties
									frmTemp.Width = 200;
									frmTemp.Height = 200;
									frmTemp.WindowState = System.Windows.Forms.FormWindowState.Normal;
									frmTemp.MaximizeBox = false;
									frmTemp.MinimizeBox = false;
									frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
									frmTemp.Text = "Modify (" + strCol.Trim() + ")";

									//define numeric text class properties
									p_txtDouble.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
									p_txtDouble.Name = "txtModify";
									p_txtDouble.TabIndex = 0;
									p_txtDouble.Tag = "";
									p_txtDouble.Visible = true;
									p_txtDouble.Enabled = true;
									frmTemp.Controls.Add(p_txtDouble);  //add the text box to the form
									p_txtDouble.Height = 100;

									p_txtDouble.Top = (int)(frmTemp.ClientSize.Height * .50) - (int)(p_txtDouble.Height * .50);
									p_txtDouble.Width = frmTemp.ClientSize.Width - 20;
									p_txtDouble.Left = 10;
									p_txtDouble.bEdit=true;
									p_txtDouble.ReadOnly=false;
									p_txtDouble.Text = "0";
									intLeft = p_txtDouble.Left;
									intTop = p_txtDouble.Top;
									frmTemp.txtNumeric = p_txtDouble;
									frmTemp.txtNumeric.Focus();
									frmTemp.txtNumeric.SelectionStart=1;
									frmTemp.strCallingFormType="TD";

									break;
						
								default:
									//------------text box------------//
									//instatiate numeric text class
									FIA_Biosum_Manager.txtNumeric p_txtDefault = new FIA_Biosum_Manager.txtNumeric(4,0);

									//define form properties
									frmTemp.Width = 200;
									frmTemp.Height = 200;
									frmTemp.WindowState = System.Windows.Forms.FormWindowState.Normal;
									frmTemp.MaximizeBox = false;
									frmTemp.MinimizeBox = false;
									frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
									frmTemp.Text = "Modify (" + strCol.Trim() + ")";
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
									p_txtDefault.Text = "0";

									intLeft = p_txtDefault.Left;
									intTop = p_txtDefault.Top;
									frmTemp.txtNumeric = p_txtDefault;
									frmTemp.txtNumeric.Focus();
									frmTemp.txtNumeric.SelectionStart = 1;
									frmTemp.strCallingFormType="TN";
									break;
							}
						}
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
                        frmMain.g_oFrmMain.ActivateStandByAnimation(
                            this.ParentForm.WindowState,
                            this.ParentForm.Left,
                            this.ParentForm.Height,
                            this.ParentForm.Width,
                            this.ParentForm.Top);
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
                            if (bApplyModifyCount)
                            {
                                //get the unique ids for each of the selected records
                                string[,] strKeyRows = new string[y, this.m_strRecordKeyColumns.Length];
                                int intKeyRowCount = 0;
                                int zz,zzz;
                                for (x = 0; x <= y - 1; x++)
                                {
                                    if (m_dg.IsSelected(x))
                                    {
                                        for (zz = 0; zz <= this.m_intRecordKeyColumns.Length - 1; zz++)
                                        {
                                            strKeyRows[intKeyRowCount, zz] = Convert.ToString(m_dg[x, m_intRecordKeyColumns[zz]]);
                                        }
                                        intKeyRowCount++;
                                    }
                                }
                                //define the primary key for the table
                                DataColumn[] colPk = new DataColumn[m_intRecordKeyColumns.Length];
                                for (x = 0; x <= this.m_intRecordKeyColumns.Length - 1; x++)
                                {
                                   colPk[x] = m_ds.Tables[this.m_dg.CaptionText].Columns[this.m_strRecordKeyColumns[x]];
                                }
                                this.m_ds.Tables[this.m_dg.CaptionText].PrimaryKey = colPk;
                               //find each record
                               for (zz = 0; zz <= intKeyRowCount-1; zz++)
                                {
                                    System.Object[] oSearch = new Object[m_intRecordKeyColumns.Length];
                                    //load up the unique id for the record
                                   for (zzz = 0; zzz <= this.m_intRecordKeyColumns.Length - 1; zzz++)
                                   {
                                        oSearch[zzz]=strKeyRows[zz,zzz].Trim();
                                   }
                                   //search for the record
                                   System.Data.DataRow oRow = this.m_ds.Tables[this.m_dg.CaptionText].Rows.Find(oSearch);
                                   if (oRow != null)
                                   {
                                        //update the record
                                        oRow.BeginEdit();
                                        oRow[m_intPopupColumn] = frmTemp.txtBox.Text.Replace("$", "");
                                        oRow.EndEdit();
                                   }
                               }
                                
                               
                            }
                            else
                            {
                                for (x = 0; x <= y - 1; x++)
                                {
                                    if (m_dg.IsSelected(x) || x == intCurrRow)
                                    {
                                        if (bApplyModifyCount == true)
                                        {
                                            m_dg.CurrentRowIndex = x - intModifyCount;
                                        }
                                        else
                                        {
                                            m_dg.CurrentRowIndex = x;
                                        }

                                        if (frmTemp.strCallingFormType == "TS")
                                        {
                                            m_dg[m_dg.CurrentRowIndex, this.m_intPopupColumn] = frmTemp.txtBox.Text;

                                        }
                                        else if (frmTemp.strCallingFormType == "TN")
                                        {
                                            m_dg[x, this.m_intPopupColumn] = frmTemp.txtNumeric.Text;
                                        }
                                        else if (frmTemp.strCallingFormType == "TD")
                                        {
                                            m_dg[x, this.m_intPopupColumn] = frmTemp.txtNumeric.Text;
                                        }
                                        else if (frmTemp.strCallingFormType == "TM")
                                        {
                                            m_dg[x, this.m_intPopupColumn] = frmTemp.txtBox.Text.Replace("$", "");
                                        }
                                        if (this.btnSave.Enabled == false) this.btnSave.Enabled = true;
                                        intModifyCount++;
                                    }
                                }
                            }
							this.m_dg.SetDataBinding(this.m_dv,"");
							this.m_dg.Update();
                            frmMain.g_oFrmMain.DeactivateStandByAnimation();
							
						}
						catch (Exception err)
						{
                            frmMain.g_oFrmMain.DeactivateStandByAnimation();
							MessageBox.Show(err.Message);
						}
					}
					frmTemp = null;
					break;
				case "REMOVE FILTER":
					this.m_dv.RowFilter="";
					this.m_strColumnFilterList="";
					this.m_mnuDataGridPopup.MenuItems[MENU_REMOVEFILTER].Enabled=false;
					this.sbDisplayedRecordCount.Text = this.m_dv.Count.ToString();
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
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					strFilter = "Max(" + strCol + ")";
					this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
					if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
					{
                         MessageBox.Show(strCol + " Maximum: " + this.m_ds.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
					}
					else
					{
						 MessageBox.Show(strCol + " Maximum: " + this.m_ds.Tables[0].Compute(strFilter, null));
					}
					break;
				case "MINIMUM":
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					strFilter = "Min(" + strCol + ")";
					this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
					if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
					{
						MessageBox.Show(strCol + " Minimum: " + this.m_ds.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
					}
					else
					{
						MessageBox.Show(strCol + " Minimum: " + this.m_ds.Tables[0].Compute(strFilter, null));
					}
					break;

				case "AVERAGE":
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					strFilter = "Avg(" + strCol + ")";
					this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
					if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
					{
						MessageBox.Show(strCol + " Average: " + this.m_ds.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
					}
					else
					{
						MessageBox.Show(strCol + " Average: " + this.m_ds.Tables[0].Compute(strFilter, null));
					}
					break;

				case "SUM":
					strCol = this.m_dg.TableStyles[0].GridColumnStyles[this.m_intPopupColumn].HeaderText.Trim();
					strFilter = "Sum(" + strCol + ")";
					this.m_dv.RowStateFilter = System.Data.DataViewRowState.CurrentRows;
					if (this.m_dv.RowFilter.ToString().Trim().Length > 0)
					{
						MessageBox.Show(strCol + " Sum: " + this.m_ds.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString()));
					}
					else
					{
						MessageBox.Show(strCol + " Sum: " + this.m_ds.Tables[0].Compute(strFilter, null));
					}
					break;

				case "COUNT BY VALUE":
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
						MessageBox.Show(strCol + " Count: " + this.m_ds.Tables[0].Compute(strFilter,this.m_dv.RowFilter.ToString() + " and " + strExp));
					}
					else
					{
						MessageBox.Show(strCol + " Count: " + this.m_ds.Tables[0].Compute(strFilter, strExp));
					}
					break;
				case "SELECT ALL":
					int intRows = this.m_dg.BindingContext[this.m_dg.DataSource,this.m_dg.DataMember].Count;
					for (x=0; x<=intRows-1;x++)
					{
						this.m_dg.Select(x);
					}
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
				this.sbDisplayedRecordCount.Text = this.m_dv.Count.ToString();
				if (this.m_strColumnFilterList.Trim().Length > 0) this.m_strColumnFilterList+=  strCol.Trim() + ",";
				else this.m_strColumnFilterList = strCol.Trim() + ",";
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
								this.m_mnuDataGridPopup.MenuItems[MENU_MODIFY].Enabled=false;
								this.m_mnuDataGridPopup.MenuItems[MENU_COUNTBYVALUE].Enabled=false;
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=false;

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
								this.m_mnuDataGridPopup.MenuItems[MENU_FILTERBYENTEREDVALUE].Enabled=true;
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

						this.m_mnuDataGridPopup.Show(this, new Point(pt.X,pt.Y + this.m_dg.Top));
					}

					break;
				case System.Windows.Forms.MouseButtons.Left :
					if (this.m_dg.CurrentRowIndex != this.m_intMouseUpCurrRow - 1)
					{
						m_intMouseUpCurrRow=this.m_dg.CurrentRowIndex+1;
					}
					Point pt2 = new Point(e.X,e.Y);
					try
					{
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
			//string str;
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
		

        public void LoadGridView(string strConn, string strSQL, string strDataSetName)
        {
            string strColumnName = "";
            this.InitializePopup();
            this.m_dg.MouseWheel += new MouseEventHandler(m_dg_MouseWheel);
            this.m_conn = new System.Data.SQLite.SQLiteConnection();

            if (strDataSetName != null && strDataSetName.Trim().Length > 0)
            {
                this.m_ds = new DataSet(strDataSetName);
            }
            else
            {
                this.m_ds = new DataSet();
            }

            this.m_da = new System.Data.SQLite.SQLiteDataAdapter();

            this.m_dg.Left = 5;
            this.toolBar1.Left = 5;
            this.btnClose.Top = this.groupBox1.Top + 10;

            this.m_dg.Width = this.groupBox1.Width - 10;
            this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - this.groupBox1.Left - 10;
            this.toolBar2.Top = this.statusBar1.Top - this.toolBar2.Height - 2;
            this.toolBar2.Left = (int)(this.groupBox1.Width * .50) - (int)(this.toolBar2.Width * .50);
            this.m_dg.Height = this.toolBar2.Top - this.m_dg.Top;

            var p_dataMgr = new SQLite.ADO.DataMgr();
            p_dataMgr.OpenConnection(strConn, ref this.m_conn);
            if (p_dataMgr.m_intError != 0)
            {
                this.m_intError = p_dataMgr.m_intError;
                p_dataMgr = null;
                return;
            }

            this.m_da.SelectCommand = new System.Data.SQLite.SQLiteCommand(strSQL, this.m_conn);
            try
            {

                this.m_da.Fill(this.m_ds, strDataSetName);
                this.m_dv = new DataView(this.m_ds.Tables[strDataSetName]);

                this.m_dv.AllowNew = false;       //cannot append new records
                this.m_dv.AllowDelete = false;    //cannot delete records
                this.m_dv.AllowEdit = false;
                this.m_dg.CaptionText = strDataSetName;
                m_dg.BackgroundColor = frmMain.g_oGridViewBackgroundColor;
                /***********************************************************************************
				 **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
				 ***********************************************************************************/
                gridview_DataGridColoredTextBoxColumn aColumnTextColumn;


                /***************************************************************
				 **custom define the grid style
				 ***************************************************************/
                DataGridTableStyle tableStyle = new DataGridTableStyle();

                /***********************************************************************
				 **map the data grid table style to the scenario rx intensity dataset
				 ***********************************************************************/
                tableStyle.MappingName = strDataSetName;
                tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
                tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
                tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
                tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;


                /******************************************************************************
				 **since the dataset has things like field name and number of columns,
				 **we will use those to create new columnstyles for the columns in our grid
				 ******************************************************************************/
                //get the number of columns from the scenario_rx_intensity data set
                int numCols = this.m_ds.Tables[strDataSetName].Columns.Count;


                /************************************************
				 **loop through all the columns in the dataset	
				 ************************************************/
                for (int i = 0; i < numCols; ++i)
                {
                    strColumnName = this.m_ds.Tables[strDataSetName].Columns[i].ColumnName;


                    /******************************************************************
					 **create a new instance of the DataGridColoredTextBoxColumn class
					 ******************************************************************/
                    aColumnTextColumn = new gridview_DataGridColoredTextBoxColumn(false, this);


                    /***********************************
					 **all columns are read-only except
					 **the edit columns
					 ***********************************/
                    aColumnTextColumn.ReadOnly = true;


                    aColumnTextColumn.HeaderText = strColumnName;

                    /********************************************************************
					 **assign the mappingname property the data sets column name
					 ********************************************************************/
                    aColumnTextColumn.MappingName = strColumnName;
                    //aColumnTextColumn
                    aColumnTextColumn.TextBox.ContextMenu = new ContextMenu(); //this.m_mnuDataGridPopup;
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

                this.m_dg.DataSource = this.m_dv;

                this.m_dg.Expand(-1);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "LoadGridViewSqlite", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.m_intError = -1;
                this.m_conn.Close();
                this.m_conn = null;
                this.m_ds.Clear();
                this.m_ds = null;
                this.m_da.Dispose();
                this.m_da = null;
                return;

            }

            p_dataMgr = null;

            this.sbMsg.Text = strDataSetName + " (Read Only)";
            if (this.m_ds.Tables[strDataSetName].Rows.Count == 0)
            {
                this.sbQueryRecordCount.Text = "0/0";
                this.sbDisplayedRecordCount.Text = "0";
            }
            else
            {
                this.m_intCurrRow = 1;
                this.sbQueryRecordCount.Text = "1/" + this.m_ds.Tables[strDataSetName].Rows.Count.ToString().Trim();

                this.sbDisplayedRecordCount.Text = Convert.ToString(this.m_dg.BindingContext[this.m_dg.DataSource, this.m_dg.DataMember].Count);
            }

            this.sbQueryRecordCount.Width =
                (int)this.CreateGraphics().MeasureString("9999999/9999999", this.statusBar1.Font).Width;
            this.sbMsg.Width = (int)(this.groupBox1.Width * .50) - (int)(this.sbQueryRecordCount.Width * .50);
            this.sbDisplayedRecordCount.Width = (int)(this.groupBox1.Width * .50) - (int)(this.sbQueryRecordCount.Width * .50);
            this.sbDisplayedRecordCount.Alignment = System.Windows.Forms.HorizontalAlignment.Right;
            this.sbQueryRecordCount.Alignment = System.Windows.Forms.HorizontalAlignment.Center;

            //event handler to keep track of current row and cell movement
            this.m_dg.CurrentCellChanged += new
                System.EventHandler(this.m_dg_CurrentCellChanged);

            this.m_strSQL = strSQL;
            if (this.m_strSQL.Trim().Length > 0) this.btnSQL.Enabled = true;

            /******************************************************************
			 **if biosum_cond_id or biosum_plot_id is part of the dataset then
			 **enable the id toolbar button
			 ******************************************************************/
            int x = 0;
            for (x = 0; x <= this.m_ds.Tables[strDataSetName].Columns.Count - 1; x++)
            {
                if (this.m_ds.Tables[strDataSetName].Columns[x].ColumnName.Trim().ToUpper() == "BIOSUM_COND_ID")
                {

                    this.m_intBiosumIdColumn = x;
                    break;
                }
                if (this.m_ds.Tables[strDataSetName].Columns[x].ColumnName.Trim().ToUpper() == "BIOSUM_PLOT_ID")
                {

                    this.m_intBiosumIdColumn = x;
                    break;
                }
            }
            if (x <= this.m_ds.Tables[strDataSetName].Columns.Count - 1)
            {
                this.toolBar1.Buttons[3].Enabled = true;
            }
            else
            {
                this.toolBar1.Buttons[3].Enabled = false;
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
					
					if (this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_UNIQUEVALUES].Enabled=true;
					if (this.m_mnuDataGridPopup.MenuItems[MENU_SELECTALL].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_SELECTALL].Enabled=true;
					if (this.m_mnuDataGridPopup.MenuItems[MENU_MIN].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_MIN].Enabled=true;
					if (this.m_mnuDataGridPopup.MenuItems[MENU_MAX].Enabled==false)
						this.m_mnuDataGridPopup.MenuItems[MENU_MAX].Enabled=true;
					   
				    break;
			}
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

		private void m_dg_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			 


		}

		private void uc_gridview_RowDeleted(object sender, DataRowChangeEventArgs e)
		{
		}

		private void uc_gridview_RowDeleting(object sender, DataRowChangeEventArgs e)
		{
		}

		private void uc_gridview_RowChanging(object sender, DataRowChangeEventArgs e)
		{
		}
		public bool CloseButton_Visible
		{
			get {return btnClose.Visible;}
			set {btnClose.Visible=value;}
		}
		public FIA_Biosum_Manager.frmGridView ReferenceGridViewForm
		{
			get {return this._frmGridView;}
			set {this._frmGridView=value;}
		}

	}
	public class gridview_DataGridColoredTextBoxColumn : DataGridTextBoxColumn
	{
		bool m_bEdit=false;
		uc_gridview uc_gridview1;
		string m_strLastKey="";
		

		public gridview_DataGridColoredTextBoxColumn(bool bEdit,uc_gridview p_uc_gridview)
		{
			this.m_bEdit = bEdit;
			this.uc_gridview1 = p_uc_gridview;
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
		}
		private void TextBox_KeyDown(object sender, KeyEventArgs e)
		{
			if (this.m_bEdit == true)
			{
				if (this.m_strLastKey.Trim().Length == 0) this.m_strLastKey = Convert.ToString(e.KeyValue);
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
					if (this.uc_gridview1.btnSave.Enabled==false) this.uc_gridview1.btnSave.Enabled=true;
				}
			}
		}

		     
	}

}

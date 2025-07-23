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
	/// Summary description for uc_rx_harvest_cost_column_list.
	/// </summary>
	public class uc_rx_harvest_cost_column_list : System.Windows.Forms.UserControl
	{
		public int m_intError=0;
		public string m_strError="";
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.ListView lvRxHarvestCostColumns;
		private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvAlternateRowColors=new ListViewAlternateBackgroundColors();
		const int COLUMN_NULL=0;
		const int COLUMN_RX=1;
		const int COLUMN_FIELD=2;
		const int COLUMN_DESC=3;
		private System.Windows.Forms.Label lblDesc;
		private Queries m_oQueries = new Queries();
		private FIA_Biosum_Manager.frmRxItem _frmRxItem=null;
        private string m_strHarvestTableColumnNameList = "";
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FIA_Biosum_Manager.frmRxItem ReferenceFormRxItem
		{
			get {return this._frmRxItem;}
			set {this._frmRxItem=value;}

		}
		public uc_rx_harvest_cost_column_list()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.lvRxHarvestCostColumns.View = System.Windows.Forms.View.Details;
			this.m_oLvAlternateRowColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvAlternateRowColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvAlternateRowColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvAlternateRowColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvAlternateRowColors.ReferenceSelectedRowBackgroundColor=frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvAlternateRowColors.CustomFullRowSelect=true;
			this.m_oLvAlternateRowColors.ReferenceListView = lvRxHarvestCostColumns;
			if (frmMain.g_oGridViewFont != null) this.lvRxHarvestCostColumns.Font = frmMain.g_oGridViewFont;

			// TODO: Add any initialization after the InitializeComponent call

		}
		public void loadvalues()
		{
			int x,y,z;

			m_oQueries.m_oFvs.LoadDatasource=true;
			m_oQueries.m_oReference.LoadDatasource=true;
            m_oQueries.m_oProcessor.LoadDatasource = true;
			m_oQueries.LoadDatasourcesNew(true);			
			this.lvRxHarvestCostColumns.Clear();
			
			this.m_oLvAlternateRowColors.InitializeRowCollection();
			this.lvRxHarvestCostColumns.Columns.Add("",2,HorizontalAlignment.Left);
			this.lvRxHarvestCostColumns.Columns.Add("Rx", 80, HorizontalAlignment.Left);
			this.lvRxHarvestCostColumns.Columns.Add("Additional Cost Component", 200, HorizontalAlignment.Left);
			this.lvRxHarvestCostColumns.Columns.Add("Description", 300, HorizontalAlignment.Left);

			this.m_intError=0;
			this.m_strError="";

			if (this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection != null)
			{
				this.lvRxHarvestCostColumns.BeginUpdate();
			
				for (x=0;x<=this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Count-1;x++)
				{

                    if (ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(x).Delete == false)
                    {
                        FIA_Biosum_Manager.RxItemHarvestCostColumnItem oItem =
                            ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(x);
                        this.lvRxHarvestCostColumns.Items.Add("");
                        this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].UseItemStyleForSubItems = false;
                        for (z = 1; z <= this.lvRxHarvestCostColumns.Columns.Count - 1; z++)
                        {
                            this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems.Add(" ");
                        }
                        this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems[COLUMN_RX].Text = oItem.RxId;
                        this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems[COLUMN_FIELD].Text = oItem.HarvestCostColumn;
                        this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems[COLUMN_DESC].Text = oItem.Description;

                        this.m_oLvAlternateRowColors.AddRow();
                        this.m_oLvAlternateRowColors.AddColumns(lvRxHarvestCostColumns.Items.Count - 1, this.lvRxHarvestCostColumns.Columns.Count);
                    }
						
					
						
				}
				this.lvRxHarvestCostColumns.EndUpdate();
			}
		
			
			this.m_oLvAlternateRowColors.ListView();

            if (this.lvRxHarvestCostColumns.Items.Count > 0)
            {
                this.lvRxHarvestCostColumns.Items[0].Selected = true;
                ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_CLEARALL] = true;
                ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_DELETE] = true;
                ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_EDIT] = true;
                
            }
		}
		public void savevalues()
		{

			if (this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection != null)
			{
			
				//make sure each command is assigned the rxid
				for (int x=0;x<=this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Count-1;x++)
				{
					if (this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(x).RxId.Trim().Length ==0)
					{
						ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(x).RxId=ReferenceFormRxItem.m_oRxItem.RxId;
						
					}
				}
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_rx_harvest_cost_column_list));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblDesc = new System.Windows.Forms.Label();
            this.lvRxHarvestCostColumns = new System.Windows.Forms.ListView();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(648, 392);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.lblDesc);
            this.panel1.Controls.Add(this.lvRxHarvestCostColumns);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 18);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(642, 371);
            this.panel1.TabIndex = 0;
            this.panel1.Resize += new System.EventHandler(this.panel1_Resize);
            // 
            // lblDesc
            // 
            this.lblDesc.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblDesc.Location = new System.Drawing.Point(0, 0);
            this.lblDesc.Name = "lblDesc";
            this.lblDesc.Size = new System.Drawing.Size(642, 29);
            this.lblDesc.TabIndex = 1;
            this.lblDesc.Text = resources.GetString("lblDesc.Text");
            this.lblDesc.Click += new System.EventHandler(this.lblDesc_Click);
            // 
            // lvRxHarvestCostColumns
            // 
            this.lvRxHarvestCostColumns.GridLines = true;
            this.lvRxHarvestCostColumns.HideSelection = false;
            this.lvRxHarvestCostColumns.Location = new System.Drawing.Point(8, 32);
            this.lvRxHarvestCostColumns.MultiSelect = false;
            this.lvRxHarvestCostColumns.Name = "lvRxHarvestCostColumns";
            this.lvRxHarvestCostColumns.Size = new System.Drawing.Size(616, 328);
            this.lvRxHarvestCostColumns.TabIndex = 0;
            this.lvRxHarvestCostColumns.UseCompatibleStateImageBehavior = false;
            this.lvRxHarvestCostColumns.View = System.Windows.Forms.View.Details;
            this.lvRxHarvestCostColumns.SelectedIndexChanged += new System.EventHandler(this.lvRxHarvestCostColumns_SelectedIndexChanged);
            this.lvRxHarvestCostColumns.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lvRxHarvestCostColumns_MouseUp);
            // 
            // uc_rx_harvest_cost_column_list
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_rx_harvest_cost_column_list";
            this.Size = new System.Drawing.Size(648, 392);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void lblDesc_Click(object sender, System.EventArgs e)
		{
		
		}

		private void panel1_Resize(object sender, System.EventArgs e)
		{
			this.lvRxHarvestCostColumns.Width = this.panel1.ClientSize.Width - this.lvRxHarvestCostColumns.Left - this.panel1.AutoScrollMargin.Width;
			this.lvRxHarvestCostColumns.Height = this.panel1.ClientSize.Height - this.lvRxHarvestCostColumns.Top - this.panel1.AutoScrollMargin.Height;
		}
        public void EditItem()
        {
            Edit("Edit");
        }
        public void AddItem()
        {
            GetExistingColumns();
            Edit("New");
            
        }
        private void AddItemToList(FIA_Biosum_Manager.RxItemHarvestCostColumnItem oItem)
        {

            this.lvRxHarvestCostColumns.Items.Add("");
            this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].UseItemStyleForSubItems = false;
            for (int z = 1; z <= this.lvRxHarvestCostColumns.Columns.Count - 1; z++)
            {
                this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems.Add(" ");
            }
           
            this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems[COLUMN_RX].Text = oItem.RxId;
            this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems[COLUMN_FIELD].Text = oItem.HarvestCostColumn;
            this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems[COLUMN_DESC].Text = oItem.Description;
           

            this.m_oLvAlternateRowColors.AddRow();
            this.m_oLvAlternateRowColors.AddColumns(lvRxHarvestCostColumns.Items.Count - 1, this.lvRxHarvestCostColumns.Columns.Count);

            
            ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_CLEARALL] = true;
            ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_DELETE] = true;
            ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_EDIT] = true;
            ReferenceFormRxItem.SetToolBarButtonsEnabled(frmRxItem.UC_HARVESTCOST);

            this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].Selected = true;
        }
        private void GetExistingColumns()
        {
            int x;
            string strCol = "";
            System.Collections.Generic.List<string> lstFields = new System.Collections.Generic.List<string>();

            DataMgr oDataMgr = new DataMgr();
            System.Collections.Generic.List<string> lstProcessorColumns = null;
            string strProcessorDefinitionsDb = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile}";
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strProcessorDefinitionsDb)))
            {
                conn.Open();
                lstProcessorColumns = oDataMgr.getStringList(conn, "SELECT DISTINCT ColumnName FROM scenario_harvest_cost_columns");
            }

            if (lstProcessorColumns.Count > 0)
            {
                 lstFields = lstProcessorColumns;
            }
                       
            string strSource = m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.RxHarvestCostColumns);
            string strTable = m_oQueries.m_oDataSource.getValidDataSourceTableName("TREATMENT PRESCRIPTIONS HARVEST COST COLUMNS").Trim();
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strSource)))
            {
                conn.Open();
                oDataMgr.m_strSQL = "SELECT DISTINCT ColumnName FROM " + strTable;
                oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
                if (oDataMgr.m_DataReader.HasRows)
                {
                    while (oDataMgr.m_DataReader.Read())
                    {
                        if (oDataMgr.m_DataReader["ColumnName"] != System.DBNull.Value &&
                            Convert.ToString(oDataMgr.m_DataReader["ColumnName"]).Trim().Length > 0)
                        {
                            strCol = Convert.ToString(oDataMgr.m_DataReader["ColumnName"]).Trim().ToUpper();
                            if (lstProcessorColumns != null)
                            {
                                for (x = 0; x <= lstProcessorColumns.Count - 1; x++)
                                {
                                    if (lstProcessorColumns[x] != null)
                                    {
                                        if (lstProcessorColumns[x].Trim().ToUpper() ==
                                            strCol) break;
                                    }
                                }
                                if (x > lstProcessorColumns.Count - 1)
                                {
                                    lstFields.Add(Convert.ToString(oDataMgr.m_DataReader["ColumnName"]).Trim());
                                }
                            }
                            else
                                lstFields.Add(Convert.ToString(oDataMgr.m_DataReader["ColumnName"]).Trim());
                        }
                    }
                }
                oDataMgr.m_DataReader.Close();
            }

            m_strHarvestTableColumnNameList = "";
            if (lstFields != null && lstFields.Count > 0)
            {
                foreach (var item in lstFields)
                {
                    m_strHarvestTableColumnNameList = m_strHarvestTableColumnNameList + $@"{item},";
                }
                m_strHarvestTableColumnNameList = m_strHarvestTableColumnNameList.Substring(0, m_strHarvestTableColumnNameList.Length - 1); //Remove trailing comma
            }
        }

		private void Edit(string strType)
		{
			int y,intIndex,intCount;
			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
			frmTemp.MaximizeBox = false;
			frmTemp.BackColor = System.Drawing.SystemColors.Control;
			frmTemp.Initialize_Scenario_Harvest_Costs_Column_Edit_Control();
			string strColumnList="";

            if (this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection != null)
            {
                for (y = 0; y <= this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Count - 1; y++)
                {
                    if (this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(y).Delete == false)
                    {
                        if (this.ReferenceFormRxItem.m_oRxItem.RxId == this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(y).RxId)
                            strColumnList = strColumnList + ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(y).HarvestCostColumn + ",";
                    }
                }
                if (strColumnList.Trim().Length > 0) strColumnList = strColumnList.Substring(0, strColumnList.Length - 1);
            }
			

			frmTemp.Height=0;
			frmTemp.Width=0;
			if (frmTemp.Top + frmTemp.uc_scenario_harvest_cost_column_edit1.Height > frmTemp.ClientSize.Height + 2)
			{
				for (y=1;;y++)
				{
					frmTemp.Height = y;
					if (frmTemp.uc_scenario_harvest_cost_column_edit1.Top + 
						frmTemp.uc_scenario_harvest_cost_column_edit1.Height < 
						frmTemp.ClientSize.Height)
					{
						break;
					}
				}

			}
			if (frmTemp.uc_scenario_harvest_cost_column_edit1.Left + frmTemp.uc_scenario_harvest_cost_column_edit1.Width > frmTemp.ClientSize.Width + 2)
			{
				for (y=1;;y++)
				{
					frmTemp.Width = y;
					if (frmTemp.uc_scenario_harvest_cost_column_edit1.Left + 
						frmTemp.uc_scenario_harvest_cost_column_edit1.Width < 
						frmTemp.ClientSize.Width)
					{
						break;
					}
				}

			}
			frmTemp.Left = 0;
			frmTemp.Top = 0;
      
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			frmTemp.uc_scenario_harvest_cost_column_edit1.Dock = System.Windows.Forms.DockStyle.Fill;		
			
			frmTemp.uc_scenario_harvest_cost_column_edit1.EditType=strType;
			
			if (strType.Trim().ToUpper() == "NEW")
			{
				frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnList = strColumnList;
				strColumnList="";
				for (y=0;y<=this.lvRxHarvestCostColumns.Items.Count-1;y++)
				{
					strColumnList=strColumnList + this.lvRxHarvestCostColumns.Items[y].SubItems[COLUMN_FIELD].Text.Trim() + ",";
				}
				if (strColumnList.Trim().Length > 0)
					strColumnList=strColumnList.Substring(0,strColumnList.Length - 1);
				frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText = "";
				frmTemp.uc_scenario_harvest_cost_column_edit1.CurrentSelectedColumnList=strColumnList;
				frmTemp.uc_scenario_harvest_cost_column_edit1.HarvestCostTableColumnList = this.m_strHarvestTableColumnNameList;
				frmTemp.uc_scenario_harvest_cost_column_edit1.loadvalues();
				frmTemp.uc_scenario_harvest_cost_column_edit1.lblEdit.Show();
			}
			else
			{
                frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText = this.lvRxHarvestCostColumns.SelectedItems[0].SubItems[COLUMN_FIELD].Text;
				frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnDescription = this.lvRxHarvestCostColumns.SelectedItems[0].SubItems[COLUMN_DESC].Text;
				frmTemp.uc_scenario_harvest_cost_column_edit1.cmbCol.Enabled=false;
				frmTemp.uc_scenario_harvest_cost_column_edit1.lblEdit.Hide();
			}
			
			frmTemp.Text = strType + " Additional CPA Component";
			System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
			if (result==System.Windows.Forms.DialogResult.OK)
			{
				if (strType.Trim().ToUpper()=="NEW")
				{
                    if (ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection == null)
                    {
                        this.ReferenceFormRxItem.m_oRxItem.m_oHarvestCostColumnItem_Collection1 = new RxItemHarvestCostColumnItem_Collection();
                        this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection = this.ReferenceFormRxItem.m_oRxItem.m_oHarvestCostColumnItem_Collection1;
                    }
                    //make sure this item was not previously deleted
                    for (y = 0; y <= this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Count - 1; y++)
                    {
                        if (this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(y).Delete == true)
                        {
                            if (this.ReferenceFormRxItem.m_oRxItem.RxId == this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(y).RxId &&
                                this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(y).HarvestCostColumn.Trim().ToUpper() ==
                                frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText.Trim().ToUpper())
                            {
                                ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(y).Delete = false;
                                ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(y).Description =
                                     frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnDescription;
                                break;

                            }

                        }
                    }
                    //new column
                    if (y > this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Count - 1)
                    {
                        FIA_Biosum_Manager.RxItemHarvestCostColumnItem oItem = new RxItemHarvestCostColumnItem();
                        oItem.Description = frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnDescription.Trim();
                        oItem.HarvestCostColumn = frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText.Trim();
                        oItem.RxId = this.ReferenceFormRxItem.m_oRxItem.RxId;
                        oItem.Add = true;
                        ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Add(oItem);
                        this.AddItemToList(oItem);
                    }
                    

                   
					
				}
				else
				{
					this.lvRxHarvestCostColumns.SelectedItems[0].SubItems[COLUMN_DESC].Text = frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnDescription;

                    for (y = 0; y <= this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Count - 1; y++)
                    {
                        if (this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(y).Delete == false)
                        {
                               if (this.ReferenceFormRxItem.m_oRxItem.RxId == this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(y).RxId &&
                                this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(y).HarvestCostColumn.Trim().ToUpper() ==
                                lvRxHarvestCostColumns.SelectedItems[0].SubItems[COLUMN_FIELD].Text.Trim().ToUpper())
                            {
                                ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(y).Description =
                                     frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnDescription;
                                break;

                            }
                                
                        }
                    }

                    
				}
				
			}
			frmTemp.Dispose();
		}
        public void RemoveItem()
        {
            if (this.lvRxHarvestCostColumns.SelectedItems.Count == 0) return;
            int x;
            int y;
            int intSelect;
            int intCurrSelect;
            /**********************************************
             **lets see if we have one to remove
             **********************************************/
            int index = lvRxHarvestCostColumns.SelectedItems[0].Index;
            intSelect = index;
            

            //locate the current property associated with the listview
            for (x = 0; x <= ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Count - 1; x++)
            {
                if (this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(x).Delete == false)
                {
                    if (this.ReferenceFormRxItem.m_oRxItem.RxId == this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(x).RxId &&
                         this.ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(x).HarvestCostColumn.Trim().ToUpper() ==
                                lvRxHarvestCostColumns.SelectedItems[0].SubItems[COLUMN_FIELD].Text.Trim().ToUpper())
                    {
                        ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(x).Delete = true;
                        ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(x).Add = false;
                        ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(x).Index = -1;
                        break;
                    }

                }
               
            }
            /**********************************************
             **remove the ONE that is selected
             **********************************************/
            if (index == 0 && lvRxHarvestCostColumns.Items.Count == 1)
            {
                lvRxHarvestCostColumns.Items.Clear();
            }
            else
            {

                //*see if were at the top of the list
                if (index == 0 && lvRxHarvestCostColumns.Items.Count > 2)
                {
                    intSelect = 0;
                }
                else
                {
                    //*see if were at the bottom
                    if (index + 1 == lvRxHarvestCostColumns.Items.Count)
                    {
                        intCurrSelect = index - 1;
                        intSelect = index - 1;
                    }
                    else
                    {
                        intSelect = index;
                    }
                }
                lvRxHarvestCostColumns.Items.Remove(lvRxHarvestCostColumns.Items[index]);
            }

            if (lvRxHarvestCostColumns.Items.Count == 0)
            {
                
                ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_OPEN] = true;
                ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_NEW] = true;
                ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_CLEARALL] = false;
                ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_DELETE] = false;
                ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_EDIT] = false;
                ReferenceFormRxItem.SetToolBarButtonsEnabled(frmRxItem.UC_HARVESTCOST);
            }
            else
                this.lvRxHarvestCostColumns.Items[intSelect].Selected = true;

        }
        public void RemoveAllItems()
        {
            if (lvRxHarvestCostColumns.SelectedItems.Count == 0) return;

            this.lvRxHarvestCostColumns.Items.Clear();
            this.m_oLvAlternateRowColors.InitializeRowCollection();

            for (int x = ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Count - 1; x >= 0; x--)
            {
                ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(x).Delete = true;
                ReferenceFormRxItem.m_oRxItem.ReferenceHarvestCostColumnCollection.Item(x).Index = -1;
            }

            
            ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_OPEN] = true;
            ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_NEW] = true;
            ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_CLEARALL] = false;
            ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_DELETE] = false;
            ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_HARVESTCOST, frmRxItem.BUTTON_EDIT] = false;
            ReferenceFormRxItem.SetToolBarButtonsEnabled(frmRxItem.UC_HARVESTCOST);

        }

        private void lvRxHarvestCostColumns_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvRxHarvestCostColumns.SelectedItems.Count > 0)
                m_oLvAlternateRowColors.DelegateListViewItem(lvRxHarvestCostColumns.SelectedItems[0]);
        }

        private void lvRxHarvestCostColumns_MouseUp(object sender, MouseEventArgs e)
        {
           
            try
            {
                if (e.Button == MouseButtons.Left)
                {
                    int intRowHt = lvRxHarvestCostColumns.Items[0].Bounds.Height;
                    double dblRow = (double)(e.Y / intRowHt);
                    this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.TopItem.Index + (int)dblRow - 1].Selected = true;

                }
            }
            catch
            {
            }
        }
        

		
	}
}

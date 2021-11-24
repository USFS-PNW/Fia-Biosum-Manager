using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;


namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario_psite.
	/// </summary>
	public class uc_optimizer_scenario_select_packages : System.Windows.Forms.UserControl
	{
		private ListViewEmbeddedControls.ListViewEx lstRxPackages;
        private System.Windows.Forms.ImageList imgSize;
		private System.Windows.Forms.ListView listView1;
		private System.ComponentModel.IContainer components;
		private string m_strTempMDBFile;
		private env m_oEnv;
		public int m_intError=0;
		public string m_strError="";
		private System.Windows.Forms.Button btnSelectAll;
		private System.Windows.Forms.Button btnUnselectAll;
		private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ComboBox m_Combo;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario=null;
		private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvRowColors = new ListViewAlternateBackgroundColors();
        private ListViewColumnSorter lvwColumnSorter;

	    const int COLUMN_CHECKBOX=0;
		const int COLUMN_FVS_VARIANT=1;
		const int COLUMN_RXPACKAGE=2;


		public uc_optimizer_scenario_select_packages()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
            lstRxPackages = new ListViewEmbeddedControls.ListViewEx();

			this.groupBox1.Controls.Add(lstRxPackages);
		    lstRxPackages.Size = this.listView1.Size;
			lstRxPackages.Location = this.listView1.Location;
			lstRxPackages.CheckBoxes = true;
			lstRxPackages.AllowColumnReorder=true;
			lstRxPackages.FullRowSelect = false;
			lstRxPackages.GridLines = true;
			lstRxPackages.MultiSelect=false;
			lstRxPackages.View = System.Windows.Forms.View.Details;
			lstRxPackages.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(lstRxPackages_ColumnClick);
			lstRxPackages.MouseUp += new System.Windows.Forms.MouseEventHandler(lstRxPackages_MouseUp);
			lstRxPackages.SelectedIndexChanged += new System.EventHandler(lstRxPackages_SelectedIndexChanged);
			
			// 

			listView1.Hide();
			lstRxPackages.Show();

			if (frmMain.g_oGridViewFont != null) lstRxPackages.Font = frmMain.g_oGridViewFont;
			this.m_oLvRowColors.ReferenceListView = this.lstRxPackages;
			this.m_oLvRowColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvRowColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvRowColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvRowColors.CustomFullRowSelect=true;

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
        public void loadvalues_FromProperties(ProcessorScenarioItem oProcItem)
        {
            // For now we don't store selected fvs_variant/rxPackages at the scenario level
            this.loadvalues(oProcItem);
        }
		public void loadvalues(ProcessorScenarioItem oProcItem)
		{
			this.lstRxPackages.Clear();
			this.m_oLvRowColors.InitializeRowCollection();

		    this.lstRxPackages.Columns.Add("",2,HorizontalAlignment.Left);
			this.lstRxPackages.Columns.Add("Variant", 75, HorizontalAlignment.Left);
			this.lstRxPackages.Columns.Add("Package", 300, HorizontalAlignment.Left);

			this.lstRxPackages.Columns[COLUMN_CHECKBOX].Width = -2;

            // Create an instance of a ListView column sorter and assign it 
            // to the ListView control.
            lvwColumnSorter = new ListViewColumnSorter();
            this.lstRxPackages.ListViewItemSorter = lvwColumnSorter;
			
            m_oEnv = new env();
            ado_data_access oAdo = new ado_data_access();

            if (String.IsNullOrEmpty(oProcItem.ScenarioId))
            {
                return;
            }

            string[] arr1 = new string[] { "PLOT" };
            string strPlotTableName = "";
            object oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourceTableName", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                    strPlotTableName = strValue;
                }
            }
            string strHarvestCostsPathAndFile = oProcItem.DbPath + "\\" + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableDbFile;

            arr1 = new string[] { "CONDITION" };
            string strCondTableName = "";
            oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourceTableName", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                   strCondTableName = strValue;
                }
            }

            /**************************************************************
             **create a temporary MDB File that will contain table links
             **to the cond, plot, and harvest_costs tables
             **************************************************************/
            dao_data_access p_dao = new dao_data_access();
            if (System.IO.File.Exists(strHarvestCostsPathAndFile) && p_dao.TableExists(strHarvestCostsPathAndFile, 
                Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName))
            {
                this.m_strTempMDBFile = this.ReferenceOptimizerScenarioForm.uc_datasource1.CreateMDBAndScenarioTableDataSourceLinks(m_oEnv.strTempDir);

                p_dao.CreateTableLink(this.m_strTempMDBFile,
                                      Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName,
                                      strHarvestCostsPathAndFile, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName);
                p_dao = null;

                string strTempConn = oAdo.getMDBConnString(this.m_strTempMDBFile, "", "");
                using (var oConn = new OleDbConnection(strTempConn))
                {
                    oConn.Open();
                    oAdo.m_strSQL = "SELECT DISTINCT plot.fvs_variant, harvest_costs.rxpackage, Count(*) AS [Count]" +
                                    " FROM (" + strCondTableName + " INNER JOIN " + strPlotTableName +
                                    " ON " + strCondTableName + ".biosum_plot_id = " + strPlotTableName + ".biosum_plot_id) " +
                                    " INNER JOIN " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName +
                                    " ON " + strCondTableName + ".biosum_cond_id = " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + ".biosum_cond_id" +
                                    " GROUP BY FVS_VARIANT, RXPACKAGE";

                    oAdo.SqlQueryReader(oConn, oAdo.m_strSQL);
                    if (oAdo.m_OleDbDataReader.HasRows == true)
                    {
                        while (oAdo.m_OleDbDataReader.Read())
                        {
                            if (oAdo.m_OleDbDataReader["fvs_variant"] != System.DBNull.Value)
                            {
                                //null column
                                this.lstRxPackages.Items.Add(" ");
                                this.lstRxPackages.Items[lstRxPackages.Items.Count - 1].UseItemStyleForSubItems = false;
                                this.lstRxPackages.Items[lstRxPackages.Items.Count - 1].Checked = true;
                                this.m_oLvRowColors.AddRow();
                                this.m_oLvRowColors.AddColumns(lstRxPackages.Items.Count - 1, lstRxPackages.Columns.Count);

                                //fvs_variant
                                this.lstRxPackages.Items[lstRxPackages.Items.Count - 1].SubItems.Add(Convert.ToString(oAdo.m_OleDbDataReader["fvs_variant"]));
                                this.m_oLvRowColors.ListViewSubItem(lstRxPackages.Items.Count - 1,
                                    COLUMN_FVS_VARIANT,
                                    lstRxPackages.Items[lstRxPackages.Items.Count - 1].SubItems[COLUMN_FVS_VARIANT], false);

                                // rxPackage
                                if (oAdo.m_OleDbDataReader["rxpackage"] != System.DBNull.Value)
                                {
                                    this.lstRxPackages.Items[this.lstRxPackages.Items.Count - 1].SubItems.Add(oAdo.m_OleDbDataReader["rxpackage"].ToString());
                                }
                                else
                                {
                                    this.lstRxPackages.Items[this.lstRxPackages.Items.Count - 1].SubItems.Add(" ");
                                }
                                this.m_oLvRowColors.ListViewSubItem(lstRxPackages.Items.Count - 1,
                                    COLUMN_RXPACKAGE,
                                    lstRxPackages.Items[lstRxPackages.Items.Count - 1].SubItems[COLUMN_RXPACKAGE], false);
                            }
                        }
                    }
                    oAdo.m_OleDbDataReader.Close();
                }
            }
		}

		// We do not currently save the list of selected packages. It is reloaded
        // each time a scenario is loaded.
        public int savevalues()
		{
			string strScenarioId;
            int x = -1;

			ado_data_access p_ado = new ado_data_access();
			strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
                Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;

			string strConn = p_ado.getMDBConnString(strScenarioMDB,"admin","");
			p_ado.OpenConnection(strConn);	
			if (p_ado.m_intError != 0)
			{
				x=p_ado.m_intError;
				p_ado=null;
				return x;
			}
			
			return x;
		}
		public int val_rxPackages()
		{
			if (this.lstRxPackages.CheckedItems.Count == 0)
			{
                MessageBox.Show("Run Scenario Failed: Select at least one FVS variant + RxPackage combination in <Filter RxPackage>", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
				return -1;
			}
			return 0;

		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_optimizer_scenario_select_packages));
            this.imgSize = new System.Windows.Forms.ImageList(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnUnselectAll = new System.Windows.Forms.Button();
            this.btnSelectAll = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // imgSize
            // 
            this.imgSize.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgSize.ImageStream")));
            this.imgSize.TransparentColor = System.Drawing.Color.Transparent;
            this.imgSize.Images.SetKeyName(0, "");
            this.imgSize.Images.SetKeyName(1, "");
            // 
            // groupBox1
            // 
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.btnUnselectAll);
            this.groupBox1.Controls.Add(this.btnSelectAll);
            this.groupBox1.Controls.Add(this.listView1);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            this.toolTip1.SetToolTip(this.groupBox1, resources.GetString("groupBox1.ToolTip"));
            this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
            // 
            // lblTitle
            // 
            resources.ApplyResources(this.lblTitle, "lblTitle");
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Name = "lblTitle";
            this.toolTip1.SetToolTip(this.lblTitle, resources.GetString("lblTitle.ToolTip"));
            // 
            // btnUnselectAll
            // 
            resources.ApplyResources(this.btnUnselectAll, "btnUnselectAll");
            this.btnUnselectAll.BackColor = System.Drawing.SystemColors.Control;
            this.btnUnselectAll.ForeColor = System.Drawing.Color.Black;
            this.btnUnselectAll.Name = "btnUnselectAll";
            this.toolTip1.SetToolTip(this.btnUnselectAll, resources.GetString("btnUnselectAll.ToolTip"));
            this.btnUnselectAll.UseVisualStyleBackColor = false;
            this.btnUnselectAll.Click += new System.EventHandler(this.btnUnselectAll_Click);
            // 
            // btnSelectAll
            // 
            resources.ApplyResources(this.btnSelectAll, "btnSelectAll");
            this.btnSelectAll.BackColor = System.Drawing.SystemColors.Control;
            this.btnSelectAll.ForeColor = System.Drawing.Color.Black;
            this.btnSelectAll.Name = "btnSelectAll";
            this.toolTip1.SetToolTip(this.btnSelectAll, resources.GetString("btnSelectAll.ToolTip"));
            this.btnSelectAll.UseVisualStyleBackColor = false;
            this.btnSelectAll.Click += new System.EventHandler(this.btnSelectAll_Click);
            // 
            // listView1
            // 
            resources.ApplyResources(this.listView1, "listView1");
            this.listView1.AllowColumnReorder = true;
            this.listView1.CheckBoxes = true;
            this.listView1.GridLines = true;
            this.listView1.MultiSelect = false;
            this.listView1.Name = "listView1";
            this.toolTip1.SetToolTip(this.listView1, resources.GetString("listView1.ToolTip"));
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // uc_optimizer_scenario_select_package
            // 
            resources.ApplyResources(this, "$this");
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_optimizer_scenario_select_package";
            this.toolTip1.SetToolTip(this, resources.GetString("$this.ToolTip"));
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void grpboxPSites_Resize(object sender, System.EventArgs e)
		{
			this.lstRxPackages.Width = this.ClientSize.Width - (int)(this.lstRxPackages.Left * 2);
		}
		private void lstRxPackages_ColumnClick(object sender, System.Windows.Forms.ColumnClickEventArgs e)
		{
            int x, y;
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }
            // Perform the sort with these new sort options.
            this.lstRxPackages.Sort();
            //reinitialize the alternate row colors
            for (x = 0; x <= this.lstRxPackages.Items.Count - 1; x++)
            {
                for (y = 0; y <= this.lstRxPackages.Columns.Count - 1; y++)
                {
                    m_oLvRowColors.ListViewSubItem(this.lstRxPackages.Items[x].Index, y, this.lstRxPackages.Items[this.lstRxPackages.Items[x].Index].SubItems[y], false);
                }
            }
           
		}

		private void btnSelectAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=this.lstRxPackages.Items.Count-1;x++)
			{
				this.lstRxPackages.Items[x].Checked=true;
			}
		}

		private void btnUnselectAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=this.lstRxPackages.Items.Count-1;x++)
			{
				this.lstRxPackages.Items[x].Checked=false;
			}
		}

		private void lstRxPackages_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = lstRxPackages.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lstRxPackages.Items[lstRxPackages.TopItem.Index + (int)dblRow-1].Selected=true;
				}
			}
			catch 
			{
			}
		}
		private void lstRxPackages_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.lstRxPackages.SelectedItems.Count > 0)
				m_oLvRowColors.DelegateListViewItem(this.lstRxPackages.SelectedItems[0]);
		}

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			lstRxPackages.Width = this.ClientSize.Width - (lstRxPackages.Left * 2);
            this.btnSelectAll.Top = this.ClientSize.Height - (int)(this.btnSelectAll.Height * 1.5);
            this.btnUnselectAll.Top = this.btnSelectAll.Top;
            lstRxPackages.Height = this.ClientSize.Height - lstRxPackages.Top - (int)(this.btnSelectAll.Height * 1.5) - 3;
		}

        public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
        {
            get { return _frmScenario; }
            set { _frmScenario = value; }
        }

        public string RxPackagesNotSelected
        {
            get
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                for (int x = 0; x <= this.lstRxPackages.Items.Count - 1; x++)
                {
                    if (this.lstRxPackages.Items[x].Checked == false)
                    {
                        string strFvsVariant = this.lstRxPackages.Items[x].SubItems[COLUMN_FVS_VARIANT].Text.Trim();
                        string strRxPackage = this.lstRxPackages.Items[x].SubItems[COLUMN_RXPACKAGE].Text.Trim();
                        if (sb.Length == 0)
                        {
                            sb.Append(strFvsVariant + strRxPackage);
                        }
                        else
                        {
                            sb.Append("," + strFvsVariant + strRxPackage);
                        }
                    }
                }
                return sb.ToString();
            }
        }
		
	}

}

using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Text;
using SQLite.ADO;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_datasource.
	/// </summary>
	public class uc_datasource : System.Windows.Forms.UserControl
	{
		public System.Windows.Forms.ListView lstRequiredTables;
		public System.Windows.Forms.GroupBox groupBox1;
		private System.ComponentModel.IContainer components;
		public int intError;
		public System.Windows.Forms.Label lblTitle;
		//private dao_data_access p_DAO;
		public string strError;
		public string strTable;
		const int COLUMN_NULL = 0;
		const int TABLETYPE = 1;
		const int PATH = 2;
		const int DBFILE = 3;
		const int FILESTATUS = 4;
		const int TABLE = 5;
		const int TABLESTATUS = 6;
		const int RECORDCOUNT = 7;

		public string m_strRandomPathAndFile = "";
		public int m_intNumberOfValidTables=0;  //MDB file is FOUND and table is FOUND




		public string m_strProjectFile="";
		public string m_strProjectDirectory="";
		public string m_strScenarioId="";
		public string m_strScenarioFile="";
        private string m_strDataSourceDBFile;
		private System.Windows.Forms.ProgressBar progressBar1;
		private System.Windows.Forms.Label lblProgress;
		private string m_strDataSourceTable;
		private System.Windows.Forms.Panel panel1;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario=null;
		private FIA_Biosum_Manager.frmProcessorScenario _frmProcessorScenario=null;
		private string _strScenarioType="optimizer";

		private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvRowColors = new ListViewAlternateBackgroundColors();
		private System.Windows.Forms.ToolBar toolBar1;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ToolBarButton tlbBtnEdit;
		private System.Windows.Forms.ToolBarButton tlbBtnRefresh;
		private System.Windows.Forms.ToolBarButton toolBarButton1;
		private System.Windows.Forms.ToolBarButton tlbBtnClose;
		private System.Windows.Forms.ToolBarButton tlbBtnHelp;

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultDatabaseXPSFile;

		public FIA_Biosum_Manager.ResizeFormUsingVisibleScrollBars m_oResizeForm = new ResizeFormUsingVisibleScrollBars();

		private ListViewColumnSorter lvwColumnSorter;

		public uc_datasource()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.lstRequiredTables.View = System.Windows.Forms.View.Details;
			this.m_oLvRowColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvRowColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvRowColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceSelectedRowBackgroundColor=frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvRowColors.CustomFullRowSelect=true;
			this.m_oLvRowColors.ReferenceListView = lstRequiredTables;
			if (frmMain.g_oGridViewFont != null) this.lstRequiredTables.Font = frmMain.g_oGridViewFont;
			m_oResizeForm.ScrollBarParentControl=panel1;
			m_oResizeForm.ResizeWidth=false;
			m_oResizeForm.ResizeHeight=false;
			m_oResizeForm.MaximumHeight = 650;
            this.m_oEnv = new env();
			
			// TODO: Add any initialization after the InitializeComponent call

		}
      
		public uc_datasource(string p_strProjectDBFile)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_strDataSourceDBFile=p_strProjectDBFile;
			this.m_strDataSourceTable="datasource";
			this.lstRequiredTables.View = System.Windows.Forms.View.Details;
			this.lblTitle.Text = "Project Data Sources";
			this.lstRequiredTables.CheckBoxes=false;
			this.m_oLvRowColors.ReferenceListView=this.lstRequiredTables;
			this.m_oLvRowColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvRowColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvRowColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceSelectedRowBackgroundColor=frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvRowColors.CustomFullRowSelect=true;
			if (frmMain.g_oGridViewFont != null) this.lstRequiredTables.Font = frmMain.g_oGridViewFont;
            this.m_oEnv = new env();
			
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
		public void InitialSize()
		{
		}
		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(uc_datasource));
			this.toolBar1 = new System.Windows.Forms.ToolBar();
			this.lstRequiredTables = new System.Windows.Forms.ListView();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.panel1 = new System.Windows.Forms.Panel();
			this.progressBar1 = new System.Windows.Forms.ProgressBar();
			this.lblProgress = new System.Windows.Forms.Label();
			this.lblTitle = new System.Windows.Forms.Label();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.tlbBtnEdit = new System.Windows.Forms.ToolBarButton();
			this.tlbBtnRefresh = new System.Windows.Forms.ToolBarButton();
			this.toolBarButton1 = new System.Windows.Forms.ToolBarButton();
			this.tlbBtnClose = new System.Windows.Forms.ToolBarButton();
			this.tlbBtnHelp = new System.Windows.Forms.ToolBarButton();
			this.groupBox1.SuspendLayout();
			this.panel1.SuspendLayout();
			this.SuspendLayout();
			// 
			// toolBar1
			// 
			this.toolBar1.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
																						this.tlbBtnEdit,
																						this.tlbBtnRefresh,
																						this.toolBarButton1,
																						this.tlbBtnHelp,
																						this.tlbBtnClose});
			this.toolBar1.ButtonSize = new System.Drawing.Size(45, 40);
			this.toolBar1.DropDownArrows = true;
			this.toolBar1.ImageList = this.imageList1;
			this.toolBar1.Location = new System.Drawing.Point(0, 0);
			this.toolBar1.Name = "toolBar1";
			this.toolBar1.ShowToolTips = true;
			this.toolBar1.Size = new System.Drawing.Size(736, 46);
			this.toolBar1.TabIndex = 3;
			this.toolBar1.Click += new System.EventHandler(this.toolBar1_Click);
			this.toolBar1.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.toolBar1_ButtonClick);
			// 
			// lstRequiredTables
			// 
			this.lstRequiredTables.CheckBoxes = true;
			this.lstRequiredTables.GridLines = true;
			this.lstRequiredTables.HideSelection = false;
			this.lstRequiredTables.Location = new System.Drawing.Point(16, 32);
			this.lstRequiredTables.MultiSelect = false;
			this.lstRequiredTables.Name = "lstRequiredTables";
			this.lstRequiredTables.Size = new System.Drawing.Size(696, 400);
			this.lstRequiredTables.TabIndex = 1;
			this.lstRequiredTables.View = System.Windows.Forms.View.Details;
			this.lstRequiredTables.MouseDown += new System.Windows.Forms.MouseEventHandler(this.lstRequiredTables_MouseDown);
			this.lstRequiredTables.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstRequiredTables_MouseUp);
			this.lstRequiredTables.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.lstRequiredTables_ColumnClick);
			this.lstRequiredTables.SelectedIndexChanged += new System.EventHandler(this.lstRequiredTables_SelectedIndexChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.panel1);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Location = new System.Drawing.Point(0, 56);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(736, 488);
			this.groupBox1.TabIndex = 2;
			this.groupBox1.TabStop = false;
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.lstRequiredTables);
			this.panel1.Controls.Add(this.progressBar1);
			this.panel1.Controls.Add(this.lblProgress);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel1.Location = new System.Drawing.Point(3, 48);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(730, 437);
			this.panel1.TabIndex = 39;
			this.panel1.Resize += new System.EventHandler(this.panel1_Resize);
			// 
			// progressBar1
			// 
			this.progressBar1.Location = new System.Drawing.Point(256, 16);
			this.progressBar1.Name = "progressBar1";
			this.progressBar1.Size = new System.Drawing.Size(440, 8);
			this.progressBar1.TabIndex = 36;
			this.progressBar1.Visible = false;
			// 
			// lblProgress
			// 
			this.lblProgress.Location = new System.Drawing.Point(16, 8);
			this.lblProgress.Name = "lblProgress";
			this.lblProgress.Size = new System.Drawing.Size(239, 16);
			this.lblProgress.TabIndex = 35;
			this.lblProgress.Text = "lblProgress";
			this.lblProgress.Visible = false;
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(730, 32);
			this.lblTitle.TabIndex = 24;
			this.lblTitle.Text = "Scenario Data Source";
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(18, 18);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// tlbBtnEdit
			// 
			this.tlbBtnEdit.ImageIndex = 0;
			this.tlbBtnEdit.Text = "Edit";
			// 
			// tlbBtnRefresh
			// 
			this.tlbBtnRefresh.ImageIndex = 1;
			this.tlbBtnRefresh.Text = "Refresh";
			// 
			// toolBarButton1
			// 
			this.toolBarButton1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
			// 
			// tlbBtnClose
			// 
			this.tlbBtnClose.ImageIndex = 2;
			this.tlbBtnClose.Text = "Close";
			// 
			// tlbBtnHelp
			// 
			this.tlbBtnHelp.ImageIndex = 3;
			this.tlbBtnHelp.Text = "Help";
			// 
			// uc_datasource
			// 
			this.Controls.Add(this.toolBar1);
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_datasource";
			this.Size = new System.Drawing.Size(736, 552);
			this.Resize += new System.EventHandler(this.uc_datasource_Resize);
			this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.uc_datasource_MouseDown);
			this.groupBox1.ResumeLayout(false);
			this.panel1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void button1_Click(object sender, System.EventArgs e)
		{
			MessageBox.Show(this.lstRequiredTables.Columns[0].Width.ToString());
		}
        public void LoadValues()
        {
            LoadLstRequiredTables();
			if (m_strScenarioId != null && this.m_strScenarioId.Trim().Length > 0)
            {
				this.m_strDataSourceDBFile = this.m_strProjectDirectory + "\\" +
				this.ScenarioType + "\\db\\scenario_" + this.ScenarioType + "_rule_definitions.db";
			}
            DataMgr dataMgr = new DataMgr();
            string strConn = dataMgr.GetConnectionString(this.m_strDataSourceDBFile);
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                con.Open();
                if (m_strScenarioId != null && this.m_strScenarioId.Trim().Length > 0)
                {
                    dataMgr.m_strSQL = "select table_type,path,file,table_name from " + this.m_strDataSourceTable + " " +
                        " where trim(lower(scenario_id)) = '" +
                        m_strScenarioId.Trim().ToLower() + "';";
                }
                else
                {
                    dataMgr.m_strSQL = "select table_type,path,file,table_name from " + this.m_strDataSourceTable + ";";
                }
                try
                {
                    dataMgr.SqlQueryReader(con, dataMgr.m_strSQL);

                    System.Collections.Generic.IDictionary<string, string[]> dictSources =
                        new System.Collections.Generic.Dictionary<string, string[]>();
                    while (dataMgr.m_DataReader.Read())
                    {
                        if (dataMgr.m_DataReader["table_type"] != DBNull.Value &&
                            dataMgr.m_DataReader["table_type"].ToString().Trim().Length > 0)
                        {
                            string strKey = dataMgr.m_DataReader["table_type"].ToString().Trim();
                            if (!dictSources.ContainsKey(strKey))
                            {
                                string[] arrSource = new string[7];
                                arrSource[Datasource.TABLETYPE] = strKey;
                                arrSource[Datasource.PATH] = dataMgr.m_DataReader["path"].ToString().Trim();
                                arrSource[Datasource.DBFILE] = dataMgr.m_DataReader["file"].ToString().Trim();
                                arrSource[Datasource.TABLE] = dataMgr.m_DataReader["table_name"].ToString().Trim();
                                dictSources.Add(strKey, arrSource);
                            }
                        }
                    }
                    dataMgr.m_DataReader.Close();
                    Datasource oDatasource = new Datasource();
                    oDatasource.ValidateDataSources(ref dictSources);
                    oDatasource = null;

                    // Add a ListItem object to the ListView for each data source
                    int x = 0;
                    foreach (var tableType in dictSources.Keys)
                    {
                        string[] arrSource = dictSources[tableType];
                        ListViewItem entryListItem = lstRequiredTables.Items.Add(" ");
                        this.m_oLvRowColors.AddRow();
                        this.m_oLvRowColors.AddColumns(x, lstRequiredTables.Columns.Count);
                        entryListItem.UseItemStyleForSubItems = false;
                        this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_datasource.COLUMN_NULL, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                        this.lstRequiredTables.Items[x].SubItems.Add(tableType);
                        this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_datasource.TABLETYPE, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                        this.lstRequiredTables.Items[x].SubItems.Add(arrSource[Datasource.PATH]);
                        this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_datasource.PATH, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                        this.lstRequiredTables.Items[x].SubItems.Add(arrSource[Datasource.DBFILE]);
                        this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_datasource.DBFILE, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                        if (arrSource[Datasource.FILESTATUS] == "F")
                        {
                            ListViewItem.ListViewSubItem FileStatusSubItem =
                                entryListItem.SubItems.Add("Found");

                            this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_datasource.FILESTATUS, FileStatusSubItem, false);

                            FileStatusSubItem.Font = frmMain.g_oGridViewFont;

                            this.lstRequiredTables.Items[x].SubItems.Add(arrSource[Datasource.TABLE]);
                            this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_datasource.TABLE, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);

                            //see if the table exists in the database container
                            if (arrSource[Datasource.TABLESTATUS] != null &&
                               arrSource[Datasource.TABLESTATUS] == "F")
                            {
                                this.lstRequiredTables.Items[x].SubItems.Add("Found");
                                this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, TABLESTATUS, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                                this.lstRequiredTables.Items[x].SubItems.Add(arrSource[Datasource.RECORDCOUNT]);
                                this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, RECORDCOUNT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);

                            }
                            else
                            {
                                ListViewItem.ListViewSubItem TableStatusSubItem =
                                    entryListItem.SubItems.Add("Not Found");
                                TableStatusSubItem.ForeColor = System.Drawing.Color.White;
                                TableStatusSubItem.BackColor = System.Drawing.Color.Red;
                                this.m_oLvRowColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(uc_datasource.TABLESTATUS).UpdateColumn = false;
                                TableStatusSubItem.Font = frmMain.g_oGridViewFont;
                                this.lstRequiredTables.Items[x].SubItems.Add("0");
                                this.m_oLvRowColors.ListViewSubItem(x, uc_datasource.RECORDCOUNT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                            }
                        }
                        else
                        {
                            ListViewItem.ListViewSubItem FileStatusSubItem =
                                entryListItem.SubItems.Add("Not Found");
                            FileStatusSubItem.ForeColor = System.Drawing.Color.White;
                            FileStatusSubItem.BackColor = System.Drawing.Color.Red;
                            this.m_oLvRowColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(uc_datasource.FILESTATUS).UpdateColumn = false;

                            FileStatusSubItem.Font = frmMain.g_oGridViewFont;
                            this.lstRequiredTables.Items[x].SubItems.Add(arrSource[Datasource.TABLE]);
                            ListViewItem.ListViewSubItem TableStatusSubItem =
                                entryListItem.SubItems.Add("Not Found");
                            TableStatusSubItem.ForeColor = System.Drawing.Color.White;
                            TableStatusSubItem.BackColor = System.Drawing.Color.Red;
                            this.m_oLvRowColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(uc_datasource.TABLE).UpdateColumn = false;
                            TableStatusSubItem.Font = frmMain.g_oGridViewFont;
                            this.lstRequiredTables.Items[x].SubItems.Add("0");
                            this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_datasource.RECORDCOUNT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                        }
                        Datasource.UpdateTableMacroVariable(entryListItem.SubItems[TABLETYPE].Text, entryListItem.SubItems[TABLE].Text);
                        this.lstRequiredTables.Items[x].SubItems.Add(Datasource.g_oCurrentSQLMacroSubstitutionVariableItem.VariableName);
                        this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, DBFILE, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                        x++;
                    }


                    }
                    catch
                    {
                        intError = -1;
                        strError = "Failed to load data source data from " + this.strDataSourceMDBFile;
                        MessageBox.Show(strError, "FIA Biosum");
                        return;
                    }
            }
            intError = 0;
            strError = "";
        }

        private void LoadLstRequiredTables()
        {
            this.lstRequiredTables.Clear();
            this.lstRequiredTables.Columns.Add(" ", 2, HorizontalAlignment.Left);
            this.lstRequiredTables.Columns.Add("Table Type", 50, HorizontalAlignment.Left);
            this.lstRequiredTables.Columns.Add("Path", 60, HorizontalAlignment.Left);
            this.lstRequiredTables.Columns.Add("DB", 60, HorizontalAlignment.Left);
            this.lstRequiredTables.Columns.Add("File Status", 80, HorizontalAlignment.Left);
            this.lstRequiredTables.Columns.Add("Table Name", 50, HorizontalAlignment.Left);
            this.lstRequiredTables.Columns.Add("Table Status", 80, HorizontalAlignment.Left);
            this.lstRequiredTables.Columns.Add("Record Count", 80, HorizontalAlignment.Left);
            this.lstRequiredTables.Columns.Add("Table Macro Variable Name", 150, HorizontalAlignment.Left);

            // Create an instance of a ListView column sorter and assign it 
            // to the ListView control.
            lvwColumnSorter = new ListViewColumnSorter();
            this.lstRequiredTables.ListViewItemSorter = lvwColumnSorter;
        }


		private void uc_datasource_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (this.m_strScenarioId.Trim().Length > 0)
			{
				if (this.ScenarioType.Trim().ToUpper()=="OPTIMIZER")
					((frmOptimizerScenario)this.ParentForm).m_bPopup = false;
				else
					this.ReferenceProcessorScenarioForm.m_bPopup=false;
			}
		}

		private void lstRequiredTables_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (this.m_strScenarioId.Trim().Length > 0)
			{
                if (this.ScenarioType.Trim().ToUpper() == "OPTIMIZER")
				 ((frmOptimizerScenario)this.ParentForm).m_bPopup = false;
				else
				  this.ReferenceProcessorScenarioForm.m_bPopup=false;
			}
		}

		
		private void uc_datasource_Resize(object sender, System.EventArgs e)
		{
			this.resize_uc_datasource();

		
		}
		public void resize_uc_datasource()
		{
			this.groupBox1.Width = this.ClientSize.Width - (this.groupBox1.Left * 2);
		    this.groupBox1.Height = this.ClientSize.Height - (this.groupBox1.Top);
			this.lstRequiredTables.Width = this.groupBox1.Width - (this.lstRequiredTables.Left * 2);
			this.lstRequiredTables.Height = this.groupBox1.Height - this.groupBox1.Top - this.lstRequiredTables.Top - 30;

		}

		private void btnUndo_Click(object sender, System.EventArgs e)
		{
			DialogResult result = MessageBox.Show("All your changes will be undone and replaced with data source values in the database?(y/n)", "Data Sources", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			switch (result)
			{
				case DialogResult.Yes :
                        this.LoadValues();
				    break;
			}
		}

		private void EditDatasource()
        {
			if (this.lstRequiredTables.SelectedItems.Count == 0)
			{
				return;
			}
			string strConn = "";
			string strSQL = "";
			int y;
			DataMgr p_dataMgr = new DataMgr();

			string strDBFullPath = this.lstRequiredTables.SelectedItems[0].SubItems[PATH].Text.Trim() + "\\" +
				this.lstRequiredTables.SelectedItems[0].SubItems[DBFILE].Text.Trim();
			string strTable = this.lstRequiredTables.SelectedItems[0].SubItems[TABLE].Text.Trim();

			FIA_Biosum_Manager.uc_datasource_edit p_uc;
			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog((frmMain)this.ParentForm.ParentForm);
			frmTemp.MaximizeBox = false;
			frmTemp.BackColor = System.Drawing.SystemColors.Control;

			if (this.m_strScenarioId.Trim().Length > 0)
			{
				if (ScenarioType.Trim().ToUpper() == "OPTIMIZER")
				{
					frmTemp.Text = "Treatment Optimizer: Edit " + this.lstRequiredTables.SelectedItems[0].SubItems[TABLETYPE].Text.Trim() + " Data Source";
					p_uc = new uc_datasource_edit(this.m_strDataSourceDBFile, this.m_strScenarioId);
				}
				else
				{
					frmTemp.Text = "Prcoessor: Edit " + this.lstRequiredTables.SelectedItems[0].SubItems[TABLETYPE].Text.Trim() + " Data Source";
					p_uc = new uc_datasource_edit(this.m_strDataSourceDBFile, this.m_strScenarioId);
				}
			}
			else
			{
				frmTemp.Text = "Database: Edit " + this.lstRequiredTables.SelectedItems[0].SubItems[TABLETYPE].Text.Trim() + " Data Source";
				p_uc = new uc_datasource_edit(this.m_strDataSourceDBFile);
			}

			frmTemp.Controls.Add(p_uc);

			p_uc.strProjectDirectory = this.strProjectDirectory;
			p_uc.strScenarioId = this.strScenarioId;

			frmTemp.Height = 0;
			frmTemp.Width = 0;
			if (p_uc.Top + p_uc.Height > frmTemp.ClientSize.Height + 2)
			{
				for (y = 1; ; y++)
				{
					frmTemp.Height = y;
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
				for (y = 1; ; y++)
				{
					frmTemp.Width = y;
					if (p_uc.Left +
						p_uc.Width <
						frmTemp.ClientSize.Width)
					{
						break;
					}
				}

			}
			frmTemp.Left = 0;
			frmTemp.Top = 0;

			p_uc.lblDBFile.Text = strDBFullPath.Trim();
			p_uc.lblTable.Text = strTable.Trim();
			p_uc.lblTableType.Text =
				this.lstRequiredTables.SelectedItems[0].SubItems[TABLETYPE].Text.Trim();
			p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

			System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
			if (result == System.Windows.Forms.DialogResult.OK)
			{
				utils p_utils = new utils();
				string strDir = p_utils.getDirectory(p_uc.lblNewDBFile.Text);
				string strFile = p_utils.getFileName(p_uc.lblNewDBFile.Text);
				this.lstRequiredTables.SelectedItems[0].SubItems[PATH].Text = strDir;
				this.lstRequiredTables.SelectedItems[0].SubItems[DBFILE].Text = strFile;
				ListViewItem.ListViewSubItem FileSubItem =
					this.lstRequiredTables.SelectedItems[0].SubItems[FILESTATUS];


				FileSubItem.Font = frmMain.g_oGridViewFont;
				this.lstRequiredTables.SelectedItems[0].SubItems[FILESTATUS].Text = "Found";
				this.m_oLvRowColors.m_oRowCollection.Item(this.lstRequiredTables.SelectedItems[0].Index).m_oColumnCollection.Item(FILESTATUS).UpdateColumn = true;
				this.m_oLvRowColors.ListViewSubItem(lstRequiredTables.SelectedItems[0].Index, FILESTATUS, FileSubItem, true);
				this.lstRequiredTables.SelectedItems[0].SubItems[TABLE].Text = p_uc.lblNewTable.Text;
				ListViewItem.ListViewSubItem TableSubItem =
					this.lstRequiredTables.SelectedItems[0].SubItems[TABLESTATUS];


				TableSubItem.Font = frmMain.g_oGridViewFont;
				this.lstRequiredTables.SelectedItems[0].SubItems[TABLESTATUS].Text = "Found";
				this.m_oLvRowColors.m_oRowCollection.Item(this.lstRequiredTables.SelectedItems[0].Index).m_oColumnCollection.Item(TABLESTATUS).UpdateColumn = true;
				this.m_oLvRowColors.ListViewSubItem(lstRequiredTables.SelectedItems[0].Index, TABLESTATUS, TableSubItem, true);
				p_utils = null;
				strConn = p_dataMgr.GetConnectionString(p_uc.lblNewDBFile.Text.Trim());
				using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
					conn.Open();
					p_dataMgr.m_strSQL = "SELECT COUNT(*) FROM " + p_uc.lblNewTable.Text.Trim();
					this.lstRequiredTables.SelectedItems[0].SubItems[RECORDCOUNT].Text =
						Convert.ToString(p_dataMgr.getRecordCount(conn, p_dataMgr.m_strSQL, p_uc.lblNewTable.Text.Trim()));
                }
			}
			frmTemp.Close();
			frmTemp.Dispose();
			frmTemp = null;
			p_dataMgr = null;
		}
		
		private void CloseDatasource()
		{
			if (this.m_strScenarioId.Trim().Length > 0)
			{
				this.Visible=false;
                if (this.ScenarioType.Trim().ToUpper() == "OPTIMIZER")
					((frmOptimizerScenario)this.ParentForm).Height = 0 ; 
				 else this.ReferenceProcessorScenarioForm.Height=0;
				
			}
			else
			{
				this.ParentForm.Dispose();
			}
		}
		private void btnClose_Click(object sender, System.EventArgs e)
		{
			if (this.m_strScenarioId.Trim().Length > 0)
			{
				this.Visible=false;
			
				((frmOptimizerScenario)this.ParentForm).Height = 0 ; 
				
			}
			else
			{
				this.ParentForm.Dispose();
			}
		}
		public void SetEnableToolbarButton(string p_strButtonText,bool p_bEnable)
		{
			for (int x=0;x<=toolBar1.Buttons.Count-1;x++)
			{
				if (toolBar1.Buttons[x].Text.Trim().ToUpper() == p_strButtonText.Trim().ToUpper())
				{
					toolBar1.Buttons[x].Enabled=p_bEnable;
					break;
				}
			}
		}
		public bool ScenarioDataSourceTableExist(string strTableName)
		{
			int x;
			for (x=0; x <= this.lstRequiredTables.Items.Count - 1; x++)
			{
				if (this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim().ToUpper()==strTableName.Trim().ToUpper())
				{
					if (this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper()=="NOT FOUND")
					{
						return false;
					}
					if (this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper()=="NOT FOUND")
					{
						return false;
					}
					return true;
				}   
			}
      
			return false;



		}
		
		 
		/********************************************************
		 ** return the row associated with the table type
		 ********************************************************/
		public int getDataSourceTableNameRow(string pcTableId)
		{
			int x;
			for (x=0; x<= this.lstRequiredTables.Items.Count-1;x++)
			{
				if (pcTableId.Trim().ToUpper() == 
					this.lstRequiredTables.Items[x].SubItems[TABLETYPE].Text.Trim().ToUpper())
				{
					return x;
				}
			}
			return -1;
		}
		/********************************************************
		 ** return the table name associated with the table type
		 ********************************************************/
		public string getDataSourceTableName(string pcTableId)
		{
			int x;
			for (x=0; x<= this.lstRequiredTables.Items.Count-1;x++)
			{
				if (pcTableId.Trim().ToUpper() == 
					this.lstRequiredTables.Items[x].SubItems[TABLETYPE].Text.Trim().ToUpper()
					&&
					this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper() =="FOUND" 
					&&
					this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper() == "FOUND")
				{
					return this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim();
				}
			}
			return "";
		}
        /********************************************************
        ** return the database path, including database name, associated with the table type
        ********************************************************/
        public string getDataSourcePathAndFile(string pcTableId)
        {
            int x;
            for (x = 0; x <= this.lstRequiredTables.Items.Count - 1; x++)
            {
                if (pcTableId.Trim().ToUpper() ==
                    this.lstRequiredTables.Items[x].SubItems[TABLETYPE].Text.Trim().ToUpper()
                    &&
                    this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper() == "FOUND"
                    &&
                    this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper() == "FOUND")
                {
                    string strPathAndFile = this.lstRequiredTables.Items[x].SubItems[PATH].Text.Trim() + "\\" +
                        this.lstRequiredTables.Items[x].SubItems[DBFILE].Text.Trim();
                    return strPathAndFile;
                }
            }
            return "";
        }
		public int val_datasources()
		{
			

            int x=0;
			for (x=0; x <= this.lstRequiredTables.Items.Count - 1; x++)
			{
				if (this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper()=="NOT FOUND")
				{
					MessageBox.Show("Run Scenario Failed: Scenario data source file " + this.lstRequiredTables.Items[x].SubItems[PATH].Text.Trim() + "\\" + 
						            this.lstRequiredTables.Items[x].SubItems[DBFILE].Text.Trim() + " is not found");
					return -1;
				}
				if (this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper()=="NOT FOUND")
				{
					MessageBox.Show("Run Scenario Failed: Scenario data source table " + this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim() + 
						 " is not found");
					return -1;
				}
				if (this.lstRequiredTables.Items[x].SubItems[RECORDCOUNT].Text.Trim().ToUpper()=="0")
				{
					//the table below does not have to have records,whereas, all
					//the other tables are required to have records
					if (this.lstRequiredTables.Items[x].SubItems[TABLETYPE].Text.Trim().ToUpper() != "TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES")
					{
						MessageBox.Show("Run Scenario Failed: Scenario data source table " + this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim() + 
							" has 0 records");
						return -1;
					}
				}
			}
      
			return 0;
		}

		private void btnClearAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=this.lstRequiredTables.Items.Count-1;x++)
			{
				this.lstRequiredTables.Items[x].Checked=false;
			}
																	
		}

		private void btnCheckAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=this.lstRequiredTables.Items.Count-1;x++)
			{
				this.lstRequiredTables.Items[x].Checked=true;
			}

		}

		private void panel1_Resize(object sender, System.EventArgs e)
		{
			
			
		}

		private void lstRequiredTables_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = lstRequiredTables.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lstRequiredTables.Items[lstRequiredTables.TopItem.Index + (int)dblRow-1].Selected=true;
					
				}
			}
			catch 
			{
			}
		}

		private void lstRequiredTables_SelectedIndexChanged(object sender, System.EventArgs e)
		{
            if (lstRequiredTables.SelectedItems.Count > 0)
            {
                this.m_oLvRowColors.DelegateListViewItem(lstRequiredTables.SelectedItems[0]);
                string strSelectedPath = lstRequiredTables.SelectedItems[0].SubItems[PATH].Text.Trim();
                if (!String.IsNullOrEmpty(strSelectedPath))
                {
                    // Disable Edit button if this is an AppData data source
                    String strAppData = "@@AppData@@".ToUpper();
                    if (strSelectedPath.Trim().ToUpper().IndexOf(strAppData) > -1)
                    {
                        tlbBtnEdit.Enabled = false;
                        return;
                    }
                }

            }
            tlbBtnEdit.Enabled = true;
		}

		private void toolBar1_Click(object sender, System.EventArgs e)
		{
			
		}

		private void toolBar1_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			switch (e.Button.Text)
			{
				case "Close":
					this.CloseDatasource();
					break;
				case "Edit":
					this.EditDatasource();
					break;
				case "Refresh":
                    this.LoadValues();
					break;
                case "Help":
                    this.showHelp();
                    break;
			}
		}

		private void lstRequiredTables_ColumnClick(object sender, System.Windows.Forms.ColumnClickEventArgs e)
		{
			int x,y;
			
			// Determine if clicked column is already the column that is being sorted.
			if ( e.Column == lvwColumnSorter.SortColumn )
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
			this.lstRequiredTables.Sort();
			//reinitialize the alternate row colors
			for (x=0;x<=this.lstRequiredTables.Items.Count-1;x++)
			{
				for (y=0;y<=this.lstRequiredTables.Columns.Count-1;y++)
				{
					this.m_oLvRowColors.ListViewSubItem(this.lstRequiredTables.Items[x].Index,y,this.lstRequiredTables.Items[this.lstRequiredTables.Items[x].Index].SubItems[y],false);
				}
			}
			

		}
		
	
		public string strScenarioId
		{
			set 
			{
				this.m_strScenarioId = value;
				
			}
			get
			{
				return this.m_strScenarioId;
			}

		}
		public string strDataSourceMDBFile
		{
			set
			{
				this.m_strDataSourceDBFile=value;
			}
			get
			{
				return this.m_strDataSourceDBFile;
			}
		}
		public string strDataSourceTable
		{
			set
			{
				this.m_strDataSourceTable=value;
			}
			get
			{
				return this.m_strDataSourceTable;
			}
		}
		public string strProjectDirectory
		{
			set
			{
				this.m_strProjectDirectory=value;
			}
			get
			{
				return this.m_strProjectDirectory;
			}
		}
        public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
		public FIA_Biosum_Manager.frmProcessorScenario ReferenceProcessorScenarioForm
		{
			get {return _frmProcessorScenario;}
			set {_frmProcessorScenario=value;}
		}
		public string ScenarioType
		{
			get {return _strScenarioType;}
			set {_strScenarioType=value;}
		}

        public void showHelp()
        {
            string strParent = "";
            string strChild = "";
            if (this.lblTitle.Text == ("Project Data Sources"))
            {
                m_xpsFile = Help.DefaultDatabaseXPSFile;
                strParent = "DATABASE";
                strChild = "PROJECT_DATA_SOURCES";
            }
            else if (this.ScenarioType.Trim().ToUpper() == "PROCESSOR")
            {
                m_xpsFile = Help.DefaultProcessorXPSFile;
                strParent = "PROCESSOR";
                strChild = "PROCESSOR_DATA_SOURCES";
            }
            else if (this.ScenarioType.Trim().ToUpper() == "OPTIMIZER")
            {
                m_xpsFile = Help.DefaultTreatmentOptimizerFile;
                strParent = "TREATMENT_OPTIMIZER";
                strChild = "DATA_SOURCES";
            }
            
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            if (!String.IsNullOrEmpty(strParent))
                m_oHelp.ShowHelp(new string[] { strParent, strChild });
        }
		
		
      
	}
}

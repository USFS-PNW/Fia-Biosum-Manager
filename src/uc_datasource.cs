using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Text;

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
		const int MDBFILE = 3;
		const int FILESTATUS = 4;
		const int TABLE = 5;
		const int TABLESTATUS = 6;
		const int RECORDCOUNT = 7;
		const int MACRONAME=8;

		public string m_strRandomPathAndFile = "";
		public int m_intNumberOfValidTables=0;  //MDB file is FOUND and table is FOUND




		public string m_strProjectFile="";
		public string m_strProjectDirectory="";
		public string m_strScenarioId="";
		public string m_strScenarioFile="";
        private bool m_bUsingSqlite = false;
        private string m_strDataSourceMDBFile;
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
      
		public uc_datasource(string p_strProjectMDBFile)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_strDataSourceMDBFile=p_strProjectMDBFile;
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


		public uc_datasource(string p_strScenarioMDBFile, string p_strScenarioId)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.lstRequiredTables.View = System.Windows.Forms.View.Details;
			this.m_strDataSourceMDBFile=p_strScenarioMDBFile;
			this.m_strDataSourceTable="scenario_datasource";
			this.m_strScenarioId = p_strScenarioId;
			this.lblTitle.Text = "Scenario Data Sources";
			this.m_oLvRowColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvRowColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvRowColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceSelectedRowBackgroundColor=frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvRowColors.CustomFullRowSelect=true;
			this.m_oLvRowColors.ReferenceListView = lstRequiredTables;
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
            macrosubst oMacroSub = new macrosubst();
            oMacroSub.ReferenceGeneralMacroSubstitutionVariableCollection = frmMain.g_oGeneralMacroSubstitutionVariable_Collection;
            
            string strPathAndFile="";
			string strSQL="";
			string strConn="";

            LoadLstRequiredTables();

			ado_data_access p_ado = new ado_data_access();
			System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
			strConn = p_ado.getMDBConnString(this.m_strDataSourceMDBFile,"","");
			oConn.ConnectionString = strConn;
			intError = 0;
			strError = "";
			try
			{
				oConn.Open();
			}
			catch (System.Data.OleDb.OleDbException oleException)
			{
				strError = "Failed to make an oleDb connection with " + strConn;
				MessageBox.Show (strError + " oleException=" + oleException.Message.Trim());
				intError = -1;
				return;
			}
			System.Data.OleDb.OleDbCommand oCommand = oConn.CreateCommand();
			if (m_strScenarioId != null && this.m_strScenarioId.Trim().Length  > 0)
			{
				oCommand.CommandText = "select table_type,path,file,table_name from " + this.m_strDataSourceTable + " " + 
					" where scenario_id = '" + 
					m_strScenarioId.Trim() +  "';";
			}
			else
			{
				oCommand.CommandText = "select table_type,path,file,table_name from " + this.m_strDataSourceTable + ";";

			}
			try
			{
				System.Data.OleDb.OleDbDataReader oDataReader = oCommand.ExecuteReader();
				int x = 0;
				
				while (oDataReader.Read())
				{
					if (oDataReader["table_type"] != System.DBNull.Value &&
						oDataReader["table_type"].ToString().Trim().Length > 0)
					{
						// Add a ListItem object to the ListView.
						System.Windows.Forms.ListViewItem entryListItem =
							lstRequiredTables.Items.Add(" ");
						this.m_oLvRowColors.AddRow();
						this.m_oLvRowColors.AddColumns(x,lstRequiredTables.Columns.Count);
						//System.Windows.Forms.ListViewItem entryListItem =
						//	lstRequiredTables.Items.Add(oDataReader["table_type"].ToString());
						entryListItem.UseItemStyleForSubItems=false;
						this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_datasource.COLUMN_NULL,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["table_type"].ToString());
						this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_datasource.TABLETYPE,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["path"].ToString());
						this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_datasource.PATH,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["file"].ToString());
						this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_datasource.MDBFILE,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
                        strPathAndFile = oMacroSub.GeneralTranslateVariableSubstitution(oDataReader["path"].ToString().Trim()) + 
                            "\\" + oDataReader["file"].ToString().Trim();
						if (System.IO.File.Exists(strPathAndFile) == true) 
						{
							ListViewItem.ListViewSubItem FileStatusSubItem = 
								entryListItem.SubItems.Add("Found");

							this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_datasource.FILESTATUS,FileStatusSubItem,false);
							//FileStatusSubItem.ForeColor = System.Drawing.Color.Black;
							//FileStatusSubItem.BackColor = System.Drawing.Color.White;
							//FileStatusSubItem.Font = new System.Drawing.Font(
							//	"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);

							FileStatusSubItem.Font = frmMain.g_oGridViewFont;
							//FileStatusSubItem.ForeColor = frmMain.g_oGridViewRowForegroundColor;
							//FileStatusSubItem.BackColor = frmMain.g_oGridViewRowBackgroundColor;


							this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["table_name"].ToString());
							this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_datasource.TABLE,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);

							//see if the table exists in the mdb database container
							dao_data_access p_dao = new dao_data_access();
							if (p_dao.TableExists(strPathAndFile,oDataReader["table_name"].ToString().Trim()) == true)
							{
								this.lstRequiredTables.Items[x].SubItems.Add("Found");
								this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_datasource.TABLESTATUS,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
								strConn = p_ado.getMDBConnString(strPathAndFile,"admin","");   //"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strPathAndFile + ";User Id=admin;Password=;";
								strSQL = "select count(*) from " + oDataReader["table_name"].ToString();
								this.lstRequiredTables.Items[x].SubItems.Add(Convert.ToString(p_ado.getRecordCount(strConn,strSQL,oDataReader["table_name"].ToString())));
								this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_datasource.RECORDCOUNT,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);

							}
							else 
							{
								ListViewItem.ListViewSubItem TableStatusSubItem = 
									entryListItem.SubItems.Add("Not Found");
								TableStatusSubItem.ForeColor = System.Drawing.Color.White;
								TableStatusSubItem.BackColor = System.Drawing.Color.Red;
								this.m_oLvRowColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(uc_datasource.TABLESTATUS).UpdateColumn=false;
								//TableStatusSubItem.Font = new System.Drawing.Font(
								//	"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
								TableStatusSubItem.Font = frmMain.g_oGridViewFont;
								this.lstRequiredTables.Items[x].SubItems.Add("0");
								this.m_oLvRowColors.ListViewSubItem(x,uc_datasource.RECORDCOUNT,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
							}
							
							
							p_dao = null;
						}
						else 
						{
							ListViewItem.ListViewSubItem FileStatusSubItem = 
								entryListItem.SubItems.Add("Not Found");
							FileStatusSubItem.ForeColor = System.Drawing.Color.White;
							FileStatusSubItem.BackColor = System.Drawing.Color.Red;
							this.m_oLvRowColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(uc_datasource.FILESTATUS).UpdateColumn=false;

							//FileStatusSubItem.Font = new System.Drawing.Font(
							//	"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
							FileStatusSubItem.Font = frmMain.g_oGridViewFont;
							this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["table_name"].ToString());
							ListViewItem.ListViewSubItem TableStatusSubItem = 
								entryListItem.SubItems.Add("Not Found");
							TableStatusSubItem.ForeColor = System.Drawing.Color.White;
							TableStatusSubItem.BackColor = System.Drawing.Color.Red;
							this.m_oLvRowColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(uc_datasource.TABLE).UpdateColumn=false;
							//TableStatusSubItem.Font = new System.Drawing.Font(
							//	"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
							TableStatusSubItem.Font = frmMain.g_oGridViewFont;
							this.lstRequiredTables.Items[x].SubItems.Add("0");
							this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,uc_datasource.RECORDCOUNT,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);

						}
						Datasource.UpdateTableMacroVariable(entryListItem.SubItems[TABLETYPE].Text,entryListItem.SubItems[TABLE].Text);
						this.lstRequiredTables.Items[x].SubItems.Add(Datasource.g_oCurrentSQLMacroSubstitutionVariableItem.VariableName);
						this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,MDBFILE,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						x++;
					}
					
				}
				oDataReader.Close();
			}
			catch
			{
				intError = -1;
				strError = "The Query Command " + oCommand.CommandText.ToString() + " Failed";
				MessageBox.Show(strError);
				oConn.Close();
				p_ado= null;
				return;
			}
			
			oConn.Close();
			

			this.lstRequiredTables.Columns[TABLETYPE].Width = -1;
			


			p_ado = null;

		}

        public void LoadValuesSqlite()
        {
            LoadLstRequiredTables();

            this.m_strDataSourceMDBFile = this.m_strProjectDirectory + "\\" +
                this.ScenarioType + "\\db\\scenario_" + this.ScenarioType + "_rule_definitions.db";
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(this.m_strDataSourceMDBFile);
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                con.Open();
                if (m_strScenarioId != null && this.m_strScenarioId.Trim().Length > 0)
                {
                    dataMgr.m_strSQL = "select table_type,path,file,table_name from " + this.m_strDataSourceTable + " " +
                        " where trim(scenario_id) = '" +
                        m_strScenarioId.Trim() + "';";
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
                                arrSource[Datasource.MDBFILE] = dataMgr.m_DataReader["file"].ToString().Trim();
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
                        this.lstRequiredTables.Items[x].SubItems.Add(arrSource[Datasource.MDBFILE]);
                        this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_datasource.MDBFILE, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
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
                        this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, MDBFILE, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
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
            this.lstRequiredTables.Columns.Add("MDB File", 60, HorizontalAlignment.Left);
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

        public void populate_listview_grid()
		{

            macrosubst oMacroSub = new macrosubst();
            oMacroSub.ReferenceGeneralMacroSubstitutionVariableCollection = frmMain.g_oGeneralMacroSubstitutionVariable_Collection;

			string strPathAndFile="";
			string strSQL="";
			string strConn="";


            //@ToDo: Consider using LoadLstRequiredTables() here
            this.lstRequiredTables.Clear();
			this.m_oLvRowColors.InitializeRowCollection();

            ado_data_access p_ado = new ado_data_access();
 
			this.lstRequiredTables.Columns.Add(" ",2,HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Table Type", 50, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Path", 60, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("MDB File", 60, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("File Status", 80, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Table Name", 50, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Table Status", 80, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Record Count", 80, HorizontalAlignment.Left);
			this.lstRequiredTables.Columns.Add("Table Macro Variable Name",150,HorizontalAlignment.Left);

			// Create an instance of a ListView column sorter and assign it 
			// to the ListView control.
			lvwColumnSorter = new ListViewColumnSorter();
			this.lstRequiredTables.ListViewItemSorter = lvwColumnSorter;


			this.m_oLvRowColors.InitializeRowCollection();

			System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
			strConn = p_ado.getMDBConnString(this.m_strDataSourceMDBFile,"","");
			oConn.ConnectionString = strConn;
			intError = 0;
			strError = "";
			try
			{
				oConn.Open();
			}
			catch (System.Data.OleDb.OleDbException oleException)
			{
				strError = "Failed to make an oleDb connection with " + strConn;
				MessageBox.Show (strError + " oleException=" + oleException.Message.Trim());
				intError = -1;
				return;
			}
			System.Data.OleDb.OleDbCommand oCommand = oConn.CreateCommand();
			if (this.m_strScenarioId.Trim().Length  > 0)
			{
				oCommand.CommandText = "select table_type,path,file,table_name from " + this.m_strDataSourceTable + " " + 
					" where scenario_id = '" + 
					this.m_strScenarioId.Trim() +  "';";
			}
			else
			{
				oCommand.CommandText = "select table_type,path,file,table_name from " + this.m_strDataSourceTable + ";";

			}
			try
			{
				System.Data.OleDb.OleDbDataReader oDataReader = oCommand.ExecuteReader();
				int x = 0;
				
				while (oDataReader.Read())
				{
					if (oDataReader["table_type"] != System.DBNull.Value &&
						oDataReader["table_type"].ToString().Trim().Length > 0)
					{
						// Add a ListItem object to the ListView.
						System.Windows.Forms.ListViewItem entryListItem =
							lstRequiredTables.Items.Add(" ");
						this.m_oLvRowColors.AddRow();
						this.m_oLvRowColors.AddColumns(x,lstRequiredTables.Columns.Count);
						entryListItem.UseItemStyleForSubItems=false;
						this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,COLUMN_NULL,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["table_type"].ToString());
						this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,TABLETYPE,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["path"].ToString());
						this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,PATH,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["file"].ToString());
						this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,MDBFILE,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
                        strPathAndFile = oMacroSub.GeneralTranslateVariableSubstitution(oDataReader["path"].ToString().Trim()) 
                            + "\\" + oDataReader["file"].ToString().Trim();
						if (System.IO.File.Exists(strPathAndFile) == true) 
						{
							ListViewItem.ListViewSubItem FileStatusSubItem = 
								entryListItem.SubItems.Add("Found");
							this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,FILESTATUS,FileStatusSubItem,false);
							//FileStatusSubItem.ForeColor = System.Drawing.Color.Black;
							//FileStatusSubItem.BackColor = System.Drawing.Color.White;
							//FileStatusSubItem.Font = new System.Drawing.Font(
							//	"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
							FileStatusSubItem.Font = frmMain.g_oGridViewFont;
							this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["table_name"].ToString());
							this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,TABLE,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);

							//see if the table exists in the mdb database container
							dao_data_access p_dao = new dao_data_access();
							if (p_dao.TableExists(strPathAndFile,oDataReader["table_name"].ToString().Trim()) == true)
							{
								this.lstRequiredTables.Items[x].SubItems.Add("Found");
								this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,TABLESTATUS,entryListItem.SubItems[entryListItem.SubItems.Count-1],lstRequiredTables.Items[x].Selected);
								strConn = p_ado.getMDBConnString(strPathAndFile,"admin","");   //"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strPathAndFile + ";User Id=admin;Password=;";
								strSQL = "select count(*) from " + oDataReader["table_name"].ToString();
								this.lstRequiredTables.Items[x].SubItems.Add(Convert.ToString(p_ado.getRecordCount(strConn,strSQL,oDataReader["table_name"].ToString())));
								this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,RECORDCOUNT,entryListItem.SubItems[entryListItem.SubItems.Count-1],lstRequiredTables.Items[x].Selected);

							}
							else 
							{
								ListViewItem.ListViewSubItem TableStatusSubItem = 
									entryListItem.SubItems.Add("Not Found");
								this.m_oLvRowColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(TABLESTATUS).UpdateColumn=false;
								TableStatusSubItem.ForeColor = System.Drawing.Color.White;
								TableStatusSubItem.BackColor = System.Drawing.Color.Red;
								//TableStatusSubItem.Font = new System.Drawing.Font(
								//	"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
								TableStatusSubItem.Font = frmMain.g_oGridViewFont;
								this.lstRequiredTables.Items[x].SubItems.Add("0");
								this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,RECORDCOUNT,entryListItem.SubItems[entryListItem.SubItems.Count-1],lstRequiredTables.Items[x].Selected);
							}
							p_dao = null;
						}
						else 
						{
							ListViewItem.ListViewSubItem FileStatusSubItem = 
								entryListItem.SubItems.Add("Not Found");
							this.m_oLvRowColors.m_oRowCollection.Item(lstRequiredTables.Items.Count-1).m_oColumnCollection.Item(FILESTATUS).UpdateColumn=false;
							FileStatusSubItem.ForeColor = System.Drawing.Color.White;
							FileStatusSubItem.BackColor = System.Drawing.Color.Red;
							//FileStatusSubItem.Font = new System.Drawing.Font(
							//	"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
							FileStatusSubItem.Font = frmMain.g_oGridViewFont;
							this.lstRequiredTables.Items[x].SubItems.Add(oDataReader["table_name"].ToString());
							ListViewItem.ListViewSubItem TableStatusSubItem = 
								entryListItem.SubItems.Add("Not Found");
							this.m_oLvRowColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(TABLESTATUS).UpdateColumn=false;
							TableStatusSubItem.ForeColor = System.Drawing.Color.White;
							TableStatusSubItem.BackColor = System.Drawing.Color.Red;
							//TableStatusSubItem.Font = new System.Drawing.Font(
							//	"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
							TableStatusSubItem.Font = frmMain.g_oGridViewFont;
							this.lstRequiredTables.Items[x].SubItems.Add("0");
							this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,RECORDCOUNT,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						}
						Datasource.UpdateTableMacroVariable(entryListItem.SubItems[TABLETYPE].Text,entryListItem.SubItems[TABLE].Text);
						this.lstRequiredTables.Items[x].SubItems.Add(Datasource.g_oCurrentSQLMacroSubstitutionVariableItem.VariableName);
						this.m_oLvRowColors.ListViewSubItem(entryListItem.Index,MDBFILE,entryListItem.SubItems[entryListItem.SubItems.Count-1],false);
						x++;
					}
					
				}
				oDataReader.Close();
			}
			catch
			{
				intError = -1;
				strError = "The Query Command " + oCommand.CommandText.ToString() + " Failed";
				MessageBox.Show(strError);
				oConn.Close();
				p_ado= null;
				return;
			}
			
			oConn.Close();
			

			this.lstRequiredTables.Columns[TABLETYPE].Width = -1;
			


			p_ado = null;
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
                    if (!m_bUsingSqlite)
                    {
                        this.populate_listview_grid();
                    }
                    else
                    {
                        this.LoadValuesSqlite();
                    }
					
				    break;
			}
		}

		private void EditDatasource()
		{
			if (this.lstRequiredTables.SelectedItems.Count == 0)
			{
				return;
			}
			string strConn="";
			string strSQL = "";
			int y;
			ado_data_access p_ado = new ado_data_access();

			string strMDBFullPath = this.lstRequiredTables.SelectedItems[0].SubItems[PATH].Text.Trim() + "\\" + 
				this.lstRequiredTables.SelectedItems[0].SubItems[MDBFILE].Text.Trim();
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
                    p_uc = new uc_datasource_edit(this.m_strDataSourceMDBFile, this.m_strScenarioId, false);
                }
                else
                {
                    frmTemp.Text = "Prcoessor: Edit " + this.lstRequiredTables.SelectedItems[0].SubItems[TABLETYPE].Text.Trim() + " Data Source";
                    p_uc = new uc_datasource_edit(this.m_strDataSourceMDBFile, this.m_strScenarioId, this.UsingSqlite);
                }
            }
			else
			{
				frmTemp.Text = "Database: Edit " + this.lstRequiredTables.SelectedItems[0].SubItems[TABLETYPE].Text.Trim() + " Data Source";
				p_uc = new uc_datasource_edit(this.m_strDataSourceMDBFile);
			}
			frmTemp.Controls.Add(p_uc);

			p_uc.strProjectDirectory = this.strProjectDirectory;
			p_uc.strScenarioId = this.strScenarioId;
           
			

			frmTemp.Height=0;
			frmTemp.Width=0;
			if (p_uc.Top + p_uc.Height > frmTemp.ClientSize.Height + 2)
			{
				for (y=1;;y++)
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
				for (y=1;;y++)
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
            
			p_uc.lblMDBFile.Text = strMDBFullPath.Trim();
			p_uc.lblTable.Text = strTable.Trim();
			p_uc.lblTableType.Text = 
				this.lstRequiredTables.SelectedItems[0].SubItems[TABLETYPE].Text.Trim();
			p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
									
			System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
			if (result==System.Windows.Forms.DialogResult.OK)
			{
				utils p_utils = new utils();
				string strDir = p_utils.getDirectory(p_uc.lblNewMDBFile.Text);
				string strFile = p_utils.getFileName(p_uc.lblNewMDBFile.Text);
				this.lstRequiredTables.SelectedItems[0].SubItems[PATH].Text = strDir;
				this.lstRequiredTables.SelectedItems[0].SubItems[MDBFILE].Text = strFile;
				ListViewItem.ListViewSubItem FileSubItem = 
					this.lstRequiredTables.SelectedItems[0].SubItems[FILESTATUS];
                
				
				FileSubItem.Font = frmMain.g_oGridViewFont;
				this.lstRequiredTables.SelectedItems[0].SubItems[FILESTATUS].Text = "Found";
				this.m_oLvRowColors.m_oRowCollection.Item(this.lstRequiredTables.SelectedItems[0].Index).m_oColumnCollection.Item(FILESTATUS).UpdateColumn=true;
				this.m_oLvRowColors.ListViewSubItem(lstRequiredTables.SelectedItems[0].Index,FILESTATUS,FileSubItem,true);
				this.lstRequiredTables.SelectedItems[0].SubItems[TABLE].Text = p_uc.lblNewTable.Text;
				ListViewItem.ListViewSubItem TableSubItem = 
					this.lstRequiredTables.SelectedItems[0].SubItems[TABLESTATUS];

			
				TableSubItem.Font = frmMain.g_oGridViewFont;
				this.lstRequiredTables.SelectedItems[0].SubItems[TABLESTATUS].Text = "Found";
				this.m_oLvRowColors.m_oRowCollection.Item(this.lstRequiredTables.SelectedItems[0].Index).m_oColumnCollection.Item(TABLESTATUS).UpdateColumn=true;
				this.m_oLvRowColors.ListViewSubItem(lstRequiredTables.SelectedItems[0].Index,TABLESTATUS,TableSubItem,true);
				p_utils = null;
				strConn=p_ado.getMDBConnString(p_uc.lblNewMDBFile.Text.Trim(),"","");
				strSQL = "select count(*) from " + p_uc.lblNewTable.Text.Trim();
				this.lstRequiredTables.SelectedItems[0].SubItems[RECORDCOUNT].Text =
					Convert.ToString(p_ado.getRecordCount(strConn,strSQL,p_uc.lblNewTable.Text.Trim()));

			}
			frmTemp.Close();
			frmTemp.Dispose();
			frmTemp = null;
			p_ado = null;
		}
		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			if (this.lstRequiredTables.SelectedItems.Count == 0)
			{
				return;
			}
            string strConn="";
			string strSQL = "";
			int y;
			ado_data_access p_ado = new ado_data_access();

            string strMDBFullPath = this.lstRequiredTables.SelectedItems[0].SubItems[PATH].Text.Trim() + "\\" + 
				             this.lstRequiredTables.SelectedItems[0].SubItems[MDBFILE].Text.Trim();
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
                    p_uc = new uc_datasource_edit(this.m_strDataSourceMDBFile, this.m_strScenarioId, false);
                }
                else
                {
                    frmTemp.Text = "Prcoessor: Edit " + this.lstRequiredTables.SelectedItems[0].SubItems[TABLETYPE].Text.Trim() + " Data Source";
                    p_uc = new uc_datasource_edit(this.m_strDataSourceMDBFile, this.m_strScenarioId, this.UsingSqlite);
                }
			}
			else
			{
				frmTemp.Text = "Database: Edit " + this.lstRequiredTables.SelectedItems[0].SubItems[TABLETYPE].Text.Trim() + " Data Source";
			   	p_uc = new uc_datasource_edit(this.m_strDataSourceMDBFile);
			}
			frmTemp.Controls.Add(p_uc);

			p_uc.strProjectDirectory = this.strProjectDirectory;
			p_uc.strScenarioId = this.strScenarioId;
           
			

			frmTemp.Height=0;
			frmTemp.Width=0;
			if (p_uc.Top + p_uc.Height > frmTemp.ClientSize.Height + 2)
			{
				for (y=1;;y++)
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
				for (y=1;;y++)
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
            
			p_uc.lblMDBFile.Text = strMDBFullPath.Trim();
			p_uc.lblTable.Text = strTable.Trim();
			p_uc.lblTableType.Text = 
				this.lstRequiredTables.SelectedItems[0].SubItems[TABLETYPE].Text.Trim();
			p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
									
			System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
			if (result==System.Windows.Forms.DialogResult.OK)
			{
				utils p_utils = new utils();
				string strDir = p_utils.getDirectory(p_uc.lblNewMDBFile.Text);
				string strFile = p_utils.getFileName(p_uc.lblNewMDBFile.Text);
				this.lstRequiredTables.SelectedItems[0].SubItems[PATH].Text = strDir;
				this.lstRequiredTables.SelectedItems[0].SubItems[MDBFILE].Text = strFile;
				ListViewItem.ListViewSubItem FileSubItem = 
					this.lstRequiredTables.SelectedItems[0].SubItems[FILESTATUS];
                
				// Change the expenseItem object's color and font.
				//FileSubItem.ForeColor = System.Drawing.Color.Black;
				//FileSubItem.BackColor = System.Drawing.Color.White;
				//FileSubItem.Font = new System.Drawing.Font(
				//	"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
				FileSubItem.Font = frmMain.g_oGridViewFont;
				this.lstRequiredTables.SelectedItems[0].SubItems[FILESTATUS].Text = "Found";
				this.m_oLvRowColors.m_oRowCollection.Item(this.lstRequiredTables.SelectedItems[0].Index).m_oColumnCollection.Item(FILESTATUS).UpdateColumn=true;
				this.m_oLvRowColors.ListViewSubItem(lstRequiredTables.SelectedItems[0].Index,FILESTATUS,FileSubItem,true);
				this.lstRequiredTables.SelectedItems[0].SubItems[TABLE].Text = p_uc.lblNewTable.Text;
				ListViewItem.ListViewSubItem TableSubItem = 
					this.lstRequiredTables.SelectedItems[0].SubItems[TABLESTATUS];

				// Change the expenseItem object's color and font.
				//TableSubItem.ForeColor = System.Drawing.Color.Black;
				//TableSubItem.BackColor = System.Drawing.Color.White;
				//FileSubItem.Font = new System.Drawing.Font(
				//	"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
				TableSubItem.Font = frmMain.g_oGridViewFont;
				this.lstRequiredTables.SelectedItems[0].SubItems[TABLESTATUS].Text = "Found";
				this.m_oLvRowColors.m_oRowCollection.Item(this.lstRequiredTables.SelectedItems[0].Index).m_oColumnCollection.Item(TABLESTATUS).UpdateColumn=true;
				this.m_oLvRowColors.ListViewSubItem(lstRequiredTables.SelectedItems[0].Index,TABLESTATUS,TableSubItem,true);
				p_utils = null;
				strConn=p_ado.getMDBConnString(p_uc.lblNewMDBFile.Text.Trim(),"","");
				strSQL = "select count(*) from " + p_uc.lblNewTable.Text.Trim();
				this.lstRequiredTables.SelectedItems[0].SubItems[RECORDCOUNT].Text =
					Convert.ToString(p_ado.getRecordCount(strConn,strSQL,p_uc.lblNewTable.Text.Trim()));

			}
			frmTemp.Close();
			frmTemp.Dispose();
			frmTemp = null;
			p_ado = null;
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

		private void btnCopy_Click(object sender, System.EventArgs e)
		{
		   /***********************************************************
		    **since we are copying all the mdb files to the same
			**scenario directory, ensure that there are not 2 of the same
			**mdb file names that reside in 2 different directories
			************************************************************/
            int x;
			int y;
			
			string strMDBFile;
			string strDir;
			string strCopyFrom;
			string strCopyTo;
			string strCopyToDir;
			string strConn;
			string strTable="";
			System.Data.DataTable p_dt;

			if (this.lstRequiredTables.CheckedItems.Count==0) return;

			
			this.progressBar1.Left = (int)(this.groupBox1.Width * .50) - (int)(this.progressBar1.Width * .50);

			this.lblProgress.Top = this.progressBar1.Top + this.progressBar1.Height + 3;
			this.lblProgress.Left = this.progressBar1.Left;


			Datasource p_datasource = new Datasource(this.m_strProjectDirectory,this.m_strScenarioId);
            string strTempMDBFile = p_datasource.CreateMDBAndTableDataSourceLinks();
			try
			{
				
				string strScenarioConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + 
					((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + 
					"\\" + ScenarioType + "\\db\\scenario_" + ScenarioType + "_rule_definitions.mdb;User Id=admin;Password=;";
				
			   /*******************************
			    ** copy the files
			    *******************************/

				this.progressBar1.Minimum=0;
				this.progressBar1.Maximum=this.lstRequiredTables.CheckedItems.Count;
				this.progressBar1.Value=0;

				this.lblProgress.Text="";
				this.progressBar1.Visible=true;
				this.lblProgress.Visible=true;

                if (this.ScenarioType.Trim().ToUpper() == "OPTIMIZER")
					strCopyToDir = ((frmOptimizerScenario)this.ParentForm).uc_scenario1.txtScenarioPath.Text.Trim() + "\\db";
				else 
					strCopyToDir = this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db";
			    for (x=0; x <= this.lstRequiredTables.Items.Count-1; x++)
			    {
					if (this.lstRequiredTables.Items[x].Checked==true)
					{
						this.progressBar1.Value=this.progressBar1.Value + 1;
						this.progressBar1.Refresh();
						this.lblProgress.Text = "Processing " + this.progressBar1.Value.ToString().Trim() + " Of " + this.progressBar1.Maximum.ToString().Trim();
						this.lblProgress.Refresh();


						
						//get structure of the table
						ado_data_access p_ado = new ado_data_access();
						strConn = p_ado.getMDBConnString(strTempMDBFile,"","");
						p_dt = p_ado.getTableSchema(strConn,"select * from " + this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim());
						

						//get the table structure

						strMDBFile = this.lstRequiredTables.Items[x].SubItems[MDBFILE].Text.Trim();
						strDir = this.lstRequiredTables.Items[x].SubItems[PATH].Text.Trim();
						strCopyFrom = strDir + "\\" + strMDBFile;
						strCopyTo = strCopyToDir + "\\" + strMDBFile;
						
						dao_data_access p_dao = new dao_data_access();
						strTable=this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim();
						//check to see if the file is already in the scenario directory
						if (strCopyFrom.ToUpper().Trim() != strCopyTo.ToUpper().Trim())
						{
							
							//create the MDB  File
							if (System.IO.File.Exists(strCopyTo) == false)
							{
								p_dao.CreateMDB(strCopyTo);
								p_dao.CreateMDBTableFromDataSetTable(strCopyTo,strTable,p_dt,true);
							}
							else
							{
								//check to see if the table is in the mdb file
								if (p_dao.TableExists(strCopyTo,strTable) == true)
								{
									FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
									frmTemp.MaximizeBox = false;
									frmTemp.BackColor = System.Drawing.SystemColors.Control;
									frmTemp.Text = "Copy Table:" + strTable + " Exists";

									FIA_Biosum_Manager.uc_table_exists_dialog p_uc = new uc_table_exists_dialog();
									frmTemp.Controls.Add(p_uc);
									
           
			

									frmTemp.Height=0;
									frmTemp.Width=0;
									if (p_uc.Top + p_uc.Height > frmTemp.ClientSize.Height + 2)
									{
										for (y=1;;y++)
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
										for (y=1;;y++)
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
									p_uc.strMDBFileLabel = "Table Contents Of " + strCopyTo.Trim();
									p_uc.listBox1.Items.Clear();
									p_dao.LoadTablesIntoListBox(strCopyTo,p_uc.listBox1);
									p_uc.strTable=strTable;
									frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
									frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
									for (;;)
									{
									   
										System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
										if (result==System.Windows.Forms.DialogResult.OK)
										{
											if (p_uc.rdoOverwrite.Checked==true)
											{
												p_dao.DeleteTableFromMDB(strCopyTo,this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim());
												p_dao.CreateMDBTableFromDataSetTable(strCopyTo,strTable,p_dt,true);
												break;
											}
											else if (p_dao.TableExists(strCopyTo,p_uc.strTable) == false)
											{
												strTable = p_uc.strTable;
												p_dao.CreateMDBTableFromDataSetTable(strCopyTo,strTable,p_dt,true);
												break;
											}
											else
											{
												frmTemp.Text = "Copy Table:" + p_uc.strTable + " Exists";
											}
										}
										else
										{
											strTable="";
											break;
										}
									}
									frmTemp.Dispose();
								}
								else
								{
									p_dao.CreateMDBTableFromDataSetTable(strCopyTo,strTable,p_dt,true);
								}
							}
							//create primary indexes and autonumbers
							p_datasource.SetPrimaryIndexesAndAutoNumbers(strCopyTo,this.lstRequiredTables.Items[x].Text.Trim(),strTable,p_dao);
							if (strTable.Trim().Length > 0)
							{

								//append the contents of the source table to the target table
								p_dao.CreateTableLink(strTempMDBFile,"temp_link",strCopyTo,strTable,true);
								p_dao=null;
								p_ado.m_strSQL = "INSERT INTO temp_link SELECT * FROM " + this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim() + ";";
								p_ado.SqlNonQuery(strConn,p_ado.m_strSQL);
								if (p_ado.m_intError==0)
								{
									p_ado.m_strSQL = "UPDATE scenario_datasource SET path = '" + strCopyToDir + "',table_name = '" + strTable.Trim() + "' " + 
										" WHERE trim(scenario_id) = '" + this.m_strScenarioId.Trim() + "' AND " +
										"trim(table_type) = '" + this.lstRequiredTables.Items[x].SubItems[TABLETYPE].Text.Trim() + "';";
									p_ado.SqlNonQuery(strScenarioConn, p_ado.m_strSQL);
									if (p_ado.m_intError == 0)
									{
										this.lstRequiredTables.Items[x].SubItems[PATH].Text = strCopyToDir;
										this.lstRequiredTables.Items[x].SubItems[TABLE].Text = strTable.Trim();
									}
								}
							}
							
						}

						p_dao=null;
						p_ado = null;
					}
			    }
				this.lstRequiredTables.Columns[PATH].Width = -1;
				this.lblProgress.Visible=false;
				this.progressBar1.Visible=false;

			}
			catch (Exception e2)
			{
				MessageBox.Show(e2.Message);

			}
			p_datasource = null;

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
		/*****************************************************
		 ** create a mdb table in the users temporary dir
		 ** and create a link to each of the scenario 
		 ** data source tables.  Return the name of the 
		 ** temporary mdb file to the calling function
		 *****************************************************/
		public string CreateMDBAndScenarioTableDataSourceLinks(string strDestinationLinkDir)
		{
			
			string strTempMDB="";
			int x;
            this.m_intNumberOfValidTables=0;
			//used to get the temporary random file name
			utils p_utils = new utils();
			//used to create a link to the table
			dao_data_access p_dao = new dao_data_access();
			for (x=0; x <= this.lstRequiredTables.Items.Count - 1; x++)
			{
					if (this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper()=="FOUND" &&
						this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper()=="FOUND")
					{
						if (strTempMDB.Trim().Length == 0)
						{
							//get temporary mdb file
							strTempMDB = 
								p_utils.getRandomFile(strDestinationLinkDir,"accdb");

							//create a temporary mdb that will contain all 
							//the links to the scenario datasource tables
							p_dao.CreateMDB(strTempMDB);

						}
						p_dao.CreateTableLink(strTempMDB,
							this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim(),
							this.lstRequiredTables.Items[x].SubItems[PATH].Text.Trim() + "\\" +
							     this.lstRequiredTables.Items[x].SubItems[MDBFILE].Text.Trim(),
							this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim());
						this.m_intNumberOfValidTables++;
						

					}
			}
			p_utils = null;
			p_dao = null;
            if (strTempMDB.Trim().Length == 0)
				MessageBox.Show("!!None of the scenario data source tables are found!!");
			return strTempMDB;
		}


		/*****************************************************
		 ** create a mdb table in the users temporary dir
		 ** and create a link to each of the scenario 
		 ** data source tables.  Return the name of the 
		 ** temporary mdb file to the calling function
		 *****************************************************/
		public void CreateScenarioTableDataSourceLinks(string strDestinationDbFile)
		{
			
			int x;
			this.m_intNumberOfValidTables=0;
			//used to get the temporary random file name
			
			//used to create a link to the table
			dao_data_access p_dao = new dao_data_access();
			for (x=0; x <= this.lstRequiredTables.Items.Count - 1; x++)
			{
				if (this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper()=="FOUND" &&
					this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper()=="FOUND")
				{
					p_dao.CreateTableLink(strDestinationDbFile,
						this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim(),
						this.lstRequiredTables.Items[x].SubItems[PATH].Text.Trim() + "\\" +
						this.lstRequiredTables.Items[x].SubItems[MDBFILE].Text.Trim(),
						this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim());
					this.m_intNumberOfValidTables++;
						

				}
			}
			p_dao = null;
			if (m_intNumberOfValidTables == 0)
				MessageBox.Show("!!None of the scenario data source tables are found!!");
			
		}

		public void getScenarioDataSourceMDBAndTableName(string strTableID,ref string strMDBPathAndFile,ref string strTable)
		{
			int x;
			for (x=0; x<= this.lstRequiredTables.Items.Count-1;x++)
			{
				if (strTableID.Trim().ToUpper() == 
					this.lstRequiredTables.Items[x].SubItems[TABLETYPE].Text.Trim().ToUpper()
					&&
					this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper() =="FOUND" 
					&&
					this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper() == "FOUND")
				{
					  strMDBPathAndFile = this.lstRequiredTables.Items[x].SubItems[PATH].Text.Trim() + "\\" + this.lstRequiredTables.Items[x].SubItems[MDBFILE].Text.Trim();
					  strTable = this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim();
					  break;
				}
			}
		}

		public void getNumberOfValidTables()
		{
			int x;
			this.m_intNumberOfValidTables=0;
			for (x=0; x <= this.lstRequiredTables.Items.Count - 1; x++)
			{
				if (this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper()=="FOUND" &&
					this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper()=="FOUND")
				{
					this.m_intNumberOfValidTables++;
				}
			}

		}

		public void LoadScenarioDataSourceTablesIntoListBox(System.Windows.Forms.ListBox listbox1)
		{
			int x;
			for (x=0; x <= this.lstRequiredTables.Items.Count - 1; x++)
			{
				if (this.lstRequiredTables.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper()=="FOUND" &&
					this.lstRequiredTables.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper()=="FOUND")
				{
					listbox1.Items.Add(this.lstRequiredTables.Items[x].SubItems[TABLE].Text.Trim());
				}
			}
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
                        this.lstRequiredTables.Items[x].SubItems[MDBFILE].Text.Trim();
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
						            this.lstRequiredTables.Items[x].SubItems[MDBFILE].Text.Trim() + " is not found");
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

		private void btnRefresh_Click(object sender, System.EventArgs e)
		{
            if (!m_bUsingSqlite)
            {
                this.populate_listview_grid();
            }
            else
            {
                this.LoadValuesSqlite();
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
                    if (!m_bUsingSqlite)
                    {
                        this.populate_listview_grid();
                    }
                    else
                    {
                        this.LoadValuesSqlite();
                    }
					
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
				this.m_strDataSourceMDBFile=value;
			}
			get
			{
				return this.m_strDataSourceMDBFile;
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
        public bool UsingSqlite
        {
            set
            {
                this.m_bUsingSqlite = value;
            }
            get
            {
                return this.m_bUsingSqlite;
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

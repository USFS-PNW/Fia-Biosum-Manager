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
	/// Summary description for uc_scenario_notes.
	/// </summary>
	public class uc_optimizer_scenario_run : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario=null;
		public System.Windows.Forms.Button btnViewLog;
		public System.Windows.Forms.Button btnSqlite;
		public System.Windows.Forms.Button btnViewResultsTables;
        public System.Windows.Forms.Button btnViewAuditTables;
		public System.Windows.Forms.Label lblMsg;
		public System.Windows.Forms.Button btnCancel;

		private int m_intError=0;
		public System.Data.DataSet m_ds;
		public System.Data.OleDb.OleDbConnection m_conn;
		public System.Data.OleDb.OleDbDataAdapter m_da;
        public RunOptimizer m_oRunOptimizer;
		public string m_strCustomPlotSQL="";
		private FIA_Biosum_Manager.frmGridView m_frmGridView;

		public bool m_bUserCancel=false;
		private bool m_bAbortThread=false;
		private System.Threading.Thread m_thread=null;
		public System.Windows.Forms.Label m_lblCurrentProcessStatus;
        private System.Windows.Forms.Panel panel1;

        private ListViewAlternateBackgroundColors m_oLvAlternateColors = new ListViewAlternateBackgroundColors();
        public ListViewEmbeddedControls.ListViewEx listViewEx1;
        private ColumnHeader colNull;
        private ColumnHeader colDesc;
        private ColumnHeader colStatus;
        private Panel pnlFileSizeMonitor;
        public uc_filesize_monitor uc_filesize_monitor3;
        public uc_filesize_monitor uc_filesize_monitor2;
        public uc_filesize_monitor uc_filesize_monitor1;
        public uc_filesize_monitor uc_filesize_monitor4;


		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_optimizer_scenario_run()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call
            LoadListView();
           


		}
        private void LoadListView()
        {
            //
            //INITIALIZE LISTVIEW ALTERNATE ROW COLORS
            //
           
            this.m_oLvAlternateColors.InitializeRowCollection();
            this.m_oLvAlternateColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceListView = this.listViewEx1;
            this.m_oLvAlternateColors.CustomFullRowSelect = true;
            m_oLvAlternateColors.ColumnsToNotUpdate("1");
            if (frmMain.g_oGridViewFont != null) this.listViewEx1.Font = frmMain.g_oGridViewFont;

          
            //
            //Validate Rule Definitions
            //
            this.AddListViewRowItem("Validate Rule Definitions",false, false);
            //
            //Save Rule Definitions
            //
            this.AddListViewRowItem("Save Rule Definitions",false, false);
            //
            //Initialize and Load Variables
            //
            this.AddListViewRowItem("Initialize and Load Variables", false, false);
            //
            //Accessibility
            //
            this.AddListViewRowItem("Determine If Stand And Conditions Are Accessible For Treatment And Harvest",false, false);
            //
            //Least Expensive Routes
            //
            this.AddListViewRowItem("Get Least Expensive Route From Stand To Wood Processing Facility",false, false);
            //
            //
            //
            this.AddListViewRowItem("Sum Tree Yields, Volume, And Value For A Stand And Treatment",false, false);
            //
            //
            //
            this.AddListViewRowItem("Apply User Defined Filters And Get Valid Stand Combinations", false, false);
            //
            //
            //
            this.AddListViewRowItem("Populate Valid Combination Audit Data", true, true);
            //
            //
            //
            this.AddListViewRowItem("Create Condition - Processing Site Table", false, false);
            //
            //
            //
            this.AddListViewRowItem("Populate Context Database", true, false);
            //
            //
            //
            //this.AddListViewRowItem("Populate FVS PRE-POST Context Database", true, false);
            //
            //
            //
            this.AddListViewRowItem("Summarize Wood Product Volume Yields, Costs, And Net Revenue For A Stand And Treatment", false, false);
            //
            //
            //
            this.AddListViewRowItem("Summarize Wood Product Volume Yields, Costs, And Net Revenue For A Stand And Treatment Package", false, false);
            //
            //
            //
            this.AddListViewRowItem("Calculate Weighted Economic Variables For Each Stand And Treatment Package", false, false);
            //
            //
            //
            this.AddListViewRowItem("Identify Effective Treatments For Each Stand", false, false);
            //
            //
            //
            this.AddListViewRowItem("Optimize the Effective Treatments For Each Stand", false, false);
            //
            //
            //
            this.AddListViewRowItem("Load Tie Breaker Tables", false, false);
            //
            //
            //
            this.AddListViewRowItem("Identify The Best Effective Treatment For Each Stand", false, false);
            this.listViewEx1.Columns[2].Width = -1;



        }

        private void AddListViewRowItem(string p_strDescription,bool p_bCheckBox, bool p_bCheckBoxChecked)
        {
            System.Windows.Forms.ListViewItem entryListItem = null;
          
            entryListItem = listViewEx1.Items.Add(" ");


            entryListItem.UseItemStyleForSubItems = false;

            this.m_oLvAlternateColors.AddRow();
            this.m_oLvAlternateColors.AddColumns(listViewEx1.Items.Count - 1, listViewEx1.Columns.Count);

            if (p_bCheckBox)
            {
                System.Windows.Forms.CheckBox oCheckBox = new CheckBox();

                listViewEx1.AddEmbeddedControl(oCheckBox, 0, listViewEx1.Items.Count - 1, System.Windows.Forms.DockStyle.Fill);
               
                oCheckBox.Text = "";
                oCheckBox = (CheckBox)listViewEx1.GetEmbeddedControl(0, listViewEx1.Items.Count - 1);
                oCheckBox.Visible = p_bCheckBox;
                oCheckBox.Show();
                oCheckBox.Enabled = p_bCheckBox;
                oCheckBox.Checked = p_bCheckBoxChecked;
            }
           
            //ProgressBar
            entryListItem.SubItems.Add(" ");
            ProgressBarBasic.ProgressBarBasic pb = new ProgressBarBasic.ProgressBarBasic();
            pb.Value = 0;
            pb.Name = "ProgressBar" + listViewEx1.Items.Count.ToString().Trim();
            pb.Orientation = System.Windows.Forms.Orientation.Horizontal;
            pb.TextStyle = ProgressBarBasic.ProgressBarBasic.TextStyleType.Percentage;
            pb.BackColor = Color.LightGray;
            pb.ForeColor = Color.Gold;
            pb.Visible = true;
            // Embed the ProgressBar in Column 1
            listViewEx1.AddEmbeddedControl(pb, 1, listViewEx1.Items.Count - 1);

            pb = (ProgressBarBasic.ProgressBarBasic)listViewEx1.GetEmbeddedControl(1, listViewEx1.Items.Count - 1);

            frmMain.g_oDelegate.SetControlPropertyValue(pb, "Maximum", 100);

            frmMain.g_oDelegate.SetControlPropertyValue(pb, "Minimum", 0);
            frmMain.g_oDelegate.SetControlPropertyValue(pb, "Value", 0);
            frmMain.g_oDelegate.SetControlPropertyValue(pb, "Text", "0%");

            //Description
            entryListItem.SubItems.Add(p_strDescription);
            this.m_oLvAlternateColors.ListViewItem(listViewEx1.Items[listViewEx1.Items.Count - 1]);
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pnlFileSizeMonitor = new System.Windows.Forms.Panel();
            this.uc_filesize_monitor4 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.uc_filesize_monitor3 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.uc_filesize_monitor2 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.uc_filesize_monitor1 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.listViewEx1 = new ListViewEmbeddedControls.ListViewEx();
            this.colNull = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colStatus = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colDesc = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lblMsg = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnSqlite = new System.Windows.Forms.Button();
            this.btnViewResultsTables = new System.Windows.Forms.Button();
            this.btnViewAuditTables = new System.Windows.Forms.Button();
            this.btnViewLog = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.pnlFileSizeMonitor.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(926, 508);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.pnlFileSizeMonitor);
            this.panel1.Controls.Add(this.listViewEx1);
            this.panel1.Controls.Add(this.lblMsg);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.lblTitle);
            this.panel1.Controls.Add(this.btnSqlite);
            this.panel1.Controls.Add(this.btnViewResultsTables);
            this.panel1.Controls.Add(this.btnViewAuditTables);
            this.panel1.Controls.Add(this.btnViewLog);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 18);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(920, 487);
            this.panel1.TabIndex = 40;
            // 
            // pnlFileSizeMonitor
            // 
            this.pnlFileSizeMonitor.AutoScroll = true;
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor4);
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor3);
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor2);
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor1);
            this.pnlFileSizeMonitor.Location = new System.Drawing.Point(12, 380);
            this.pnlFileSizeMonitor.Name = "pnlFileSizeMonitor";
            this.pnlFileSizeMonitor.Size = new System.Drawing.Size(829, 99);
            this.pnlFileSizeMonitor.TabIndex = 69;
            // 
            // uc_filesize_monitor4
            // 
            this.uc_filesize_monitor4.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor4.Information = "";
            this.uc_filesize_monitor4.Location = new System.Drawing.Point(616, 10);
            this.uc_filesize_monitor4.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.uc_filesize_monitor4.Name = "uc_filesize_monitor4";
            this.uc_filesize_monitor4.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor4.TabIndex = 3;
            this.uc_filesize_monitor4.Visible = false;
            // 
            // uc_filesize_monitor3
            // 
            this.uc_filesize_monitor3.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor3.Information = "";
            this.uc_filesize_monitor3.Location = new System.Drawing.Point(429, 10);
            this.uc_filesize_monitor3.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.uc_filesize_monitor3.Name = "uc_filesize_monitor3";
            this.uc_filesize_monitor3.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor3.TabIndex = 2;
            this.uc_filesize_monitor3.Visible = false;
            // 
            // uc_filesize_monitor2
            // 
            this.uc_filesize_monitor2.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor2.Information = "";
            this.uc_filesize_monitor2.Location = new System.Drawing.Point(206, 10);
            this.uc_filesize_monitor2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.uc_filesize_monitor2.Name = "uc_filesize_monitor2";
            this.uc_filesize_monitor2.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor2.TabIndex = 1;
            this.uc_filesize_monitor2.Visible = false;
            // 
            // uc_filesize_monitor1
            // 
            this.uc_filesize_monitor1.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor1.Information = "";
            this.uc_filesize_monitor1.Location = new System.Drawing.Point(5, 10);
            this.uc_filesize_monitor1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.uc_filesize_monitor1.Name = "uc_filesize_monitor1";
            this.uc_filesize_monitor1.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor1.TabIndex = 0;
            this.uc_filesize_monitor1.Visible = false;
            // 
            // listViewEx1
            // 
            this.listViewEx1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colNull,
            this.colStatus,
            this.colDesc});
            this.listViewEx1.FullRowSelect = true;
            this.listViewEx1.GridLines = true;
            this.listViewEx1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listViewEx1.HideSelection = false;
            this.listViewEx1.Location = new System.Drawing.Point(13, 50);
            this.listViewEx1.MultiSelect = false;
            this.listViewEx1.Name = "listViewEx1";
            this.listViewEx1.Size = new System.Drawing.Size(896, 275);
            this.listViewEx1.TabIndex = 42;
            this.listViewEx1.UseCompatibleStateImageBehavior = false;
            this.listViewEx1.View = System.Windows.Forms.View.Details;
            // 
            // colNull
            // 
            this.colNull.Text = "Optional";
            this.colNull.Width = 55;
            // 
            // colStatus
            // 
            this.colStatus.Text = "Run Status";
            this.colStatus.Width = 250;
            // 
            // colDesc
            // 
            this.colDesc.Text = "Description";
            this.colDesc.Width = 502;
            // 
            // lblMsg
            // 
            this.lblMsg.Enabled = false;
            this.lblMsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMsg.ForeColor = System.Drawing.Color.Black;
            this.lblMsg.Location = new System.Drawing.Point(9, 360);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(660, 16);
            this.lblMsg.TabIndex = 38;
            this.lblMsg.Text = "lblMsg";
            this.lblMsg.Visible = false;
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.ForeColor = System.Drawing.Color.Black;
            this.btnCancel.Location = new System.Drawing.Point(83, 8);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(72, 30);
            this.btnCancel.TabIndex = 39;
            this.btnCancel.Text = "Start";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(8, 5);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(69, 23);
            this.lblTitle.TabIndex = 26;
            this.lblTitle.Text = "Run";
            // 
            // btnSQqlite
            // 
            this.btnSqlite.Enabled = false;
            this.btnSqlite.ForeColor = System.Drawing.Color.Black;
            this.btnSqlite.Location = new System.Drawing.Point(349, 8);
            this.btnSqlite.Name = "btnSqlite";
            this.btnSqlite.Size = new System.Drawing.Size(120, 30);
            this.btnSqlite.TabIndex = 33;
            this.btnSqlite.Text = "SQLite";
            this.btnSqlite.Click += new System.EventHandler(this.btnSqlite_Click);
            // 
            // btnViewResultsTables
            // 
            this.btnViewResultsTables.ForeColor = System.Drawing.Color.Black;
            this.btnViewResultsTables.Location = new System.Drawing.Point(188, 8);
            this.btnViewResultsTables.Name = "btnViewResultsTables";
            this.btnViewResultsTables.Size = new System.Drawing.Size(150, 30);
            this.btnViewResultsTables.TabIndex = 32;
            this.btnViewResultsTables.Text = "View Results Tables";
            this.btnViewResultsTables.Click += new System.EventHandler(this.btnViewScenarioTables_Click);
            // 
            // btnViewAuditTables
            // 
            this.btnViewAuditTables.ForeColor = System.Drawing.Color.Black;
            this.btnViewAuditTables.Location = new System.Drawing.Point(480, 8);
            this.btnViewAuditTables.Name = "btnViewAuditTables";
            this.btnViewAuditTables.Size = new System.Drawing.Size(120, 30);
            this.btnViewAuditTables.TabIndex = 31;
            this.btnViewAuditTables.Text = "View Audit Data";
            this.btnViewAuditTables.Click += new System.EventHandler(this.btnViewAuditTables_Click);
            // 
            // btnViewLog
            // 
            this.btnViewLog.ForeColor = System.Drawing.Color.Black;
            this.btnViewLog.Location = new System.Drawing.Point(610, 8);
            this.btnViewLog.Name = "btnViewLog";
            this.btnViewLog.Size = new System.Drawing.Size(100, 30);
            this.btnViewLog.TabIndex = 34;
            this.btnViewLog.Text = "View Log File";
            this.btnViewLog.Click += new System.EventHandler(this.btnViewLog_Click);
            // 
            // uc_optimizer_scenario_run
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_optimizer_scenario_run";
            this.Size = new System.Drawing.Size(926, 508);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.pnlFileSizeMonitor.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		
		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.Visible=false;
			
			((frmOptimizerScenario)this.ParentForm).Height = 0 ; 
		}


        private void btnSqlite_Click(object sender, System.EventArgs e)
        {
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.UseShellExecute = true;
            try
            {
                proc.StartInfo.FileName = ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsSqliteDbFile;
            }
            catch
            {
            }
            try
            {
                proc.Start();
            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - uc_db:btnOpen_Click() \n" +
                    "Err Msg - " + err.Message,
                    "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);

                this.m_intError = -1;
            }
        }

		private void btnViewAuditTables_Click(object sender, System.EventArgs e)
		{
            viewAuditTables();	
           
		}
        
        private void viewAuditTables()
        {
            string strConn = "";
            int x;
            DataMgr oDataMgr = new DataMgr();
            this.m_frmGridView = new frmGridView();
            this.m_frmGridView.Text = "Treatment Optimizer: Audit";
            lblMsg.Text = "";
            lblMsg.Show();
            string strDbPathAndFile = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\" + Tables.Audit.DefaultCondAuditTableSqliteDbFile;

            if (System.IO.File.Exists(strDbPathAndFile) == true)
            {
                strConn = oDataMgr.GetConnectionString(strDbPathAndFile);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    if (oDataMgr.TableExist(conn, Tables.Audit.DefaultCondAuditTableName) == true)
                    {
                        this.lblMsg.Text = Tables.Audit.DefaultCondAuditTableName;
                        this.lblMsg.Refresh();
                        this.m_frmGridView.UsingSQLite = true;
                        this.m_frmGridView.LoadDataSet(strConn, "select * from " + Tables.Audit.DefaultCondAuditTableName, Tables.Audit.DefaultCondAuditTableName);
                    }
                    if (oDataMgr.TableExist(conn, Tables.Audit.DefaultCondRxAuditTableName) == true)
                    {
                        this.lblMsg.Text = Tables.Audit.DefaultCondRxAuditTableName;
                        this.lblMsg.Refresh();
                        this.m_frmGridView.UsingSQLite = true;
                        this.m_frmGridView.LoadDataSet(strConn, "select * from " + Tables.Audit.DefaultCondRxAuditTableName, Tables.Audit.DefaultCondRxAuditTableName);
                    }
                }
            }
            this.m_frmGridView.TileGridViews();
            this.m_frmGridView.Show();
            this.m_frmGridView.Focus();
            lblMsg.Text = "";
            lblMsg.Refresh();
            lblMsg.Hide();
            oDataMgr = null;
        }
        
        private void btnViewScenarioTables_Click(object sender, System.EventArgs e)
		{
			viewResultsTables();
		}

		/// <summary>
		/// every optimizer_results.accdb table is viewed in a uc_gridview control
		/// </summary>
        private void viewResultsTables()
        {
            string strDBPathAndFile = "";
            string strConn = "";
            string strSQL = "";
            string[] strTableNames;
            strTableNames = new string[1];
            DataMgr p_dataMgr = new DataMgr();

            strDBPathAndFile = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsSqliteDbFile;
            strConn = p_dataMgr.GetConnectionString(strDBPathAndFile);

            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                conn.Open();

                strTableNames = p_dataMgr.getTableNames(conn);

                if (p_dataMgr.m_intError == 0)
                {
                    if (strTableNames.Length > 0)
                    {
                        this.lblMsg.Text = "";
                        this.lblMsg.Visible = true;

                        this.m_frmGridView = new frmGridView();
                        this.m_frmGridView.Text = "Treatment Optimizer: Run Scenario Results (" + this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim() + ")";

                        foreach (string strTable in strTableNames)
                        {
                            this.lblMsg.Text = strTable;
                            this.lblMsg.Refresh();
                            strSQL = "select * from " + strTable.Trim();
                            this.m_frmGridView.UsingSQLite = true;
                            this.m_frmGridView.LoadDataSet(strConn, strSQL, strTable.Trim());
                        }

                        this.lblMsg.Text = "";
                        this.lblMsg.Visible = false;
                        if (strTableNames.Length > 1) this.m_frmGridView.TileGridViews();
                        this.m_frmGridView.Show();
                        this.m_frmGridView.Focus();
                    }

                    else
                    {
                        MessageBox.Show("No Tables Found In " + strDBPathAndFile);
                    }
                }
            }
        }

        private void btnViewLog_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process proc = new System.Diagnostics.Process();
			proc.StartInfo.UseShellExecute = true;
			try
			{
				proc.StartInfo.FileName = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\runlog.txt";
			}
			catch
			{
			}
			try
			{
				proc.Start();
			}
			catch (Exception caught)
			{
				MessageBox.Show(caught.Message);
			}
		}


		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			DialogResult result;

			if (this.btnCancel.Text.Trim().ToUpper() == "CANCEL")
			{

				if (this.m_bAbortThread==true) return;

				if (frmMain.g_oDelegate.m_oThread == null)
				{
					result =  MessageBox.Show("Cancel Running The Scenario (Y/N)?","Cancel Process", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
					switch (result) 
					{
						case DialogResult.Yes:
							this.btnCancel.Text = "Start";
							this.m_bUserCancel=true;
                            uc_filesize_monitor1.EndMonitoringFile();
                            uc_filesize_monitor2.EndMonitoringFile();
							return;
						case DialogResult.No:
							return;
					}
				}
				else
				{
					if (frmMain.g_oDelegate.m_oThread.IsAlive)
					{  
						frmMain.g_oDelegate.AbortProcessing("Cancel Processing","Cancel Running The Scenario (Y/N)?");
						if (frmMain.g_oDelegate.CurrentThreadProcessAborted)
						{
							frmMain.g_oDelegate.StopThread();
                            if (FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic != null)
                            {
                                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                                frmMain.g_oDelegate.SetControlPropertyValue((Control)FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "TextStyle", ProgressBarBasic.ProgressBarBasic.TextStyleType.Text);
                                frmMain.g_oDelegate.SetControlPropertyValue((Control)FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Text", "!!Cancelled!!");
                            }
                            if (m_lblCurrentProcessStatus != null)
                            {
                                frmMain.g_oDelegate.SetControlPropertyValue((Control)m_lblCurrentProcessStatus, "ForeColor", Color.Red);
                                frmMain.g_oDelegate.SetControlPropertyValue((Control)m_lblCurrentProcessStatus, "Text", "!!Cancelled!!");

                            }
                            frmMain.g_oDelegate.SetControlPropertyValue((Control)btnCancel, "Text", "Start");
							this.m_bUserCancel=true;
                            uc_filesize_monitor1.EndMonitoringFile();
                            uc_filesize_monitor2.EndMonitoringFile();

						}
					}
				}
			}
			else
			{
				RunScenario();
			}
		}
        public static void UpdateThermPercent(ProgressBarBasic.ProgressBarBasic p_oPb, int p_intMin, int p_intMax, int p_intValue)
        {
            p_oPb.TextStyle = ProgressBarBasic.ProgressBarBasic.TextStyleType.Percentage;
            int intPercent = (int)(((double)(p_intValue - p_intMin) /
                (double)(p_intMax - p_intMin)) * 100);

            frmMain.g_oDelegate.SetControlPropertyValue(p_oPb, "Value", intPercent);

            frmMain.g_oDelegate.ExecuteControlMethod(p_oPb, "Refresh");

        }
        public static void UpdateThermText(ProgressBarBasic.ProgressBarBasic p_oPb, string p_strText)
        {
           p_oPb.TextStyle = ProgressBarBasic.ProgressBarBasic.TextStyleType.Text;
           frmMain.g_oDelegate.SetControlPropertyValue(p_oPb,"Text",p_strText);
           frmMain.g_oDelegate.ExecuteControlMethod(p_oPb, "Refresh");
        }
        public static int GetListViewItemIndex(ListViewEmbeddedControls.ListViewEx p_oLv, string p_strDesc)
        {
            
            int intCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(p_oLv,"Count",false);
            for (int x = 0; x <= intCount - 1; x++)
            {
                string strDesc = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(p_oLv, x, 2, "Text", false);
                if (strDesc.Trim().ToUpper() == p_strDesc.Trim().ToUpper())
                {
                    return x;
                }
            }
            return -1;
        }

        private void RunScenario()
		{
            int x;

			this.lblMsg.Text = "";
            this.btnCancel.Enabled = false;
            this.btnViewAuditTables.Enabled = false;
            this.btnViewLog.Enabled = false;
            this.btnViewResultsTables.Enabled = false;
            this.btnSqlite.Enabled = false;

            for (x = 0; x <= listViewEx1.Items.Count - 1; x++)
            {
                ProgressBarBasic.ProgressBarBasic pb = new ProgressBarBasic.ProgressBarBasic();
               
                pb = (ProgressBarBasic.ProgressBarBasic)listViewEx1.GetEmbeddedControl(1, x);
                pb.TextStyle = ProgressBarBasic.ProgressBarBasic.TextStyleType.Percentage;
                pb.TextColor = Color.Black;
                frmMain.g_oDelegate.SetControlPropertyValue(pb, "Maximum", 100);

                frmMain.g_oDelegate.SetControlPropertyValue(pb, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue(pb, "Value", 0);
                frmMain.g_oDelegate.SetControlPropertyValue(pb, "Text", "0%");
            }


			this.Refresh();
            this.listViewEx1.Items[0].Selected = true;
            FIA_Biosum_Manager.RunOptimizer.g_bOptimizerRun = true;
            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic =(ProgressBarBasic.ProgressBarBasic)listViewEx1.GetEmbeddedControl(1, 0);
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = 0;

			this.val_OptimizerRunData();
			if (this.m_intError==0)
			{
                this.btnCancel.Enabled = true;
				this.btnCancel.Text = "Cancel";
				this.btnCancel.Refresh();
				this.btnViewAuditTables.Enabled=false;
				this.btnViewResultsTables.Enabled=false;
				this.btnSqlite.Enabled=false;
				this.btnViewLog.Enabled=false;
				this.m_bUserCancel=false;


				frmMain.g_oDelegate.CurrentThreadProcessIdle=false;
				frmMain.g_oDelegate.CurrentThreadProcessAborted=false;
				frmMain.g_oDelegate.CurrentThreadProcessDone=false;
				frmMain.g_oDelegate.CurrentThreadProcessStarted=false;
				frmMain.g_oDelegate.m_oThread = new System.Threading.Thread(new System.Threading.ThreadStart(this.StartScenarioRunProcess));
				frmMain.g_oDelegate.InitializeThreadEvents();
				frmMain.g_oDelegate.m_oThread.IsBackground=true;
				frmMain.g_oDelegate.m_oThread.Start();
			}
			else
			{
                this.btnCancel.Enabled = true;
                this.btnViewAuditTables.Enabled = true;
                this.btnViewLog.Enabled = true;
                this.btnViewResultsTables.Enabled = true;
                this.btnSqlite.Enabled = true;
				if (this.ReferenceOptimizerScenarioForm.WindowState == System.Windows.Forms.FormWindowState.Minimized)
					this.ReferenceOptimizerScenarioForm.WindowState = System.Windows.Forms.FormWindowState.Normal;
				ReferenceOptimizerScenarioForm.Focus();

			}

		}
		private void StartScenarioRunProcess()
		{
			frmMain.g_oDelegate.CurrentThreadProcessStarted=true;
             this.m_oRunOptimizer = new RunOptimizer(this);
			if (!this.m_bAbortThread)
			{
				
			}
			System.Threading.Thread.Sleep(1000);
			
		}
        public static void UpdateThermPercent()
        {
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep++;
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent(
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic,
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps,
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps,
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep);
        }

		/// <summary>
        /// validate each component required for running Optimizer
		/// </summary>
		private void val_OptimizerRunData()
		{
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 10;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;

			this.m_intError=0;

            if (this.m_intError == 0)
            {
                this.m_intError = ReferenceOptimizerScenarioForm.uc_scenario_owner_groups1.ValInput();
            }
           
            if (this.m_intError == 0)
            {
                UpdateThermPercent();
                this.m_intError = ReferenceOptimizerScenarioForm.uc_scenario_costs1.val_costs();
            }

            if (this.m_intError == 0)
            {
                UpdateThermPercent();
                this.m_intError = ReferenceOptimizerScenarioForm.uc_scenario_psite1.val_psites();
            }

            if (this.m_intError == 0)
            {
                UpdateThermPercent();
                this.m_intError = ReferenceOptimizerScenarioForm.uc_scenario_processor_scenario_select1.val_processorscenario();
            }

            if (this.m_intError == 0)
            {
                UpdateThermPercent();
                this.m_intError = ReferenceOptimizerScenarioForm.uc_optimizer_scenario_select_packages1.val_rxPackages();
            }

			if (this.m_intError==0)  
			{
                UpdateThermPercent();
                this.m_intError = ReferenceOptimizerScenarioForm.uc_scenario_filter1.Val_PlotFilter(ReferenceOptimizerScenarioForm.uc_scenario_filter1.txtCurrentSQL.Text.Trim());
                if (this.m_intError!=0)
					MessageBox.Show(ReferenceOptimizerScenarioForm.uc_scenario_filter1.m_strError,"FIA Biosum");
			}
            
			if (this.m_intError==0)  
			{
                UpdateThermPercent();
                this.m_intError = ReferenceOptimizerScenarioForm.uc_scenario_cond_filter1.Val_CondFilter(ReferenceOptimizerScenarioForm.uc_scenario_cond_filter1.txtCurrentSQL.Text.Trim());
                if (this.m_intError!=0)
					MessageBox.Show(ReferenceOptimizerScenarioForm.uc_scenario_cond_filter1.m_strError,"FIA Biosum");
			}
            
			if (this.m_intError==0) 
			{
                UpdateThermPercent();
				this.m_intError = ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_effective1.Audit(false);
				if (this.m_intError!=0)
					MessageBox.Show(ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_strError,"FIA Biosum");
			}
           
			if (this.m_intError==0) 
			{
                UpdateThermPercent();
				ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_optimization1.DisplayAuditMessage=false;
				ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_optimization1.Audit();
				this.m_intError = ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_optimization1.m_intError;
				if (this.m_intError!=0)
					MessageBox.Show(ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_optimization1.m_strError,"FIA Biosum");
			}

			if (this.m_intError==0) 
			{
                UpdateThermPercent();
				this.m_intError = ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.Audit(false);
				if (this.m_intError!=0)
					MessageBox.Show(ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.m_strError,"FIA Biosum");
			}

            if (this.m_intError == 0)
            {
                UpdateThermPercent();
                this.m_intError = val_travelTimesDatasourceMatch();
            }

            if (this.m_intError == 0)
            {
                /***************************************************************************
                **make sure all the scenario datasource tables and files are available
                **and ready for use
                ***************************************************************************/
                UpdateThermPercent();
                if (ReferenceOptimizerScenarioForm.m_ldatasourcefirsttime == true)
                {
                    ReferenceOptimizerScenarioForm.uc_datasource1.LoadValues();
                    ReferenceOptimizerScenarioForm.m_ldatasourcefirsttime = false;
                }
                this.m_intError = ReferenceOptimizerScenarioForm.uc_datasource1.val_datasources();

                if (this.m_intError == 0)
                {
                    UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");

                    FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = 1;
                    listViewEx1.Items[1].Selected = true;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)listViewEx1.GetEmbeddedControl(1, 1);

                    ReferenceOptimizerScenarioForm.SaveRuleDefinitions();

                    UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
                }
            }
            else
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
            }
		}

        private int val_travelTimesDatasourceMatch()
        {
            DataMgr oDataMgr = new DataMgr();
            string strProcessorScenarioTravelTimesSource = ReferenceOptimizerScenarioForm.uc_datasource1.m_strProjectDirectory + "\\" + Tables.TravelTime.DefaultTravelTimePathAndDbFile;

            int x = 0;
            for (x = 0; x <= ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Count - 1; x++)
            {
                if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).Selected)
                {
                    string strProcessorRuleDefsDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;

                    string strProcessorScenarioId = ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).ScenarioId;

                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strProcessorRuleDefsDb)))
                    {
                        conn.Open();

                        oDataMgr.m_strSQL = "SELECT path, file FROM scenario_datasource WHERE TRIM(UPPER(table_type)) = 'TRAVEL TIMES' AND TRIM(UPPER(scenario_id)) = '" + strProcessorScenarioId.Trim().ToUpper() + "'";

                        oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
                        if (oDataMgr.m_DataReader.HasRows)
                        {
                            while (oDataMgr.m_DataReader.Read())
                            {
                                strProcessorScenarioTravelTimesSource = oDataMgr.m_DataReader["path"].ToString().Trim() + "\\" + oDataMgr.m_DataReader["file"].ToString().Trim();
                                break;
                            }
                        }
                        oDataMgr.m_DataReader.Close();
                    }
                }
            }

            string strOptimizerScenarioTravelTimesSource = ReferenceOptimizerScenarioForm.LoadedQueries.m_oTravelTime.m_strDbFile;

            if (strProcessorScenarioTravelTimesSource.Trim().ToLower() != strOptimizerScenarioTravelTimesSource.Trim().ToLower())
            {
                MessageBox.Show("Run Scenario Failed: Optimizer Scenario Travel Times Datasource does not match Travel Times Datasource of selected Processor scenario",
                    "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return -1;
            }

            return 0;
        }

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
            groupBox1_Resize();
		}
        public void groupBox1_Resize()
        {
            this.listViewEx1.Width = this.groupBox1.Width - (int)(this.listViewEx1.Left * 2);
            this.lblMsg.Width = this.listViewEx1.Width;
            this.pnlFileSizeMonitor.Top = this.groupBox1.Height - this.groupBox1.Top - this.pnlFileSizeMonitor.Height - 2;
            this.lblMsg.Top = pnlFileSizeMonitor.Top - this.lblMsg.Height - 2;
            this.listViewEx1.Height = this.lblMsg.Top - this.listViewEx1.Top -5;

            if (uc_filesize_monitor1.lblMaxSize.Left + uc_filesize_monitor1.lblMaxSize.Width > uc_filesize_monitor1.Width)
            {
                for (; ; )
                {
                    uc_filesize_monitor1.Width = uc_filesize_monitor1.Width + 1;
                    if (uc_filesize_monitor1.lblMaxSize.Left + uc_filesize_monitor1.lblMaxSize.Width < uc_filesize_monitor1.Width)
                        break;

                }
            }
            if (uc_filesize_monitor2.lblMaxSize.Left + uc_filesize_monitor2.lblMaxSize.Width > uc_filesize_monitor2.Width)
            {
                for (; ; )
                {
                    uc_filesize_monitor2.Width = uc_filesize_monitor2.Width + 1;
                    if (uc_filesize_monitor2.lblMaxSize.Left + uc_filesize_monitor2.lblMaxSize.Width < uc_filesize_monitor2.Width)
                        break;

                }
            }
            if (uc_filesize_monitor3.lblMaxSize.Left + uc_filesize_monitor3.lblMaxSize.Width > uc_filesize_monitor3.Width)
            {
                for (; ; )
                {
                    uc_filesize_monitor3.Width = uc_filesize_monitor3.Width + 1;
                    if (uc_filesize_monitor3.lblMaxSize.Left + uc_filesize_monitor3.lblMaxSize.Width < uc_filesize_monitor3.Width)
                        break;

                }
            }

            if (uc_filesize_monitor4.lblMaxSize.Left + uc_filesize_monitor4.lblMaxSize.Width > uc_filesize_monitor4.Width)
            {
                for (; ; )
                {
                    uc_filesize_monitor4.Width = uc_filesize_monitor4.Width + 1;
                    if (uc_filesize_monitor4.lblMaxSize.Left + uc_filesize_monitor4.lblMaxSize.Width < uc_filesize_monitor4.Width)
                        break;

                }
            }

            this.uc_filesize_monitor2.Left = this.uc_filesize_monitor1.Left + uc_filesize_monitor2.Width + 2;
            this.uc_filesize_monitor3.Left = this.uc_filesize_monitor2.Left + uc_filesize_monitor3.Width + 2;
            this.uc_filesize_monitor4.Left = this.uc_filesize_monitor3.Left + uc_filesize_monitor4.Width + 2;
        }

        public void UpdateOptimizationVariableGroupboxText(string p_strOptimizationVariableName)
        {
            
        }

		public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
		{
            get { return _frmScenario; }
			set {_frmScenario=value;}
		}

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }
	}
	/// <summary>
    /// main class used for running the Optimizer scenario
	/// </summary>
	public class RunOptimizer
	{
		FIA_Biosum_Manager.uc_optimizer_scenario_run _uc_scenario_run;
		FIA_Biosum_Manager.frmOptimizerScenario _frmScenario;
		private int m_intError;
        public DataMgr m_dataMgr;
		public string m_strSystemResultsDbPathAndFile="";
        public string m_strContextDbPathAndFile = "";
        public string m_strFvsContextDbPathAndFile = "";
        public string m_strFVSPreValidComboDbPathAndFile = "";
        public string m_strFVSPostValidComboDbPathAndFile = "";
        public string m_strFVSPrePostValidComboDbPathAndFile = "";
        public string m_strWorkTablesDb = "";
        public string m_strWorkTablesConn = "";
        
		private env m_oEnv;
		private utils m_oUtils;
		public string m_strPlotTable;
        public string m_strPlotPathAndFile;
		public string m_strRxTable;
        public string m_strRxPathAndFile;
        public string m_strRxPackageTable;
		public string m_strTravelTimeTable;
		public string m_strCondTable;
        public string m_strCondPathAndFile;
        public string m_strHarvestMethodsTable;
        public string m_strHarvestMethodsPathAndFile;
		public string m_strFFETable;
		public string m_strHvstCostsTable;
		public string m_strPSiteWorkTable;
		public string m_strTreeVolValBySpcDiamGroupsTable;
        public string m_strTreeVolValSumTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsTreeVolValSumTableName;
		public string m_strUserDefinedPlotSQL;
	    public string m_strUserDefinedCondSQL;
        public string m_strPSiteTable;
        public string m_strPSitePathAndFile;
        public string m_strProcessorResultsPathAndFile;
        private string m_strEconByRxWorkTableName = Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName + "_work_table";
        
		
		private string m_strLine;
        private OptimizerScenarioItem.OptimizationVariableItem m_oOptimizationVariable = new OptimizerScenarioItem.OptimizationVariableItem();
		private string m_strOptimizationTableName="";
		private string m_strOptimizationSourceTableName="";
		private string m_strOptimizationTableNameSql="";
		private string m_strOptimizationColumnNameSql="";
		private string m_strOptimizationSourceColumnName="";
		private string m_strOptimizationAggregateSql="";
		private string m_strOptimizationAggregateColumnName="";
        private ProcessorScenarioItem m_oProcessorScenarioItem = null;
        private RxTools m_oRxTools = new RxTools();
        private RxPackageItem_Collection m_oRxPackageItem_Collection = new RxPackageItem_Collection();
		private FIA_Biosum_Manager.macrosubst m_oVarSub = new macrosubst();

        public static bool g_bOptimizerRun = false;
        public static ProgressBarBasic.ProgressBarBasic g_oCurrentProgressBarBasic = null;
        public static int g_intCurrentProgressBarBasicMinimumSteps = 1;
        public static int g_intCurrentProgressBarBasicMaximumSteps = -1;
        public static int g_intCurrentProgressBarBasicCurrentStep = -1;
        public static int g_intCurrentListViewItem = 0;
        System.Windows.Forms.CheckBox oCheckBox = null;
        int intListViewIndex = -1;
        private string m_strDebugFile = "";

        public RunOptimizer(FIA_Biosum_Manager.uc_optimizer_scenario_run p_form)
        {
            m_intError = 0;
            _uc_scenario_run = p_form;
            ReferenceOptimizerScenarioForm = p_form.ReferenceOptimizerScenarioForm;
            try
            {
                m_strDebugFile = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\runlog.txt";


                if (frmMain.g_bDebug)
                {
                    if (System.IO.File.Exists(m_strDebugFile)) System.IO.File.Delete(m_strDebugFile);
                    m_strLine = "START: Optimizer Run Log " + System.DateTime.Now.ToString();
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_strLine + "\r\n\r\n");
                    if (frmMain.g_intDebugLevel > 1)
                    {

                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Project: " + frmMain.g_oFrmMain.frmProject.uc_project1.txtProjectId.Text + "\r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Project Directory: " + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text + "\r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Scenario Directory: " + ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "---------------------------------------------------------------\r\n");
                    }

                }
            }
            catch (Exception caught)
            {
                MessageBox.Show(caught.ToString());
                m_intError = -1;
                return;
            }
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Create A Temporary MDB File With Links To All The Optimizer Tables And Scenario Result Tables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "------------------------------------------------------------------------------------------------\r\n");
            }
            try
            {
                //
                //INITIALIZE AND LOAD VARIABLES
                //
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, 2);
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = 2;
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 15;
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;

                frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, 2);
                frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, 2, "Selected", true);
                frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, 2, "focused", true);

                m_oUtils = new utils();
                m_oEnv = new env();
                m_dataMgr = new DataMgr();

                //get the selected processor scenario item
                if (ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_processor_scenario_select1.m_oProcessorScenarioItem != null)
                    m_oProcessorScenarioItem = this.ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_processor_scenario_select1.m_oProcessorScenarioItem;

                m_strProcessorResultsPathAndFile = m_oProcessorScenarioItem.DbPath + "\\" + Tables.ProcessorScenarioRun.DefaultScenarioResultsTableDbFile;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //load the treatment packages
                m_oRxTools.LoadAllRxPackageItems(m_oRxPackageItem_Collection);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                m_oVarSub.ReferenceSQLMacroSubstitutionVariableCollection =
                    frmMain.g_oSQLMacroSubstitutionVariable_Collection;
                m_oVarSub.ReferenceGeneralMacroSubstitutionVariableCollection =
                    frmMain.g_oGeneralMacroSubstitutionVariable_Collection;

                string strScenarioOutputFolder = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim();
                m_strSystemResultsDbPathAndFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
                CopyScenarioResultsTable(m_strSystemResultsDbPathAndFile, strScenarioOutputFolder + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsSqliteDbFile);

                intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                    ReferenceUserControlScenarioRun.listViewEx1, "Populate Context Database");
                oCheckBox = (CheckBox)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(0, intListViewIndex);
                if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)oCheckBox, "Checked", false) == true)
                {
                    m_strContextDbPathAndFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
                    CopyScenarioResultsTable(m_strContextDbPathAndFile, strScenarioOutputFolder + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsContextSqliteDbFile);
                }

                m_strFVSPrePostValidComboDbPathAndFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
                CopyScenarioResultsTable(m_strFVSPrePostValidComboDbPathAndFile, ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\validcombo.db");

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                m_strUserDefinedPlotSQL =
                    m_oVarSub.SQLTranslateVariableSubstitution(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_filter1.txtCurrentSQL.Text.Trim());

                m_strUserDefinedCondSQL =
                    m_oVarSub.SQLTranslateVariableSubstitution(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_cond_filter1.txtCurrentSQL.Text.Trim());

                // Create temporary SQLite database for work tables
                m_strWorkTablesDb = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
                m_dataMgr.CreateDbFile(m_strWorkTablesDb);
                m_strWorkTablesConn = m_dataMgr.GetConnectionString(m_strWorkTablesDb);

                if (m_strWorkTablesDb.Trim().Length == 0)
                {
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!RunOptimizer: Error Creating DB File Containing Work Tables!!\r\n");
                    MessageBox.Show("RunOptimizer: Error Creating DB File Containing Work Tables");
                }
                else
                {
                    if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.UserControl)_uc_scenario_run.uc_filesize_monitor1, "Visible", false) == false)
                    {
                        _uc_scenario_run.uc_filesize_monitor1.BeginMonitoringFile(
                            m_strWorkTablesDb, 2000000000, "2GB");
                        _uc_scenario_run.uc_filesize_monitor1.Information = "Work table DB file";

                        _uc_scenario_run.uc_filesize_monitor2.BeginMonitoringFile(
                            m_strSystemResultsDbPathAndFile, 2000000000, "2GB");
                        _uc_scenario_run.uc_filesize_monitor2.Information = "Scenario results DB file";
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Work Db File: " + m_strWorkTablesDb + "\r\n");
                    ReferenceUserControlScenarioRun.btnSqlite.Enabled = false;

                    getTableNames();
                    if (m_intError != 0)
                    {
                        MessageBox.Show("An error occurred while retrieving Treatment Optimizer table names");
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    //effective variable
                    FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_effective.Variables oFvsVar =
                        ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar;

                    /********************************************************************
					 **get optimization variable
					 ********************************************************************/
                    OptimizerScenarioItem.OptimizationVariableItem_Collection oOptimizationVariableCollection =
                        ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection;

                    for (int x = 0; x <= oOptimizationVariableCollection.Count - 1; x++)
                    {
                        if (oOptimizationVariableCollection.Item(x).bSelected)
                        {
                            m_oOptimizationVariable.Copy(oOptimizationVariableCollection.Item(x), ref m_oOptimizationVariable);
                            break;
                        }
                    }

                    if (m_oOptimizationVariable.strMaxYN == "Y")
                    {
                        m_strOptimizationTableName = "Max";
                        m_strOptimizationAggregateSql = "MAX";
                        m_strOptimizationAggregateColumnName = "max_optimization_value";
                    }
                    else
                    {
                        m_strOptimizationTableName = "Min";
                        m_strOptimizationAggregateSql = "MIN";
                        m_strOptimizationAggregateColumnName = "min_optimization_value";
                    }

                    if (m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "REVENUE")
                    {
                        m_strOptimizationTableName = this.m_strOptimizationTableName + "NR";
                        m_strOptimizationSourceTableName = Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;
                        m_strOptimizationTableNameSql = "cycle1_effective_" + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;
                        m_strOptimizationSourceColumnName = "max_nr_dpa";
                        m_strOptimizationColumnNameSql = "post_variable_value";
                    }
                    else if (m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "MERCHANTABLE VOLUME")
                    {
                        m_strOptimizationSourceTableName = Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;
                        m_strOptimizationTableNameSql = "cycle1_effective_" + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;
                        m_strOptimizationColumnNameSql = "post_variable_value";
                        m_strOptimizationSourceColumnName = "merch_vol_cf";
                        m_strOptimizationTableName = m_strOptimizationTableName + "MerchVol";
                        if (m_oOptimizationVariable.bUseFilter) m_strOptimizationTableName = m_strOptimizationTableName + "NR";
                    }
                    else if (m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "ECONOMIC ATTRIBUTE")
                    {
                        m_strOptimizationColumnNameSql = "post_variable_value";
                        m_strOptimizationTableName = m_oOptimizationVariable.strFVSVariableName;
                        if (m_oOptimizationVariable.bUseFilter) m_strOptimizationTableName = m_strOptimizationTableName + "NR";
                    }
                    else
                    {
                        string[] strCol = frmMain.g_oUtils.ConvertListToArray(m_oOptimizationVariable.strFVSVariableName, ".");
                        m_strOptimizationSourceTableName = strCol[0];
                        m_strOptimizationSourceColumnName = strCol[1];
                        m_strOptimizationTableNameSql = "cycle1_effective_optimization_treatments";
                        if (m_oOptimizationVariable.strValueSource == "POST")
                        {
                            m_strOptimizationColumnNameSql = "post_variable_value";
                        }
                        else if (m_oOptimizationVariable.strValueSource == "POST-PRE")
                        {
                            m_strOptimizationColumnNameSql = "change_value";
                        }

                        m_strOptimizationTableName = m_strOptimizationTableName + strCol[1].Trim();
                        if (m_oOptimizationVariable.bUseFilter) m_strOptimizationTableName = m_strOptimizationTableName + "NR";
                    }

                    //revenue filter
                    if (m_oOptimizationVariable.bUseFilter)
                    {
                        switch (m_oOptimizationVariable.strFilterOperator)
                        {
                            case "<":
                                m_strOptimizationTableName = m_strOptimizationTableName + "lt";
                                break;
                            case ">":
                                m_strOptimizationTableName = m_strOptimizationTableName + "gt";
                                break;
                            case "<=":
                                m_strOptimizationTableName = m_strOptimizationTableName + "lte";
                                break;
                            case ">=":
                                m_strOptimizationTableName = m_strOptimizationTableName + "gte";
                                break;
                            case "<>":
                                m_strOptimizationTableName = m_strOptimizationTableName + "ne";
                                break;
                            case "=":
                                m_strOptimizationTableName = m_strOptimizationTableName + "e";
                                break;
                        }
                        if (m_oOptimizationVariable.dblFilterValue >= 0)
                            m_strOptimizationTableName = m_strOptimizationTableName + "P" + Convert.ToString(m_oOptimizationVariable.dblFilterValue).Trim();
                        else
                            m_strOptimizationTableName = m_strOptimizationTableName + "N" + Convert.ToString(m_oOptimizationVariable.dblFilterValue * -1).Trim();
                    }

                    ReferenceOptimizerScenarioForm.OutputTablePrefix = getFileNamePrefix();
                    m_strOptimizationTableName = ReferenceOptimizerScenarioForm.OutputTablePrefix + "_" + m_strOptimizationTableName;

                    CreateAuditTables();
                    CreateOptimizerResultTables();
                    if (!String.IsNullOrEmpty(m_strContextDbPathAndFile))
                        CreateContextTables();
                    CreateValidComboTables();

                    if (m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    // Processor scenario results tables
                    m_strTreeVolValBySpcDiamGroupsTable = Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName;
                    m_strHvstCostsTable = Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName;

                    //CREATE WORK TABLES
                    CreateTableStructureOfHarvestCosts();
                    if (m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    CreateTableStructureForIntensity();
                    if (m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    CreateTableStructureForScenarioProcessingSites();
                    if (m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    CreateTableStructureForHaulCosts();
                    if (m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    CreateTableStructureForPlotCondAccessiblity();
                    if (m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    /******************************************************************
					**create table structure of user defined sql plot filter statement
					******************************************************************/
                    CreateTableStructureOfUserDefinedPlotSQL();

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    /********************************************************************
                     **create table structure for condition table filters
                     ********************************************************************/
                    CreateTableStructureForUserDefinedConditionTable();

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    /********************************************************************
                     **filter scenario selected processing sites
                     ********************************************************************/
                    FilterPSites();

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    if (m_intError == 0)
                    {
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
                    }
                    else
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    }

                    /***********************************************************
					**identify the plots that are accessible
					***********************************************************/
                    FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem + 1;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1,
                        FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                    frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                    frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
                    frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);

                    if (m_intError == 0)
                    {
                        CondAccessible();
                    }
                    else
                    {
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "NA");
                    }

                    /**************************************************************
                     **get the fastest travel time from plot to processing site
                     **************************************************************/
                    intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                       ReferenceUserControlScenarioRun.listViewEx1, "Get Least Expensive Route From Stand To Wood Processing Facility");

                    FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1,
                        FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                    frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                    frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
                    frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);

                    if (m_intError == 0)
                    {
                        getHaulCosts();
                    }
                    else
                    {
                        if (ReferenceUserControlScenarioRun.m_bUserCancel == false)
                        {
                            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "NA");
                        }
                    }

                    /***************************************************************************
                     **sum up tree volumes and values by plot+condition, treatment and species
                     ***************************************************************************/
                    intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                       ReferenceUserControlScenarioRun.listViewEx1, "Sum Tree Yields, Volume, And Value For A Stand And Treatment");

                    FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1,
                        FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                    frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                    frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
                    frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);

                    if (m_intError == 0)
                    {
                        sumTreeVolVal();
                    }
                    else
                    {
                        if (ReferenceUserControlScenarioRun.m_bUserCancel == false)
                        {
                            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "NA");
                        }
                    }

                    /***************************************************************************
                     **valid combos
                     ***************************************************************************/
                    if (m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                    {
                        validcombos();
                    }

                    /***************************************************************************
                     **cond psite table
                     ***************************************************************************/
                    if (m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                    {
                        CondPsiteTable();
                    }

                    /***************************************************************************
                     **context reference database
                     ***************************************************************************/
                    if (m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false &&
                        !String.IsNullOrEmpty(m_strContextDbPathAndFile))
                    {
                        ContextReferenceTables();
                        ContextTextFiles(strScenarioOutputFolder);
                    }

                    /**********************************************************************
                     **wood product yields net revenue and costs summary by treatment table
                     **********************************************************************/
                    if (m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                    {
                        econ_by_rx_cycle();
                    }

                    /*******************************************************************************
                     **wood product yields net revenue and costs summary by treatment package table 
                     *******************************************************************************/
                    if (m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                    {
                        econ_by_rx_utilized_sum();
                    }

                    /**********************************************************************
                     **Calculate custom economic variables if needed
                     **********************************************************************/
                    if (m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                    {
                        calculate_weighted_econ_variables();
                    }

                    /***************************************************************************
                     **effective treatments
                     ***************************************************************************/
                    if (m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                    {
                        Effective();
                    }
                    /**************************************************************************
                     **optimization
                     **************************************************************************/

                    if (m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                    {
                        Optimization();
                    }

                    /**************************************************************************
                     **tie breakers
                     **************************************************************************/
                    if (m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                    {
                        tiebreaker();
                    }

                    /*********************************************************************
                     **find the best treatments for revenue, torch/crown index improvement,
                     **and merch removal
                     *********************************************************************/
                    if (m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                    {
                        Best_rx_summary();
                    }

                    if (m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                    {
                        System.DateTime oDate = System.DateTime.Now;
                        string strDateFormat = "yyyy-MM-dd_HH-mm";
                        string strFileDate = oDate.ToString(strDateFormat);
                        strFileDate = strFileDate.Replace("/", "_"); strFileDate = strFileDate.Replace(":", "_");
                        CreateHtml();
                        CopyScenarioResultsTable(strScenarioOutputFolder + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsSqliteDbFile, m_strSystemResultsDbPathAndFile);
                        if (!String.IsNullOrEmpty(m_strContextDbPathAndFile))
                            CopyScenarioResultsTable(strScenarioOutputFolder + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsContextSqliteDbFile, m_strContextDbPathAndFile);
                        m_strSystemResultsDbPathAndFile = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\" +
                            Tables.OptimizerScenarioResults.DefaultScenarioResultsSqliteDbFile;
                        m_strFVSPrePostValidComboDbPathAndFile = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\validcombo_fvspre.db";

                        CopyScenarioResultsTable(
                            ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\optimizer_results_" + m_strOptimizationTableName + "_" + strFileDate.Trim() + ".db",
                            ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsSqliteDbFile);
                    }

                    if (frmMain.g_bDebug)
                    {
                        m_strLine = "END: Optimizer Analysis Run Log " + System.DateTime.Now.ToString();
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_strLine + "\r\n\r\n");
                    }

                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", false);
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnCancel, "Text", "Start");
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnViewResultsTables, "Enabled", true);
                    if (m_intError == 0) frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnSqlite, "Enabled", true);

                    if (m_intError != 0)
                    {
                        MessageBox.Show("Completed With Errors");
                    }
                    else
                    {
                        if (ReferenceUserControlScenarioRun.m_bUserCancel == false)
                            MessageBox.Show("Successfully Completed");
                        else
                            MessageBox.Show("Process Cancelled");
                    }
                }

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnViewLog, "Enabled", true);

                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, 
                    FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
                frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);

                //audit
                intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                           ReferenceUserControlScenarioRun.listViewEx1, "Populate Valid Combination Audit Data");
                oCheckBox = (CheckBox)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(
                    0, intListViewIndex);
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
                if ((bool)frmMain.g_oDelegate.GetControlPropertyValue(oCheckBox, "Enabled", false) == true)
                {
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnViewAuditTables, "Enabled", true);
                }
            }
            catch (Exception err)
            {
                if (err.Message.Trim().ToUpper() != "THREAD WAS BEING ABORTED.")
                {
                    MessageBox.Show("Run Scenario " + err.Message, "FIA Biosum");
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.btnCancel, "Text", "Start");
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!Error!!\r\n");
                }
            }

            _uc_scenario_run.uc_filesize_monitor1.EndMonitoringFile();
            _uc_scenario_run.uc_filesize_monitor2.EndMonitoringFile();
            if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.UserControl)_uc_scenario_run.uc_filesize_monitor3, "Visible", false))
                _uc_scenario_run.uc_filesize_monitor3.EndMonitoringFile();
            if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.UserControl)_uc_scenario_run.uc_filesize_monitor4, "Visible", false))
                _uc_scenario_run.uc_filesize_monitor4.EndMonitoringFile();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.ReferenceUserControlScenarioRun.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        }


        public RunOptimizer()
		{

		}
		~RunOptimizer()
		{
			
		}

        private void CreateAuditTables()
        {
            DataMgr oDataMgr = new DataMgr();

            string strDbPathAndFile = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\" + Tables.Audit.DefaultCondAuditTableSqliteDbFile;
            if (!System.IO.File.Exists(strDbPathAndFile))
            {
                oDataMgr.CreateDbFile(strDbPathAndFile);
            }

            string strConn = oDataMgr.GetConnectionString(strDbPathAndFile);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                if (!oDataMgr.TableExist(oConn, Tables.Audit.DefaultCondAuditTableName))
                {
                    frmMain.g_oTables.m_oAudit.CreateCondAuditTableSqlite(oDataMgr, oConn, Tables.Audit.DefaultCondAuditTableName);
                }
                if (!oDataMgr.TableExist(oConn, Tables.Audit.DefaultCondRxAuditTableName))
                {
                    frmMain.g_oTables.m_oAudit.CreatePlotCondRxAuditTableSqlite(oDataMgr, oConn, Tables.Audit.DefaultCondRxAuditTableName);
                }
            }
            oDataMgr = null;
        }
		/// <summary>
		/// create a table structure that will hold
		/// the plot data that results when running the user 
		/// defined sql
		/// </summary>
		/// <param name="strUserDefinedSQL"></param>
        private void CreateTableStructureOfUserDefinedPlotSQL()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureOfUserDefinedPlotSQL\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Create table structure userdefinedplotfilter_work\r\n");

                if (p_dataMgr.TableExist(workConn, "userdefinedplotfilter_work"))
                {
                    p_dataMgr.m_strSQL = "DROP TABLE userdefinedplotfilter_work";
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    if (p_dataMgr.m_intError != 0)
                    {
                        if (frmMain.g_bDebug)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "!! Error Deleting userdefinedplotfilter_work Table!!\r\n");
                        this.m_intError = p_dataMgr.m_intError;
                        p_dataMgr = null;
                        return;
                    }
                }

                // Attach all databases available for use for plot filter
                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strPlotPathAndFile + "' AS master";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strPSitePathAndFile + "' AS gis";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strHarvestMethodsPathAndFile + "' AS ref";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                string[] arrFields = p_dataMgr.getFieldNamesArray(workConn, this.m_strUserDefinedPlotSQL);

                string fieldsAndDataTypes = "";

                foreach (string column in arrFields)
                {
                    string field = "";
                    string dataType = "";
                    p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT " + column + " FROM plot", ref field, ref dataType);
                    dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                    fieldsAndDataTypes = fieldsAndDataTypes + field + " " + dataType + ", ";
                }
                p_dataMgr.m_strSQL = "CREATE TABLE userdefinedplotfilter_work (" + fieldsAndDataTypes + "PRIMARY KEY (biosum_plot_id))";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strSystemResultsDbPathAndFile + "' AS results";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                /***********************************************************************
			     **make a copy of the userdefinedplot filter table and give it the
			     **name ruledefinitionsplotfilter. This will apply the owngrpcd
			     **filters and any other future filters to the userdefinedplotfilter_work table.
			     ***********************************************************************/

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Delete table ruledefinitionsplotfilter\r\n");

                if (p_dataMgr.AttachedTableExist(workConn, "ruledefinitionsplotfilter"))
                {
                    p_dataMgr.m_strSQL = "DROP TABLE results.ruledefinitionsplotfilter";
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    if (p_dataMgr.m_intError != 0)
                    {
                        if (frmMain.g_bDebug)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "!! Error Deleting ruledefinitionsplotfilter Table!!\r\n");
                        this.m_intError = p_dataMgr.m_intError;
                        p_dataMgr = null;
                        return;
                    }
                }

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Copy table structure userdefinedplotfilter_work to ruledefinitionsplotfilter\r\n");

                arrFields = p_dataMgr.getFieldNamesArray(workConn, "SELECT * FROM userdefinedplotfilter_work");

                fieldsAndDataTypes = "";

                foreach (string column in arrFields)
                {
                    string field = "";
                    string dataType = "";
                    p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT " + column + " FROM userdefinedplotfilter_work", ref field, ref dataType);
                    dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                    fieldsAndDataTypes = fieldsAndDataTypes + field + " " + dataType + ", ";
                }
                p_dataMgr.m_strSQL = "CREATE TABLE results.ruledefinitionsplotfilter (" + fieldsAndDataTypes + "PRIMARY KEY (biosum_plot_id))";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
            }
        }
		/// <summary>
		/// create links to the tables located in the optimizer_results.accdb file
		/// </summary>
		
        private void CreateOptimizerResultTables()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateOptimizerResultsTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            string[] strTableNames = new string[1];
            DataMgr oDataMgr = new DataMgr();
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(m_strSystemResultsDbPathAndFile)))
            {
                conn.Open();
                strTableNames = oDataMgr.getTableNames(conn);
                for (int x = 0; x <= strTableNames.Length - 1; x++)
                {
                    if (strTableNames[x] != null && strTableNames[x].Trim().Length > 0)
                    {
                        oDataMgr.SqlNonQuery(conn, "DROP TABLE " + strTableNames[x]);
                    }
                }

                // Query the optimization variable for the selected revenue attribute so we can pass it to the table
                string strColumnFilterName = "";
                if (this.m_oOptimizationVariable.bUseFilter)
                {
                    strColumnFilterName = this.m_oOptimizationVariable.strRevenueAttribute;
                }

                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteValidComboTable(oDataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsValidCombosTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteValidComboFVSPrePostTable(oDataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsValidCombosFVSPrePostTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteTieBreakerTable(oDataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsTieBreakerTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteOptimizationTable(oDataMgr, conn, this.ReferenceOptimizerScenarioForm.OutputTablePrefix + Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix, strColumnFilterName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteEffectiveTable(oDataMgr, conn, this.ReferenceOptimizerScenarioForm.OutputTablePrefix + Tables.OptimizerScenarioResults.DefaultScenarioResultsEffectiveTableSuffix, strColumnFilterName);
                string strScenarioResultsBestRxSummaryTableName = this.ReferenceOptimizerScenarioForm.OutputTablePrefix + Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryTableSuffix;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteBestRxSummaryTable(oDataMgr, conn, strScenarioResultsBestRxSummaryTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteBestRxSummaryTable(oDataMgr, conn, strScenarioResultsBestRxSummaryTableName + "_before_tiebreaks");
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteProductYieldsTable(oDataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteEconByRxUtilSumTable(oDataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqlitePostEconomicWeightedTable(oDataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteHaulCostTable(oDataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsHaulCostsTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteCondPsiteTable(oDataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsCondPsiteTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteVersionTable(oDataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsVersionTableName);

            }
        }


        private void CreateContextTables()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateContextTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            string[] strTableNames;
            strTableNames = new string[1];
            DataMgr p_dataMgr = new DataMgr();

            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(p_dataMgr.GetConnectionString(this.m_strContextDbPathAndFile)))
            {
                conn.Open();

                strTableNames = p_dataMgr.getTableNames(conn);
                for (int x = 0; x <= strTableNames.Length - 1; x++)
                {
                    if (strTableNames[x] != null &&
                        strTableNames[x].Trim().Length > 0)
                    {
                        p_dataMgr.SqlNonQuery(conn, "DROP TABLE " + strTableNames[x]);
                    }
                }

                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteDiameterSpeciesGroupRefTable(p_dataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsDiameterSpeciesGroupRefTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteFvsWeightedVariableRefTable(p_dataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsFvsWeightedVariablesRefTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteEconWeightedVariableRefTable(p_dataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsEconWeightedVariablesRefTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteSpeciesGroupRefTable(p_dataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsSpeciesGroupRefTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateSqliteScenarioAdditionalHarvestCostsTable(p_dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName + "_C");
                frmMain.g_oTables.m_oProcessor.CreateSqliteHarvestCostsTable(p_dataMgr, conn, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + "_C");
                frmMain.g_oTables.m_oProcessor.CreateSqliteTreeVolValSpeciesDiamGroupsTable(p_dataMgr, conn, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + "_C", true);
                frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioPSitesTable(p_dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableName + "_C");
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteHarvestMethodRefTable(p_dataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteRxPackageRefTable(p_dataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsRxPackageRefTableName);
                frmMain.g_oTables.m_oFvs.CreateSqliteRxHarvestCostColumnTable(p_dataMgr, conn, Tables.FVS.DefaultRxHarvestCostColumnsTableName + "_C");

                // Add the ad hoc additional harvest cost columns to table
                string strProcessorPath = ((frmMain)this._frmScenario.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + strProcessorPath + "' AS processor_defs";
                p_dataMgr.SqlNonQuery(conn, p_dataMgr.m_strSQL);

                string strSourceColumnsList = p_dataMgr.getFieldNames(conn, "SELECT * FROM scenario_additional_harvest_costs");
                string[] strSourceColumnsArray = frmMain.g_oUtils.ConvertListToArray(strSourceColumnsList, ",");
                foreach (string strColumn in strSourceColumnsArray)
                {
                    if (!p_dataMgr.ColumnExist(conn,
                        Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName + "_C", strColumn))
                    {
                        p_dataMgr.AddColumn(conn, Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName + "_C",
                            strColumn, "DOUBLE", "");
                    }
                }
            }
        }

        /// <summary>
        /// Copy the scenario results db file from the scenario?\db directory to the temp directory
        /// where the temp directory version is used during a single Optimizer run. Once
        /// the run successfully completes it is copied back to the scenario?\db directory.
        /// </summary>
        private void CopyScenarioResultsTable(string p_strDestPathAndDbFileName, string p_strSourcePathAndDbFileName)
        {
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();

            if (System.IO.File.Exists(p_strSourcePathAndDbFileName))
            {
                System.IO.File.Copy(p_strSourcePathAndDbFileName, p_strDestPathAndDbFileName, true);
            }
            else
            {
                oDataMgr.CreateDbFile(p_strDestPathAndDbFileName);
            }
            oDataMgr = null;
        }

        /// <summary>
        /// create links to the tables located in the validcombo.mdb file
        /// </summary>
        private void CreateValidComboTables()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateValidComboTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            string[] strTableNames;
            strTableNames = new string[1];
            DataMgr oDataMgr = new DataMgr();

            //
            //FVS PRE AND POST VALID COMBO TABLE
            //
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(this.m_strFVSPrePostValidComboDbPathAndFile)))
            {
                conn.Open();

                strTableNames = oDataMgr.getTableNames(conn);
                for (int x = 0; x <= strTableNames.Length - 1; x++)
                {
                    if (strTableNames[x] != null &&
                        strTableNames[x].Trim().Length > 0)
                    {
                        oDataMgr.SqlNonQuery(conn, "DROP TABLE " + strTableNames[x]);
                    }
                }
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteValidComboFVSPreTable(oDataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsValidCombosFVSPreTableName);
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteValidComboFVSPostTable(oDataMgr, conn, Tables.OptimizerScenarioResults.DefaultScenarioResultsValidCombosFVSPostTableName);

            }

        }
       
		
		/// <summary>
        /// get the names of the Optimizer tables
		/// </summary>
		private void getTableNames()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nGet Optimizer Table Names\r\n---------------------\r\n");
            }
			/**************************************************************
			 **get the plot table name
			 **************************************************************/
            string[] arr1 = new string[] { "PLOT" };
            object oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourceTableName", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                    this.m_strPlotTable = strValue;
                }
            }
            
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Plot:" + this.m_strPlotTable + "\r\n");

            oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourcePathAndFile", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                    this.m_strPlotPathAndFile = strValue;
                }
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Plot Path and File:" + this.m_strPlotPathAndFile + "\r\n");



			/**************************************************************
			 **get the treatment prescriptions table
			 **************************************************************/
            arr1 = new string[] { "TREATMENT PRESCRIPTIONS" };
            oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourceTableName", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                    this.m_strRxTable = strValue;
                }
            }
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Treatment Prescriptions:" + this.m_strRxTable + "\r\n");

            oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourcePathAndFile", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                    this.m_strRxPathAndFile = strValue;
                }
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Treatment Prescriptions Path and File:" + this.m_strRxPathAndFile + "\r\n");

            /**************************************************************
			 **get the treatment package table
			 **************************************************************/
            arr1 = new string[] { "TREATMENT PACKAGES" };
            oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourceTableName", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                    this.m_strRxPackageTable = strValue;
                }
            }
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Treatment Package:" + this.m_strRxPackageTable + "\r\n");


			/**************************************************************
			 **get the travel time table name
			 **************************************************************/
            arr1 = new string[] { "TRAVEL TIMES" };
            oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourceTableName", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                    this.m_strTravelTimeTable = strValue;
                }
            }
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Travel Time:" + m_strTravelTimeTable + "\r\n");

			
			/**************************************************************
			 **get the cond table name
			 **************************************************************/
            arr1 = new string[] { "CONDITION" };
            oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourceTableName", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                    this.m_strCondTable = strValue;
                }
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Cond:" + this.m_strCondTable + "\r\n");

            oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourcePathAndFile", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                    this.m_strCondPathAndFile = strValue;
                }
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Cond Path and File:" + this.m_strCondPathAndFile + "\r\n");

            /**************************************************************
             **get the processing sites table name and path
             **************************************************************/
            arr1 = new string[] { "PROCESSING SITES" };
            oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourceTableName", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                    this.m_strPSiteTable = strValue;
                }
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Processing Sites Table:" + this.m_strPSiteTable + "\r\n");

            oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourcePathAndFile", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                    this.m_strPSitePathAndFile = strValue;
                }
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Processing Sites Path and file:" + m_strPSitePathAndFile + "\r\n");
            this.m_strTreeVolValSumTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsTreeVolValSumTableName;

            /**************************************************************
             **get the harvest methods table name and path
             **************************************************************/
            arr1 = new string[] { "HARVEST METHODS" };
            oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourceTableName", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                    this.m_strHarvestMethodsTable = strValue;
                }
                else
                {
                    this.m_intError = -1;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Unable to locate Harvest Methods table on data sources tab!! \r\n");
                    return;
                }
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Harvest methods Table:" + this.m_strHarvestMethodsTable + "\r\n");

            oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_datasource1,
                "getDataSourcePathAndFile", arr1, true);
            if (oValue != null)
            {
                string strValue = Convert.ToString(oValue);
                if (strValue != "false")
                {
                    this.m_strHarvestMethodsPathAndFile = strValue;
                }
            }

            this.m_strHarvestMethodsPathAndFile = this.m_oVarSub.GeneralTranslateVariableSubstitution(this.m_strHarvestMethodsPathAndFile);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Harvest Methods Path and file:" + this.m_strHarvestMethodsPathAndFile + "\r\n");
            this.m_strTreeVolValSumTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsTreeVolValSumTableName;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Tree Sum Volume And Value:" + m_strTreeVolValSumTable + "\r\n");

		}
        private void FilterPSites()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//FilterPSites\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();
            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + 
                    Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile + "' AS rule_defs";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "DELETE FROM scenario_psites_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "INSERT INTO scenario_psites_work_table (psite_id,trancd,biocd) " +
                    "SELECT psite_id,trancd,biocd " +
                    "FROM scenario_psites " +
                    "WHERE TRIM(scenario_id)='" + ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim() + "' AND " +
                    "selected_yn='Y';";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
            }

            p_dataMgr = null;
        }
        /// <summary>
        /// set the cond_accessible_yn flag
        /// </summary>
        private void CondAccessible()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CondAccessible\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 7;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;

            DataMgr p_dataMgr = new DataMgr();

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strCondPathAndFile + "' AS master";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                /**************************************************
                 * Insert plot_id/cond_id records into PSITE_ACCESSIBLE_WORKTABLE
                 * ************************************************/
                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName +
                    " (biosum_cond_id, biosum_plot_id) " +
                    "SELECT DISTINCT biosum_cond_id, biosum_plot_id " +
                    "FROM " + m_strCondTable + " AS c " +
                    "GROUP BY c.biosum_cond_id, c.biosum_plot_id";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                /********************************************************************
			     **update cond_too_far_steep field to Y if slope is 
			     **<= 40% and the yarding distance >= to maximum yarding distance 
			     **for a slope <= 40%
			     ********************************************************************/
                p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS w " +
                    "SET cond_too_far_steep_yn = 'Y' " +
                    "FROM " + m_strCondTable + " AS c, " + m_strPlotTable + " AS p " +
                    "WHERE c.biosum_cond_id = w.biosum_cond_id AND c.biosum_plot_id = p.biosum_plot_id " +
                    "AND w.biosum_plot_id = p.biosum_plot_id AND c.slope IS NOT NULL " +
                    "AND c.slope < " + m_oProcessorScenarioItem.m_oHarvestMethod.SteepSlopePercent + " AND " +
                    "p.gis_yard_dist_ft >= " + ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_cond_filter1.strNonSteepYardingDistance.Trim();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                /********************************************************************
			     **update cond_too_far_steep field to Y if slope is 
			     **> 40% and the yarding distance >= to maximum yarding distance 
			     **for a slope > 40%
			     ********************************************************************/
                p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS w " +
                    "SET cond_too_far_steep_yn = 'Y' " +
                    "FROM " + m_strCondTable + " AS c, " + m_strPlotTable + " AS p " +
                    "WHERE c.biosum_cond_id = w.biosum_cond_id AND c.biosum_plot_id = p.biosum_plot_id " +
                    "AND w.biosum_plot_id = p.biosum_plot_id AND c.slope IS NOT NULL " +
                    "AND c.slope >= " + m_oProcessorScenarioItem.m_oHarvestMethod.SteepSlopePercent + " AND " +
                    "p.gis_yard_dist_ft >= " + ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_cond_filter1.strSteepYardingDistance.Trim();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                /*************************************************************
			     **set the remainder of the cond_too_far_steep_yn fields to N
			     *************************************************************/
                p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName +
                    " SET cond_too_far_steep_yn = 'N' " +
                    "WHERE cond_too_far_steep_yn IS NULL " +
                    "OR (cond_too_far_steep_yn <> 'Y' AND cond_too_far_steep_yn <> 'N')";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                /*************************************************************
                 **update the condition accessible flag
                 *************************************************************/
                p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName +
                    " SET cond_accessible_yn = " +
                    "CASE WHEN cond_too_far_steep_yn = 'Y' THEN 'N' ELSE 'Y' END";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                /********************************************************************
                 **get the current condition record counts for each plot
                 ********************************************************************/
                //insert the condition counts into the work table
                p_dataMgr.m_strSQL = "INSERT INTO plot_cond_accessible_work_table (biosum_plot_id, num_cond) " +
                    " SELECT biosum_plot_id , COUNT(biosum_plot_id) " +
                    " FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName +
                    " GROUP BY biosum_plot_id";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                p_dataMgr.m_strSQL = "INSERT INTO plot_cond_accessible_work_table2 (biosum_plot_id, num_cond, num_cond_not_accessible) " +
                    "SELECT a.biosum_plot_id, a.num_cond,b.cond_not_accessible_count AS num_cond_not_accessible FROM plot_cond_accessible_work_table AS a," +
                    "(SELECT biosum_plot_id, COUNT(*) AS cond_not_accessible_count FROM " +
                    Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " " +
                    "WHERE cond_accessible_yn='N' GROUP BY biosum_plot_id) AS b " +
                    "WHERE a.biosum_plot_id = b.biosum_plot_id";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            }

            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

            if (p_dataMgr.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
            }
            else
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!Error Executing SQL!!\r\n");
                this.m_intError = p_dataMgr.m_intError;
            }

            p_dataMgr = null;
        }
        /// <summary>
        /// populate the haul_costs table and plot table with 
        /// the cheapest route for hauling merch and chip
        /// </summary>
        
        private void getHaulCosts()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//getHaulCosts\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            string strTruckHaulCost;
            string strRailHaulCost;
            string strTransferMerchCost;
            string strTransferChipCost;

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 27;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;

            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Processing Haul Costs");
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", true);
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nDelete all records from haul_costs table\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "-------------------------------------------------------------\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();
            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                // attach scenario results db
                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strSystemResultsDbPathAndFile + "' AS results";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                // attach gis travel times db
                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strPSitePathAndFile + "' AS travel_times";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                /********************************************
                **delete all records in the table
                ********************************************/
                p_dataMgr.m_strSQL = "DELETE FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHaulCostsTableName;
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nUpdate Plot And Haul Cost Tables With Merch And Chip Haul Costs\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "-------------------------------------------------------------\r\n");
                }

                /********************************************
                **get the haul cost per green ton per hour
                ********************************************/
                strTruckHaulCost = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_costs1.RoadHaulCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
                strTruckHaulCost = strTruckHaulCost.Replace(",", "");

                if (strTruckHaulCost.Trim().Length == 1) strTruckHaulCost = "0.00";

                strRailHaulCost = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_costs1.RailHaulCostDollarsPerGreenTonPerMile.Replace("$", "").ToString();
                strRailHaulCost = strRailHaulCost.Replace(",", "");
                if (strRailHaulCost.Trim().Length == 1) strRailHaulCost = "0.00";

                /***********************************************
                 **get the transfer cost per green to per hour
                 ***********************************************/
                strTransferMerchCost = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_costs1.RailMerchTransferCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
                strTransferMerchCost = strTransferMerchCost.Replace(",", "");
                if (strTransferMerchCost.Trim().Length == 1) strTransferMerchCost = "0.00";

                strTransferChipCost = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_costs1.RailChipTransferCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
                strTransferChipCost = strTransferChipCost.Replace(",", "");
                if (strTransferChipCost.Trim().Length == 1) strTransferChipCost = "0.00";

                /*******************************************************************************
			    **zap the haul_costs table
			    *******************************************************************************/
                p_dataMgr.m_strSQL = "DELETE FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHaulCostsTableName;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ndelete records in haul_costs table\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Null The PSite Work Table's Haul Cost Fields");
                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

                p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName +
                    " SET merch_haul_cost_id = NULL," +
                    "merch_haul_psite = NULL, " +
                    "merch_haul_psite_name = NULL, " +
                    "merch_haul_cost_dpgt = NULL," +
                    "chip_haul_cost_id = NULL," +
                    "chip_haul_psite = NULL," +
                    "chip_haul_psite_name = NULL," +
                    "chip_haul_cost_dpgt = NULL";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nnull the psite work table's haul cost fields\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                /*****************************************************************
			    **delete any records that may exist in the work tables
			    *****************************************************************/
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Delete Records In Work Tables");

                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

                String[] arrWorkTables = { "all_road_merch_haul_costs_work_table", "all_road_chip_haul_costs_work_table", "merch_rh_to_collector_haul_costs_work_table", "chip_rh_to_collector_haul_costs_work_table",
                "merch_plot_to_rh_to_collector_haul_costs_work_table", "chip_plot_to_rh_to_collector_haul_costs_work_table", "cheapest_road_merch_haul_costs_work_table",
                "cheapest_road_chip_haul_costs_work_table", "cheapest_rail_merch_haul_costs_work_table", "cheapest_rail_chip_haul_costs_work_table", "combine_merch_rail_road_haul_costs_work_table",
                "combine_chip_rail_road_haul_costs_work_table", "cheapest_merch_haul_costs_work_table", "cheapest_chip_haul_costs_work_table"};

                foreach (string table in arrWorkTables)
                {
                    p_dataMgr.m_strSQL = "DELETE FROM " + table;
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                }

                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                //MERCH AND CHIP ROAD PROCESSING SITE HAUL COSTS
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Road Haul Costs For Merchantable Wood Processing Sites");
                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\n--Merchantable wood processing site haul costs--\r\n");

                /***************************************************************
			    **process the merch travel times first
			    ***************************************************************/
                /*****************************************************************
                 **Insert into a table all travel time records where the psite 
                 **has road only or road/rail access and processes 
                 **merch only or merch/chip
                 *****************************************************************/

                // complete_haul_costs_dpgt is calculated using the same calculation as road_cost_dpgt
                // since transfer_cost_dpgt and rail_cost_dpgt are set to 0
                p_dataMgr.m_strSQL = "INSERT into all_road_merch_haul_costs_work_table " +
                    "(biosum_plot_id, railhead_id, psite_id, transfer_cost_dpgt, road_cost_dpgt, rail_cost_dpgt, complete_haul_cost_dpgt, materialcd)" +
                    "SELECT t.biosum_plot_id, 0 AS railhead_id," +
                    "s.psite_id, 0 AS transfer_cost_dpgt," +
                    "(" + strTruckHaulCost.Trim() + " * t.one_way_hours) AS road_cost_dpgt," +
                    "0 AS rail_cost_dpgt, (" + strTruckHaulCost.Trim() + " * t.one_way_hours) AS complete_haul_cost_dpgt," +
                    "'M' as materialcd " +
                    "FROM " + m_strTravelTimeTable + " AS t," +
                    m_strPSiteWorkTable + " AS s " + 
                    "WHERE t.psite_id=s.psite_id AND " +
                    "(s.trancd='1' OR s.trancd ='3') AND " +
                    "(s.biocd='3' OR s.biocd='1')  AND t.one_way_hours > 0";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into work table all travel time records where psite has road access and processes merch\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                /**************************************************************************
			    **Find the cheapest plot to merch processing site road route.
			    **The first query (a) returns all rows with biosum_plot_id, road_cost,
			    **and materialcd . The first subquery (b) finds the minimum haul cost
			    **for a plot. The second subquery (c) finds the minimum haul cost for each
			     **plot,psite combination. The where clause returns the desired row.
			     **************************************************************************/
                p_dataMgr.m_strSQL = "INSERT INTO cheapest_road_merch_haul_costs_work_table " +
                    "SELECT null, b.biosum_plot_id,null AS railhead_id,c.psite_id," +
                    "0 AS transfer_cost_dpgt, a.road_cost_dpgt, 0 AS rail_cost_dpgt," +
                    "b.min_cost AS complete_haul_cost_dpgt," +
                    "'M' AS materialcd " +
                    "FROM  all_road_merch_haul_costs_work_table AS a," +
                    "(SELECT biosum_plot_id,MIN(complete_haul_cost_dpgt) AS min_cost " +
                    "FROM all_road_merch_haul_costs_work_table " +
                    "GROUP BY biosum_plot_id) AS b," +
                    "(SELECT biosum_plot_id,  psite_id ," +
                    "MIN(complete_haul_cost_dpgt) AS min_cost2 " +
                    "FROM all_road_merch_haul_costs_work_table " +
                    "GROUP BY biosum_plot_id, psite_id) AS c " +
                    "WHERE  c.biosum_plot_id = b.biosum_plot_id AND " +
                    "a.biosum_plot_id = b.biosum_plot_id AND " +
                    "a.psite_id = c.psite_id AND " +
                    "b.min_cost = c.min_cost2";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into work table. Find the cheapest plot to merch processing site road route.\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Road Haul Costs For Chip Wood Processing Sites");
                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

                /***********************************************************************
			     **Append to a table all travel time records where the psite 
			     **has road only or road/rail access and processes 
			     **chip only or merch/chip
			     ***********************************************************************/
                // complete_haul_costs_dpgt is calculated using the same calculation as road_cost_dpgt
                // since transfer_cost_dpgt and rail_cost_dpgt are set to 0
                p_dataMgr.m_strSQL = "INSERT INTO all_road_chip_haul_costs_work_table " +
                    "SELECT null, t.biosum_plot_id, 0 AS railhead_id," +
                    "s.psite_id, 0 AS transfer_cost_dpgt," +
                    "(" + strTruckHaulCost.Trim() + " * t.one_way_hours) AS road_cost_dpgt," +
                    "0 AS rail_cost_dpgt, " +
                    "(" + strTruckHaulCost.Trim() + " * t.one_way_hours) AS complete_haul_cost_dpgt," +
                    "'C' AS materialcd " +
                    "FROM " + m_strTravelTimeTable + " AS t," +
                    m_strPSiteWorkTable + " AS s " +
                    "WHERE t.psite_id=s.psite_id AND " +
                    "(s.trancd='1' OR s.trancd='3') AND " +
                    "(s.biocd='3' OR s.biocd='2')  AND t.one_way_hours > 0";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into work table all travel time records where psite has road access and processes chips.\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                /******************************************************************
                 **For each plot get the cheapest road route to a psite. 
                 ******************************************************************/
                p_dataMgr.m_strSQL = "INSERT INTO cheapest_road_chip_haul_costs_work_table " +
                    "SELECT null, b.biosum_plot_id, null AS railhead_id, c.psite_id, " +
                    "0 AS transfer_cost_dpgt,a.road_cost_dpgt," +
                    "0 AS rail_cost_dpgt, b.min_cost AS complete_haul_cost_dpgt," +
                    "'C' AS materialcd " +
                    "FROM all_road_chip_haul_costs_work_table AS a," +
                    "(SELECT biosum_plot_id,MIN(complete_haul_cost_dpgt) AS min_cost " +
                    "FROM all_road_chip_haul_costs_work_table " +
                    "GROUP BY biosum_plot_id) AS b," +
                    "(SELECT biosum_plot_id,  psite_id ," +
                    "MIN(complete_haul_cost_dpgt) AS min_cost2 " +
                    "FROM all_road_chip_haul_costs_work_table " +
                    "GROUP BY biosum_plot_id, psite_id) AS c " +
                    "WHERE c.biosum_plot_id=b.biosum_plot_id AND " +
                    "a.biosum_plot_id=b.biosum_plot_id AND " +
                    "a.psite_id=c.psite_id AND b.min_cost=c.min_cost2";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into work table. Find the cheapest plot to chip processing site road route.\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                //MERCH AND CHIP RAIL PROCESSING SITE HAUL COSTS
                /*********************************************************
                 **Append to a table all travel time collector_id (psite)
                 **records where the psite has rail access
                 *********************************************************/
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Rail Haul Costs For Merchantable Wood Processing Sites");
                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

                p_dataMgr.m_strSQL = "INSERT INTO merch_rh_to_collector_haul_costs_work_table " +
                    "SELECT t.psite_id AS railhead_id," +
                    "t.collector_id AS psite_id," +
                    strTransferMerchCost.Trim() + " AS transfer_cost_dpgt," +
                    "0 AS road_cost_dpgt," +
                    "((t.one_way_hours * 45) * " + strRailHaulCost.Trim() + ") AS rail_cost_dpgt," +
                    "0 AS complete_haul_cost_dpgt,  'M' AS materialcd " +
                    "FROM " + m_strTravelTimeTable + " AS t  " +
                    "INNER JOIN  " + m_strPSiteWorkTable + " AS s " +
                    "ON t.collector_id = s.psite_id " +
                    "WHERE  s.trancd='3' And (s.biocd='3' Or s.biocd='1')  AND t.one_way_hours > 0 AND " +
                    "EXISTS (SELECT ss.psite_id " +
                    "FROM " + m_strPSiteWorkTable + " AS ss " +
                    "WHERE t.psite_id=ss.psite_id AND ss.trancd='2' AND (ss.biocd='3' Or ss.biocd='1'))";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into work table all travel time collector_id (psite) records where the psite has rail access and processes merch.\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                /***************************************************************************
                 **Combine records from the travel time table and the 
                 **merch_rh_to_collector_haul_costs_work_table table by matching the 
                 **r.railhead_id with the travel time psite_id. By doing this,
                 **we can calculate the road_cost and get the total cost by summing 
                 **together the plot to railhead road cost, the transfer of material cost,
                 ** and the railhead to collector site rail cost.
                 ***************************************************************************/
                p_dataMgr.m_strSQL = "INSERT INTO merch_plot_to_rh_to_collector_haul_costs_work_table " +
                    "SELECT null, t.biosum_plot_id, r.railhead_id, r.psite_id," +
                    "r.transfer_cost_dpgt," +
                    "(" + strTruckHaulCost.Trim() + " * t.one_way_hours) AS road_cost_dpgt," +
                    "r.rail_cost_dpgt, (r.transfer_cost_dpgt + (" + strTruckHaulCost.Trim() + " * t.one_way_hours) + r.rail_cost_dpgt) AS complete_haul_cost_dpgt," +
                    "'M' AS materialcd " +
                    "FROM  " + m_strTravelTimeTable + " AS t," +
                    "merch_rh_to_collector_haul_costs_work_table AS r " +
                    "WHERE r.railhead_id = t.psite_id  AND t.one_way_hours > 0";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into work table travel time plot records and previous work rail/merch table results\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                p_dataMgr.m_strSQL = "UPDATE merch_plot_to_rh_to_collector_haul_costs_work_table " +
                    "SET complete_haul_cost_dpgt = transfer_cost_dpgt + road_cost_dpgt + rail_cost_dpgt";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nupdate merch by road and rail total haul cost\r\n ");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                /*******************************************************************
                 **Find the cheapest plot to merch processing site rail route.
                 **The first query (a) returns all rows with biosum_plot_id,
                 **railhead_id, transfer_cost, road_cost. The first subquery (b)
                 **finds the minimum haul cost for a plot. The second subquery (c)
                 **finds the minimum haul cost for each plot,psite combination.
                 **The where clause returns the desired row.
                 *******************************************************************/
                p_dataMgr.m_strSQL = "INSERT INTO cheapest_rail_merch_haul_costs_work_table " +
                    "SELECT null, a.biosum_plot_id, a.railhead_id, c.psite_id, " +
                    "a.transfer_cost_dpgt, a.road_cost_dpgt,a.rail_cost_dpgt," +
                    "c.min_cost AS complete_haul_cost_dpgt,'M' as materialcd " +
                    "FROM merch_plot_to_rh_to_collector_haul_costs_work_table AS a," +
                    "(SELECT biosum_plot_id, MIN(complete_haul_cost_dpgt) AS min_cost2 " +
                    "FROM merch_plot_to_rh_to_collector_haul_costs_work_table " +
                    "GROUP BY biosum_plot_id) AS b," +
                    "(SELECT biosum_plot_id, psite_id," +
                    "MIN(complete_haul_cost_dpgt) AS min_cost " +
                    "FROM merch_plot_to_rh_to_collector_haul_costs_work_table " +
                    "GROUP BY biosum_plot_id,psite_id) AS c " +
                    "WHERE  c.biosum_plot_id = b.biosum_plot_id AND " +
                    "a.biosum_plot_id = c.biosum_plot_id AND " +
                    "a.psite_id = c.psite_id AND " +
                    "a.complete_haul_cost_dpgt = c.min_cost AND " +
                    "min_cost2 = min_cost";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert into work table. Find the cheapest plot to merch processing site rail routes\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Rail Haul Costs For Chip Wood Processing Sites");
                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

                /***********************************************************************
                 **Append to a table all travel time collector_id (psite) records
                 **where the psite has rail access and processes chips only or 
                 **both merch/chips
                 ***********************************************************************/
                p_dataMgr.m_strSQL = "INSERT INTO chip_rh_to_collector_haul_costs_work_table " +
                    "SELECT  t.psite_id AS railhead_id," +
                    "t.collector_id AS psite_id," +
                    strTransferChipCost.Trim() + " AS transfer_cost_dpgt," +
                    "0 AS road_cost_dpgt," +
                    "((t.one_way_hours * 45) * " + strRailHaulCost.Trim() + ") AS rail_cost_dpgt," +
                    "0 AS complete_haul_cost_dpgt,  'C' AS materialcd " +
                    "FROM " + m_strTravelTimeTable + " AS t  " +
                    "INNER JOIN  " + m_strPSiteWorkTable + " AS s " +
                    "ON t.collector_id = s.psite_id " +
                    "WHERE s.trancd='3' AND  " +
                    "(s.biocd='3' OR s.biocd='2')  AND t.one_way_hours > 0 AND " +
                    "EXISTS (SELECT ss.psite_id " +
                    "FROM " + m_strPSiteWorkTable + " AS ss " +
                    "WHERE t.psite_id=ss.psite_id AND ss.trancd='2' AND (ss.biocd='3' Or ss.biocd='2'))";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert into work table all travel time collector_id (psite) records where the psite has rail access and processes chips.\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                /*************************************************************************
                 **Combine records from the travel time table and the 
                 **chip_rh_to_collector_haul_costs_work_table by matching the
                 **r.railhead_id with the travel time psite_id. By doing this,
                 **we can calculate the road_cost and get the total cost by summing 
                 **together the plot to railhead road cost, the transfer of material cost,
                 **and the railhead to collector site rail cost.
                 *************************************************************************/
                p_dataMgr.m_strSQL = "INSERT INTO chip_plot_to_rh_to_collector_haul_costs_work_table " +
                    "SELECT null, t.biosum_plot_id, r.railhead_id, r.psite_id," +
                    "r.transfer_cost_dpgt," +
                    "(" + strTruckHaulCost.Trim() + " * t.one_way_hours) AS road_cost_dpgt," +
                    "r.rail_cost_dpgt, " +
                    "(r.transfer_cost_dpgt + road_cost_dpgt + r.rail_cost_dpgt) AS complete_haul_cost_dpgt," +
                    "'C' AS materialcd " +
                    "FROM  " + m_strTravelTimeTable + " AS t," +
                    "chip_rh_to_collector_haul_costs_work_table AS r " +
                    "WHERE r.railhead_id = t.psite_id  AND t.one_way_hours > 0";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into work table travel time plot records and previous rail/chips work table results\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                p_dataMgr.m_strSQL = "UPDATE chip_plot_to_rh_to_collector_haul_costs_work_table " +
                    "SET complete_haul_cost_dpgt = transfer_cost_dpgt + road_cost_dpgt + rail_cost_dpgt";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nupdate chips by road and rail total haul cost\r\n ");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                /************************************************************************
                 **Find the cheapest plot to Chips processing site rail route.
                 **The first query (a) returns all rows with biosum_plot_id, railhead_id,
                 **transfer_cost, road_cost. The first subquery (b) finds the minimum
                 **haul cost for a plot. The second subquery (c) finds the minimum haul
                 **cost for each plot,psite combination. The where clause returns the
                 **desired row.
                 *************************************************************************/
                p_dataMgr.m_strSQL = "INSERT INTO cheapest_rail_chip_haul_costs_work_table " +
                    "SELECT null, a.biosum_plot_id, a.railhead_id, b.psite_id," +
                    "a.transfer_cost_dpgt, a.road_cost_dpgt, a.rail_cost_dpgt," +
                    "b.min_cost AS complete_haul_cost_dpgt,'C' AS materialcd " +
                    "FROM chip_plot_to_rh_to_collector_haul_costs_work_table AS a," +
                    "(SELECT biosum_plot_id, " +
                    "MIN(complete_haul_cost_dpgt) AS min_cost2 " +
                    "FROM chip_plot_to_rh_to_collector_haul_costs_work_table " +
                    "GROUP BY biosum_plot_id) AS c, " +
                    "(SELECT biosum_plot_id, psite_id," +
                    "MIN(complete_haul_cost_dpgt) AS min_cost " +
                    "FROM chip_plot_to_rh_to_collector_haul_costs_work_table " +
                    "GROUP BY biosum_plot_id,psite_id) AS b " +
                    "WHERE  b.biosum_plot_id = c.biosum_plot_id AND " +
                    "a.biosum_plot_id = b.biosum_plot_id AND " +
                    "a.psite_id = b.psite_id AND  " +
                    "a.complete_haul_cost_dpgt = b.min_cost AND " +
                    "min_cost2 = min_cost";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert into work table. Find the cheapest plot to chip processing site rail routes\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n"); ;
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Combine Road And Rail Haul Costs For Merchantable Wood Processing Sites");
                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

                /**************************************************************
                 **combine the cheapest road and rail total cost for each plot
                 **to a merch psite
                 **After the insert there should be two records for each
                 **plot - one with cheapest haul cost by road and another
                 **with cheapest haul cost by rail
                 ***************************************************************/
                p_dataMgr.m_strSQL = "INSERT INTO combine_merch_rail_road_haul_costs_work_table " +
                    "SELECT * FROM cheapest_road_merch_haul_costs_work_table";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert into work table. Cheapest road route to merch psite\r\n ");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n"); ;
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                // Need to specify all fields except the haul_cost_id because there may be duplicate haul_cost_id's  
                // between the rail and road tables. This allows MS Access to auto-assign the haul_cost_id for the
                // inserted records
                p_dataMgr.m_strSQL = "INSERT INTO combine_merch_rail_road_haul_costs_work_table " +
                    "SELECT null, biosum_plot_id, railhead_id, psite_id, transfer_cost_dpgt, road_cost_dpgt, " +
                    "rail_cost_dpgt, complete_haul_cost_dpgt, materialcd FROM cheapest_rail_merch_haul_costs_work_table";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert into work table. Cheapest rail route to merch psite\r\n ");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                /***************************************************
                 **Get the overall cheapest merch route
                 ***************************************************/
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Get Overall Least Expensive Merch Route");
                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
                p_dataMgr.m_strSQL = "INSERT INTO cheapest_merch_haul_costs_work_table " +
                    "SELECT null, a.biosum_plot_id,a.railhead_id,b.psite_id," +
                    "a.transfer_cost_dpgt, a.road_cost_dpgt,  a.rail_cost_dpgt," +
                    "b.min_cost AS complete_haul_cost_dpgt,'M' AS materialcd " +
                    "FROM combine_merch_rail_road_haul_costs_work_table AS a," +
                    "(SELECT biosum_plot_id, MIN(complete_haul_cost_dpgt) AS min_cost2 " +
                    "FROM combine_merch_rail_road_haul_costs_work_table " +
                    "GROUP BY biosum_plot_id) AS c, " +
                    "(SELECT biosum_plot_id, psite_id," +
                    "MIN(complete_haul_cost_dpgt) AS min_cost " +
                    "FROM combine_merch_rail_road_haul_costs_work_table " +
                    "GROUP BY biosum_plot_id,psite_id) AS b " +
                    "WHERE  b.biosum_plot_id = c.biosum_plot_id AND " +
                    "a.biosum_plot_id = b.biosum_plot_id AND " +
                    "a.psite_id = b.psite_id AND " +
                    "a.complete_haul_cost_dpgt = b.min_cost AND " +
                    "min_cost2 = min_cost";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert into work table. Get the overall cheapest merch route\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n"); ;
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Combine Road And Rail Haul Costs For Chip Wood Processing Sites");
                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

                /**************************************************************
                 **combine the cheapest road and rail total cost for each plot
                 **to a chips psite
                 **After the insert there should be two records for each
                 **plot - one with cheapest haul cost by road and another
                 **with cheapest haul cost by rail
                 ***************************************************************/
                p_dataMgr.m_strSQL = "INSERT INTO combine_chip_rail_road_haul_costs_work_table " +
                    "SELECT * FROM cheapest_road_chip_haul_costs_work_table";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert into work table. Cheapest road route to chip psite\r\n ");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                // Need to specify all fields except the haul_cost_id because there may be duplicate haul_cost_id's  
                // between the rail and road tables. This allows MS Access to auto-assign the haul_cost_id for the
                // inserted records
                p_dataMgr.m_strSQL = "INSERT INTO combine_chip_rail_road_haul_costs_work_table " +
                    "SELECT null, biosum_plot_id, railhead_id, psite_id, transfer_cost_dpgt, road_cost_dpgt, " +
                    "rail_cost_dpgt, complete_haul_cost_dpgt, materialcd FROM cheapest_rail_chip_haul_costs_work_table";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert into work table. Cheapest rail route to chip psite\r\n ");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Get Overall Least Expensive Chip Route");
                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

                /******************************************
                 **Get the overall cheapest chip route
                 ******************************************/
                p_dataMgr.m_strSQL = "INSERT INTO cheapest_chip_haul_costs_work_table " +
                    "SELECT null, a.biosum_plot_id, a.railhead_id, b.psite_id, " +
                    "a.transfer_cost_dpgt, a.road_cost_dpgt,  a.rail_cost_dpgt," +
                    "b.min_cost AS complete_haul_cost_dpgt,'C' AS materialcd " +
                    "FROM combine_chip_rail_road_haul_costs_work_table AS a, " +
                    "(SELECT biosum_plot_id,MIN(complete_haul_cost_dpgt) AS min_cost2 " +
                    "FROM combine_chip_rail_road_haul_costs_work_table " +
                    "GROUP BY biosum_plot_id) AS c, " +
                    "(SELECT biosum_plot_id, psite_id," +
                    "MIN(complete_haul_cost_dpgt) AS min_cost " +
                    "FROM combine_chip_rail_road_haul_costs_work_table " +
                    "GROUP BY biosum_plot_id,psite_id) AS b " +
                    "WHERE  b.biosum_plot_id = c.biosum_plot_id AND " +
                    "a.biosum_plot_id = b.biosum_plot_id AND " +
                    "a.psite_id = b.psite_id AND " +
                    "a.complete_haul_cost_dpgt = b.min_cost AND " +
                    "min_cost2 = min_cost";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert into work table. Get the overall cheapest chip route\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                //INSERT INTO HAUL_COSTS TABLE
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Inserting Results Into Haul Costs Table");
                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

                p_dataMgr.m_strSQL = "INSERT INTO haul_costs (biosum_plot_id, railhead_id, psite_id, " +
                    "transfer_cost_dpgt, road_cost_dpgt, rail_cost_dpgt, complete_haul_cost_dpgt, materialcd) " +
                    "SELECT biosum_plot_id, railhead_id, psite_id," +
                    "transfer_cost_dpgt, road_cost_dpgt, rail_cost_dpgt, complete_haul_cost_dpgt, materialcd " +
                    "FROM cheapest_merch_haul_costs_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert into haul_costs table cheapest merch route for each plot\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n"); 
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                p_dataMgr.m_strSQL = "INSERT INTO haul_costs(biosum_plot_id, railhead_id, psite_id, " +
                    "transfer_cost_dpgt, road_cost_dpgt, rail_cost_dpgt, complete_haul_cost_dpgt, materialcd) " +
                    "SELECT biosum_plot_id, railhead_id, psite_id, " +
                    "transfer_cost_dpgt, road_cost_dpgt, rail_cost_dpgt, complete_haul_cost_dpgt, materialcd " +
                    "FROM cheapest_chip_haul_costs_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert into haul_costs table cheapest chip route for each plot\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                //UPDATE PSITE_ACCESSIBLE_WORKTABLE TABLE

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Updating PSite Work Table");
                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

                /**************************************************
                 **Update cheapest merch route fields
                 **************************************************/
                p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS w " +
                    "SET merch_haul_cost_id = h.haul_cost_id, merch_haul_psite = h.psite_id, merch_haul_cost_dpgt = h.complete_haul_cost_dpgt " +
                    "FROM haul_costs AS h WHERE w.biosum_plot_id = h.biosum_plot_id AND h.materialcd = 'M'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate plot merch haul cost fields\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS w " +
                    "SET merch_haul_psite_name = (SELECT p.name FROM processing_site AS p " +
                    "WHERE w.merch_haul_psite = p.psite_id)";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate merch psite name\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n"); ;
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                /*****************************************
                 **Update  cheapest chip routes
                 *****************************************/
                p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS p " +
                    "SET chip_haul_cost_id = (SELECT h.haul_cost_id FROM haul_costs AS h WHERE p.biosum_plot_id = h.biosum_plot_id AND h.materialcd = 'C'), " +
                    "chip_haul_psite = (SELECT h.psite_id FROM haul_costs AS h WHERE p.biosum_plot_id = h.biosum_plot_id AND h.materialcd = 'C'), " +
                    "chip_haul_cost_dpgt = (SELECT h.complete_haul_cost_dpgt FROM haul_costs AS h WHERE p.biosum_plot_id = h.biosum_plot_id AND h.materialcd = 'C')";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate plot chip haul cost fields\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS w " +
                    "SET chip_haul_psite_name = (SELECT p.name FROM processing_site AS p WHERE w.chip_haul_psite = p.psite_id)";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate chip psite name\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n"); ;
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                /******************************************
                 **clean up work tables
                 ******************************************/
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Cleaning Up Haul Cost Work Tables...Stand By");
                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nCleaning up haul cost work tables\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }
                if (this.m_intError == 0)
                {
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
                }
                frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", false);

                frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
            }
        }
        
        /// <summary>
        /// sum the tree_vol_val_by_species_diam_groups table values to tree_vol_val_sum_by_rx_cycle_work
        /// </summary>
        private void sumTreeVolVal()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//sumTreeVolVal\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 3;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;

            /**************************************************************
			 **sum the tree_vol_val_by_species_diam_groups table to
			 **        tree_vol_val_sum_by_rx_cycle_work
			 **************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nSum Tree Volumes and Values By Treatment\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "----------------------------------------\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                // create work table
                if (p_dataMgr.TableExist(workConn, m_strTreeVolValSumTable))
                {
                    p_dataMgr.m_strSQL = "DROP TABLE " + m_strTreeVolValSumTable;
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                }

                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteTreeVolValSumTable(p_dataMgr, workConn, m_strTreeVolValSumTable);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    this.m_intError = p_dataMgr.m_intError;
                    return;
                }

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strProcessorResultsPathAndFile + "' AS processor_results";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "INSERT INTO " + m_strTreeVolValSumTable +
                    " (biosum_cond_id,rxpackage,rx,rxcycle,chip_vol_cf," +
                    "chip_wt_gt,chip_val_dpa,merch_vol_cf," +
                    "merch_wt_gt,merch_val_dpa,place_holder) " +
                    "SELECT s.biosum_cond_id, " +
                    "s.rxpackage,s.rx,s.rxcycle," +
                    "SUM(s.chip_vol_cf) AS chip_vol_cf," +
                    "SUM(s.chip_wt_gt) AS chip_wt_gt," +
                    "SUM(s.chip_val_dpa) AS chip_val_dpa," +
                    "SUM(s.merch_vol_cf) AS merch_vol_cf," +
                    "SUM(s.merch_wt_gt) AS merch_wt_gt," +
                    "SUM(s.merch_val_dpa) AS merch_val_dpa, " +
                    "s.place_holder " +
                    "FROM " + m_strTreeVolValBySpcDiamGroupsTable.Trim() + " AS s " + 
                    "GROUP BY biosum_cond_id,rxpackage,rx,rxcycle, place_holder " +
                    "ORDER BY biosum_cond_id,rxpackage,rx,rxcycle, place_holder";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into tree_vol_val_sum_by_rx_cycle_work table tree volume and value sums\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");

                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    this.m_intError = p_dataMgr.m_intError;
                    return;
                }

                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                if (this.m_intError == 0)
                {
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
                }
            }

            p_dataMgr = null;
        }

        /// <summary>
        /// load the validcombos table with biosum_cond_id,rxpackage,rx and rxcycle values
        /// that exist in the user defined plot filters, condition, ffe, travel times, and
        /// harvest cost, and tree volume/value tables
        /// </summary>
        /// 
        private void validcombos()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//validcombos\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            string strRxList = "";
            string strGrpCd = "";
            int x = 0;
            string strTable = "";
            int intListViewIndex = -1;
            System.Windows.Forms.CheckBox oCheckBox = null;

            intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                         ReferenceUserControlScenarioRun.listViewEx1, "Apply User Defined Filters And Get Valid Stand Combinations");

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 21;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);

            DataMgr p_dataMgr = new DataMgr();

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\" + Tables.Audit.DefaultCondAuditTableSqliteDbFile + "' AS audit";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strPlotPathAndFile + "' AS master";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strPSitePathAndFile + "' AS gis";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strHarvestMethodsPathAndFile + "' AS ref";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile + "' AS calc_prepost_fvsout";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutPrePostDbFile + "' AS prepost_fvsout";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strFVSPrePostValidComboDbPathAndFile + "' AS valid_combo";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strSystemResultsDbPathAndFile + "' AS results";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strProcessorResultsPathAndFile + "' AS processor_results";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                /*****************************************************************
			     **delete audit tables
			     *****************************************************************/
                p_dataMgr.m_strSQL = "DELETE FROM " + Tables.Audit.DefaultCondAuditTableName;
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (p_dataMgr.m_intError != 0)
                {
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    this.m_intError = p_dataMgr.m_intError;
                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                p_dataMgr.m_strSQL = "DELETE FROM " + Tables.Audit.DefaultCondRxAuditTableName;
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (p_dataMgr.m_intError != 0)
                {
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    this.m_intError = p_dataMgr.m_intError;
                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                /**************************
			     **get the treatment list
			     **************************/
                p_dataMgr.m_strSQL = "SELECT rx FROM " + m_strRxTable;
                p_dataMgr.SqlQueryReader(workConn, p_dataMgr.m_strSQL);
                if (!p_dataMgr.m_DataReader.HasRows)
                {
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    this.m_intError = -1;
                    MessageBox.Show("No Treatments Found In The Treatment Table");
                    return;
                }
                while (p_dataMgr.m_DataReader.Read())
                {
                    strRxList += p_dataMgr.m_DataReader["rx"].ToString().Trim();
                }
                p_dataMgr.m_DataReader.Close();

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute User Defined Plot SQL And Insert Resulting Records Into Table userdefinedplotfilter_work\r\n");
                p_dataMgr.m_strSQL = "INSERT INTO userdefinedplotfilter_work " + m_strUserDefinedPlotSQL;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--Execute User Defined Cond SQL And Insert Resulting Records Into Table userdefinedcondfilter_work--\r\n");
                p_dataMgr.m_strSQL = "INSERT INTO userdefinedcondfilter_work " + m_strUserDefinedCondSQL;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--Execute rule definition filters for the condition table. The filters include ownership and condition accessible--\r\n");
                p_dataMgr.m_strSQL = "INSERT INTO ruledefinitionscondfilter SELECT * FROM " + m_strCondTable + " AS c " +
                    "WHERE c.owngrpcd IN (";

                //usfs ownnership
                if (ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_owner_groups1.chkOwnGrp10.Checked == true)
                {
                    strGrpCd = "10,1";
                }
                //other federal ownership
                if (ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_owner_groups1.chkOwnGrp20.Checked == true)
                {
                    if (strGrpCd.Trim().Length == 0)
                    {
                        strGrpCd = "20,2";
                    }
                    else
                    {
                        strGrpCd += ",20,2";
                    }
                }
                //state and local govt ownership
                if (ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_owner_groups1.chkOwnGrp30.Checked == true)
                {
                    if (strGrpCd.Trim().Length == 0)
                    {
                        strGrpCd = "30,3";
                    }
                    else
                    {
                        strGrpCd += ",30,3";
                    }
                }
                //private ownership
                if (ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_owner_groups1.chkOwnGrp40.Checked == true)
                {
                    if (strGrpCd.Trim().Length == 0)
                    {
                        strGrpCd = "40,4";
                    }
                    else
                    {
                        strGrpCd += ",40,4";
                    }
                }
                p_dataMgr.m_strSQL += strGrpCd + ")";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--Execute SQL that deletes from the condition rule definitions table (ruledefinitionscondfilter) those biosum_cond_id that are not found in the user defined condition SQL filter table (userdefinedcondfilter_work)--\r\n");

                p_dataMgr.m_strSQL = "DELETE FROM ruledefinitionscondfilter AS a " +
                        "WHERE NOT EXISTS " +
                        "(SELECT b.biosum_cond_id " +
                        "FROM userdefinedcondfilter_work AS b " +
                        "WHERE a.biosum_cond_id = b.biosum_cond_id)";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--Execute SQL That Includes  Rule Definitions Defined By The User Into Table ruledefinitionsplotfilter--\r\n");
                p_dataMgr.m_strSQL = "INSERT INTO ruledefinitionsplotfilter " +
                        "SELECT DISTINCT p.* from userdefinedplotfilter_work AS p " +
                        "INNER JOIN ruledefinitionscondfilter AS c ON p.biosum_plot_id = c.biosum_plot_id";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                p_dataMgr.m_strSQL = "DELETE FROM validcombos";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                /**********************************************************************
			     **create valid combiniations of biosum_cond_id and treatment and 
			     **also user defined filters by owngrpcd
			     **********************************************************************/
                //insert all the possible valid POST plot+rxpackage+rxcycle records
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--insert all possible valid post plot+rxpackage+rxcycle records--\r\n");

                //cycle1
                p_dataMgr.m_strSQL = "INSERT INTO validcombos_fvspost " +
                           "(biosum_cond_id, rxpackage, rx, rxcycle) " +
                           "SELECT a.biosum_cond_id, b.rxpackage, b.simyear1_rx AS rx, '1' AS rxcycle " +
                           "FROM " + m_strCondTable + " AS a, " + m_strRxPackageTable + " AS b " +
                           "WHERE b.simyear1_rx IS NOT NULL AND LENGTH(TRIM(b.simyear1_rx)) > 0 AND b.simyear1_rx<>'000'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //cycle2
                p_dataMgr.m_strSQL = "INSERT INTO validcombos_fvspost " +
                           "(biosum_cond_id, rxpackage, rx, rxcycle) " +
                           "SELECT a.biosum_cond_id, b.rxpackage, b.simyear2_rx AS rx, '2' AS rxcycle " +
                           "FROM " + m_strCondTable + " AS a, " + m_strRxPackageTable + " AS b " +
                           "WHERE b.simyear2_rx IS NOT NULL AND LENGTH(TRIM(b.simyear2_rx)) > 0 AND b.simyear2_rx<>'000'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //cycle3
                p_dataMgr.m_strSQL = "INSERT INTO validcombos_fvspost " +
                           "(biosum_cond_id, rxpackage, rx, rxcycle) " +
                           "SELECT a.biosum_cond_id, b.rxpackage, b.simyear3_rx AS rx ,'3' AS rxcycle " +
                           "FROM " + m_strCondTable + " AS a, " + m_strRxPackageTable + " AS b " +
                           "WHERE b.simyear3_rx IS NOT NULL AND LENGTH(TRIM(b.simyear3_rx)) > 0 AND b.simyear3_rx<>'000'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //cycle4
                p_dataMgr.m_strSQL = "INSERT INTO validcombos_fvspost " +
                        "(biosum_cond_id, rxpackage, rx, rxcycle) " +
                        "SELECT a.biosum_cond_id, b.rxpackage, b.simyear4_rx AS rx, '4' AS rxcycle " +
                        "FROM " + m_strCondTable + " AS a, " + m_strRxPackageTable + " AS b " +
                        "WHERE b.simyear4_rx IS NOT NULL AND LENGTH(TRIM(b.simyear4_rx)) > 0 AND b.simyear4_rx<>'000'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //insert all the possible valid pre plot+rxpackage+rxcycle records
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--insert all possible valid PRE plot+rxpackage+rxcycle records--\r\n");

                //cycle1
                p_dataMgr.m_strSQL = "INSERT INTO validcombos_fvspre " +
                        "(biosum_cond_id, rxpackage, rx, rxcycle) " +
                        "SELECT a.biosum_cond_id, b.rxpackage, b.simyear1_rx AS rx, '1' AS rxcycle " +
                        "FROM " + m_strCondTable + " AS a, " + m_strRxPackageTable + " AS b " +
                        "WHERE b.simyear1_rx IS NOT NULL AND LENGTH(TRIM(b.simyear1_rx)) > 0 AND b.simyear1_rx<>'000'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //cycle2
                p_dataMgr.m_strSQL = "INSERT INTO validcombos_fvspre " +
                           "(biosum_cond_id, rxpackage, rx, rxcycle) " +
                           "SELECT a.biosum_cond_id, b.rxpackage, b.simyear2_rx AS rx, '2' AS rxcycle " +
                           "FROM " + m_strCondTable + " AS a, " + m_strRxPackageTable + " AS b " +
                           "WHERE b.simyear2_rx IS NOT NULL AND LENGTH(TRIM(b.simyear2_rx)) > 0 AND b.simyear2_rx<>'000'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //cycle3
                p_dataMgr.m_strSQL = "INSERT INTO validcombos_fvspre " +
                           "(biosum_cond_id, rxpackage, rx, rxcycle) " +
                           "SELECT a.biosum_cond_id, b.rxpackage, b.simyear3_rx AS rx, '3' AS rxcycle " +
                           "FROM " + m_strCondTable + " AS a, " + m_strRxPackageTable + " AS b " +
                           "WHERE b.simyear3_rx IS NOT NULL AND LENGTH(TRIM(b.simyear3_rx)) > 0 AND b.simyear3_rx<>'000'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //cycle4
                p_dataMgr.m_strSQL = "INSERT INTO validcombos_fvspre " +
                        "(biosum_cond_id, rxpackage, rx, rxcycle) " +
                        "SELECT a.biosum_cond_id, b.rxpackage, b.simyear4_rx AS rx, '4' AS rxcycle " +
                        "FROM " + m_strCondTable + " AS a, " + m_strRxPackageTable + " AS b " +
                        "WHERE b.simyear4_rx IS NOT NULL AND LENGTH(TRIM(b.simyear4_rx)) > 0 AND b.simyear4_rx<>'000'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                string strWhere = "";
                for (x = 0; x <= FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES - 1; x++)
                {
                    strTable = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.TableName(
                        ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.m_strPostVarArray[x]);
                    if (strTable.Trim().Length > 0)
                    {
                        p_dataMgr.m_strSQL = "UPDATE validcombos_fvspost AS v " +
                            "SET variable" + Convert.ToString(x + 1).Trim() + "_yn = 'Y' " +
                            "WHERE EXISTS (SELECT biosum_cond_id, rxpackage, rx, rxcycle " +
                            "FROM " + strTable + " AS b " +
                            "WHERE v.biosum_cond_id = b.biosum_cond_id AND v.rxpackage = b.rxpackage " +
                            "AND v.rx = b.rx AND v.rxcycle = b.rxcycle)";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                        p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                        p_dataMgr.m_strSQL = "UPDATE validcombos_fvspost SET variable" + Convert.ToString(x + 1).Trim() + "_yn='N' WHERE variable" + Convert.ToString(x + 1).Trim() + "_yn IS NULL";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                        p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                        p_dataMgr.m_strSQL = "UPDATE validcombos_fvspost SET variable" + Convert.ToString(x + 1).Trim() + "_yn='N' WHERE LENGTH(TRIM(variable" + Convert.ToString(x + 1).Trim() + "_yn))=0";
                        p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                        strWhere = strWhere + "b.variable" + Convert.ToString(x + 1).Trim() + "_yn <> 'N' AND ";
                    }

                    strTable = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.TableName(
                        ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.m_strPreVarArray[x]);
                    if (strTable.Trim().Length > 0)
                    {
                        p_dataMgr.m_strSQL = "UPDATE validcombos_fvspre AS v " +
                            "SET variable" + Convert.ToString(x + 1).Trim() + "_yn = 'Y' " +
                            "WHERE EXISTS (SELECT biosum_cond_id, rxpackage, rx, rxcycle " +
                            "FROM " + strTable + " AS b " +
                            "WHERE v.biosum_cond_id = b.biosum_cond_id AND v.rxpackage = b.rxpackage " +
                            "AND v.rx = b.rx AND v.rxcycle = b.rxcycle)";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                        p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                        p_dataMgr.m_strSQL = "UPDATE validcombos_fvspre SET variable" + Convert.ToString(x + 1).Trim() + "_yn='N' WHERE variable" + Convert.ToString(x + 1).Trim() + "_yn IS NULL";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                        p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                        p_dataMgr.m_strSQL = "UPDATE validcombos_fvspre SET variable" + Convert.ToString(x + 1).Trim() + "_yn='N' WHERE LENGTH(TRIM(variable" + Convert.ToString(x + 1).Trim() + "_yn))=0";
                        p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                        strWhere = strWhere + "a.variable" + Convert.ToString(x + 1).Trim() + "_yn <> 'N' AND ";
                    }
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (strWhere.Trim().Length > 0)
                {
                    strWhere = "WHERE a.biosum_cond_id = b.biosum_cond_id AND a.rxpackage=b.rxpackage AND a.rx=b.rx AND a.rxcycle=b.rxcycle AND " + strWhere.Substring(0, strWhere.Length - 5);
                }

                p_dataMgr.m_strSQL = "INSERT INTO validcombos_fvsprepost " +
                    "SELECT a.biosum_cond_id,a.rxpackage,a.rx,a.rxcycle, null AS fvs_variant " +
                    "FROM validcombos_fvspost AS a, validcombos_fvspre AS b " +
                    strWhere;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                // Set the fvs_variant from the plot table
                p_dataMgr.m_strSQL = "UPDATE validcombos_fvsprepost AS v " +
                    "SET fvs_variant = p.fvs_variant " +
                    "FROM " + m_strCondTable + " AS c, " + m_strPlotTable + " AS p " +
                    "WHERE v.biosum_cond_id = c.biosum_cond_id AND c.biosum_plot_id = p.biosum_plot_id";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                // Delete records from the validcombos_fvsprepost if they are in variant/packages on the filter's exclusion list
                string strUnselectedPackages =
                    (string)frmMain.g_oDelegate.GetControlPropertyValue((uc_optimizer_scenario_select_packages)_frmScenario.uc_optimizer_scenario_select_packages1, "RxPackagesNotSelected", false);
                if (!string.IsNullOrEmpty(strUnselectedPackages))
                {
                    string[] strUnselectedPackagesArray = frmMain.g_oUtils.ConvertListToArray(strUnselectedPackages, ",");
                    string strTempWhere = " WHERE FVS_VARIANT || RXPACKAGE IN (";
                    foreach (string strVariantPackage in strUnselectedPackagesArray)
                    {
                        strTempWhere = strTempWhere + "'" + strVariantPackage + "',";
                    }
                    string strTrimWhere = strTempWhere.TrimEnd(',') + ")";
                    p_dataMgr.m_strSQL = "DELETE FROM validcombos_fvsprepost " + strTrimWhere;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                p_dataMgr.m_strSQL = "INSERT INTO validcombos (biosum_cond_id,rxpackage,rx,rxcycle) " +
                    "SELECT c.biosum_cond_id, v.rxpackage, v.rx, v.rxcycle " +
                    "FROM ruledefinitionsplotfilter AS p JOIN ruledefinitionscondfilter AS c " +
                    "ON p.biosum_plot_id = c.biosum_plot_id " +
                    "JOIN validcombos_fvsprepost AS v ON c.biosum_cond_id = v.biosum_cond_id " +
                    "JOIN " + m_strHvstCostsTable + " AS h ON v.rxpackage = h.rxpackage " +
                    "AND v.rx = h.rx AND v.rxcycle = h.rxcycle AND v.biosum_cond_id = h.biosum_cond_id " +
                    "GROUP BY c.biosum_cond_id, v.rxpackage, v.rx, v.rxcycle";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nCreate Valid Combinations\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                p_dataMgr.m_strSQL = "DELETE FROM validcombos AS v " +
                    "WHERE EXISTS (SELECT * from " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS p " +
                    "WHERE v.biosum_cond_id = p.biosum_cond_id AND (p.merch_haul_psite IS NULL))";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nDelete combinations that don't have rows in the travel_times table\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Delete inaccessible plots from Table validcombos\r\n");

                p_dataMgr.m_strSQL = "DELETE FROM validcombos AS v " +
                                "WHERE EXISTS " +
                                 "(SELECT p.* " +
                                 "FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS p " +
                                 "WHERE v.biosum_cond_id =  p.biosum_cond_id AND p.COND_ACCESSIBLE_YN = 'N')";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");

                intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                        ReferenceUserControlScenarioRun.listViewEx1, "Populate Valid Combination Audit Data");

                FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 16;
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);


                oCheckBox = (CheckBox)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(
                               0, intListViewIndex);

                _uc_scenario_run.uc_filesize_monitor4.EndMonitoringFile();

                if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)oCheckBox, "Checked", false) == true)
                {
                    frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Creating Audit Data");
                    frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", true);
                    frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun, "Refresh");

                    _uc_scenario_run.uc_filesize_monitor4.BeginMonitoringFile(
                        ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\" + Tables.Audit.DefaultCondAuditTableSqliteDbFile, 2000000000, "2GB");

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n--cond_audit--\r\n\r\n");
                    p_dataMgr.m_strSQL = "INSERT INTO " + Tables.Audit.DefaultCondAuditTableName + " (biosum_cond_id) " +
                        "SELECT r.biosum_cond_id " +
                        "FROM ruledefinitionscondfilter AS r, userdefinedplotfilter_work AS u " +
                        "WHERE r.biosum_plot_id = u.biosum_plot_id";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Insert All Plots From ruledefinitionscoldfilter table into cond_audit\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    if (m_dataMgr.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        this.m_intError = m_dataMgr.m_intError;
                        return;
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    /************************************************************************
                     **check to see if the plot record exists in the harvest costs table
                     ************************************************************************/
                    p_dataMgr.m_strSQL = "SELECT COUNT(*) FROM " +
                        "(SELECT a.biosum_cond_id FROM " + Tables.Audit.DefaultCondAuditTableName + " AS a " +
                        "JOIN " + m_strHvstCostsTable + " AS b ON a.biosum_cond_id = b.biosum_cond_id " +
                        "LIMIT 1)";

                    if ((int)p_dataMgr.getRecordCount(workConn, p_dataMgr.m_strSQL, "temp") > 0)
                    {
                        p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + 
                                    " SET harvest_costs_yn = 'Y' " +
                                    "WHERE biosum_cond_id " +
                                    "IN (SELECT biosum_cond_id FROM " + m_strHvstCostsTable + ")";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSee if cond record exists in the harvest_costs table\r\n");
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                        p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + 
                        " SET harvest_costs_yn = 'N' " +
                        "WHERE harvest_costs_yn IS NULL OR LENGTH(TRIM(harvest_costs_yn))=0;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet harvest_costs_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    /****************************************************************************
                     **check to see if the plot record exists in the validcombos_fvsprepost table
                     ****************************************************************************/
                    p_dataMgr.m_strSQL = "SELECT COUNT(*) FROM " +
                                "(SELECT a.biosum_cond_id " +
                                "FROM " + Tables.Audit.DefaultCondAuditTableName + " AS a, " +
                                "validcombos_fvsprepost b " +
                                "WHERE a.biosum_cond_id = b.biosum_cond_id LIMIT 1)";

                    if ((int)p_dataMgr.getRecordCount(workConn, p_dataMgr.m_strSQL, "temp") > 0)
                    {
                        p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + 
                                        " SET fvs_prepost_variables_yn = 'Y' " +
                                        "WHERE biosum_cond_id " +
                                        "IN (SELECT biosum_cond_id FROM validcombos_fvsprepost)";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet fvs_prepost_variables_yn=Y if plot record exists in validcombos_fvsprepost table\r\n");
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                        p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + 
                        " SET fvs_prepost_variables_yn = 'N' " +
                        "WHERE fvs_prepost_variables_yn IS NULL OR " +
                        "fvs_prepost_variables_yn <> 'Y'";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet fvs_prepost_variables_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    /********************************************************************************************************
                     **check to see if the plot record exists in the processor tree volume and value tableharvest cost table
                     ********************************************************************************************************/
                    p_dataMgr.m_strSQL = "SELECT COUNT(*) FROM " +
                                "(SELECT a.biosum_cond_id " +
                                "FROM " + Tables.Audit.DefaultCondAuditTableName + " AS a, " +
                                m_strTreeVolValSumTable + " AS b " +
                                "WHERE a.biosum_cond_id = b.biosum_cond_id LIMIT 1)";

                    if ((int)p_dataMgr.getRecordCount(workConn, p_dataMgr.m_strSQL, "temp") > 0)
                    {
                        p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + 
                                        " SET processor_tree_vol_val_yn = 'Y' " +
                                        "WHERE biosum_cond_id " +
                                        "IN (SELECT biosum_cond_id FROM " + m_strTreeVolValSumTable + ");";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet processor_tree_vol_val_yn=Y if plot record exists in " + m_strTreeVolValSumTable + " table\r\n");
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                        p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    /**********************************************************************
                     **check to see if the plot record exists in the gis travel times table
                     **********************************************************************/
                    p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName +
                        " SET processor_tree_vol_val_yn = 'N' " +
                        "WHERE processor_tree_vol_val_yn IS NULL OR processor_tree_vol_val_yn<>'Y'";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet processor_tree_vol_val_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    p_dataMgr.m_strSQL = "SELECT COUNT(*) FROM " +
                                "(SELECT a.biosum_cond_id " +
                                "FROM " + Tables.Audit.DefaultCondAuditTableName + " AS a," +
                                "ruledefinitionscondfilter AS b," +
                                m_strTravelTimeTable + " AS c " +
                                "WHERE a.biosum_cond_id = b.biosum_cond_id AND c.biosum_plot_id=b.biosum_plot_id LIMIT 1)";

                    if ((int)p_dataMgr.getRecordCount(workConn, p_dataMgr.m_strSQL, "temp") > 0)
                    {
                        p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName +
                                       " SET gis_travel_times_yn = 'Y' " +
                                       "WHERE biosum_cond_id " +
                                       "IN (SELECT biosum_cond_id FROM ruledefinitionscondfilter " +
                                       "WHERE biosum_plot_id " +
                                       "IN (SELECT biosum_plot_id FROM " +
                                       m_strTravelTimeTable + "))";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet gis_travel_times_yn=Y if plot record exists in " + m_strTravelTimeTable + " table\r\n");
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                        p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName +
                                    " SET gis_travel_times_yn = 'N' " +
                                    "WHERE gis_travel_times_yn IS NULL OR  " +
                                    "gis_travel_times_yn <>'Y' ;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet gis_travel_times_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + " AS a " +
                        "SET cond_too_far_steep_yn = p.cond_too_far_steep_yn, psite_merch_yn = " +
                        "CASE WHEN p.merch_haul_psite IS NULL THEN 'N' ELSE 'Y' END, " +
                        "psite_chip_yn = CASE WHEN p.chip_haul_psite IS NULL THEN 'N' ELSE 'Y' END " +
                        "FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS p " +
                        "WHERE a.biosum_cond_id = p.biosum_cond_id";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet accessibility and psite flags\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    //BIOSUM_COND_ID + RX RECORD AUDIT
                    /**********************************************************************************
                    **Insert all the biosum_cond_id + rx combinations into the cond_rx_audit table
                    ***********************************************************************************/
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n--cond_rx_audit--\r\n");

                    //cycle1
                    p_dataMgr.m_strSQL = "INSERT INTO " + Tables.Audit.DefaultCondRxAuditTableName +
                         "(biosum_cond_id,rxpackage,rx,rxcycle)  " +
                         "SELECT a.biosum_cond_id, b.rxpackage,b.rx,b.rxcycle " +
                         "FROM " + Tables.Audit.DefaultCondAuditTableName + " AS a, " +
                        "(SELECT DISTINCT rxpackage,simyear1_rx AS rx,'1' AS rxcycle " +
                         "FROM " + m_strRxPackageTable + " " +
                         "WHERE simyear1_rx IS NOT NULL AND LENGTH(TRIM(simyear1_rx)) > 0 AND simyear1_rx<>'000') AS b ;";  //+ 
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    //cycle2
                    p_dataMgr.m_strSQL = "INSERT INTO " + Tables.Audit.DefaultCondRxAuditTableName +
                         "(biosum_cond_id,rxpackage,rx,rxcycle)  " +
                         "SELECT a.biosum_cond_id, b.rxpackage,b.rx,b.rxcycle " +
                         "FROM " + Tables.Audit.DefaultCondAuditTableName + " AS a, " +
                        "(SELECT DISTINCT rxpackage,simyear2_rx AS rx,'2' AS rxcycle " +
                         "FROM " + m_strRxPackageTable + " " +
                         "WHERE simyear2_rx IS NOT NULL AND LENGTH(TRIM(simyear2_rx)) > 0 AND simyear2_rx<>'000') AS b ;";  //+ 
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    //cycle3
                    p_dataMgr.m_strSQL = "INSERT INTO " + Tables.Audit.DefaultCondRxAuditTableName +
                        "(biosum_cond_id,rxpackage,rx,rxcycle)  " +
                         "SELECT a.biosum_cond_id, b.rxpackage,b.rx,b.rxcycle " +
                         "FROM " + Tables.Audit.DefaultCondAuditTableName + " AS a, " +
                        "(SELECT DISTINCT rxpackage,simyear3_rx AS rx,'3' AS rxcycle " +
                         "FROM " + m_strRxPackageTable + " " +
                         "WHERE simyear3_rx IS NOT NULL AND LENGTH(TRIM(simyear3_rx)) > 0 AND simyear3_rx<>'000') AS b ;";  //+ 
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    //cycle4
                    p_dataMgr.m_strSQL = "INSERT INTO " + Tables.Audit.DefaultCondRxAuditTableName +
                         "(biosum_cond_id,rxpackage,rx,rxcycle)  " +
                         "SELECT a.biosum_cond_id, b.rxpackage,b.rx,b.rxcycle " +
                         "FROM " + Tables.Audit.DefaultCondAuditTableName + " AS a, " +
                        "(SELECT DISTINCT rxpackage,simyear4_rx AS rx,'4' AS rxcycle " +
                         "FROM " + m_strRxPackageTable + " " +
                         "WHERE simyear4_rx IS NOT NULL AND LENGTH(TRIM(simyear4_rx)) > 0 AND simyear4_rx<>'000') AS b ;";  //+ 
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    /*********************************************************************************
                    **check to see if the cond + rx record exists in the fvs prepost variables table
                    *********************************************************************************/
                    p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondRxAuditTableName +
                        " SET fvs_prepost_variables_yn = 'Y' " +
                        "WHERE EXISTS (SELECT biosum_cond_id,rxpackage,rx,rxcycle " +
                        "FROM validcombos_fvsprepost " +
                        "WHERE " + Tables.Audit.DefaultCondRxAuditTableName + ".biosum_cond_id = " +
                        "validcombos_fvsprepost.biosum_cond_id AND " +
                        Tables.Audit.DefaultCondRxAuditTableName + ".rxpackage = validcombos_fvsprepost.rxpackage AND " +
                        Tables.Audit.DefaultCondRxAuditTableName + ".rx = validcombos_fvsprepost.rx AND " +
                        Tables.Audit.DefaultCondRxAuditTableName + ".rxcycle=validcombos_fvsprepost.rxcycle)";
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondRxAuditTableName +
                        " SET fvs_prepost_variables_yn = 'N' " +
                        "WHERE fvs_prepost_variables_yn IS NULL OR LENGTH(TRIM(fvs_prepost_variables_yn))=0";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet fvs_prepost_variables_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    /****************************************************************************
                     **check to see if the plot + rx record exists in the harvest costs table
                     ****************************************************************************/
                    p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondRxAuditTableName + " AS c " +
                    "SET harvest_costs_yn = 'Y' " +
                    "WHERE EXISTS (SELECT biosum_cond_id,rxpackage,rx,rxcycle " +
                    "FROM " + m_strHvstCostsTable + " AS h " +
                    "WHERE c.biosum_cond_id = h.biosum_cond_id AND " +
                    "c.rxpackage = h.rxpackage AND " +
                    "c.rx = h.rx AND " +
                    "c.rxcycle = h.rxcycle)";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet harvest_costs_yn=Y if plot + rx + rxpackage + rxcycle record exists in " + m_strHvstCostsTable + " table\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondRxAuditTableName +
                        " SET harvest_costs_yn = 'N' " +
                        "WHERE harvest_costs_yn IS NULL OR LENGTH(TRIM(harvest_costs_yn))=0 ;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet harvest_costs_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    /*********************************************************************************
                     **check to see if the plot + rx record exists in the processor tree vol val table
                     *********************************************************************************/
                    p_dataMgr.m_strSQL = "SELECT COUNT(*) FROM " +
                                    "(SELECT a.biosum_cond_id " +
                                    "FROM " + Tables.Audit.DefaultCondRxAuditTableName + " AS a," +
                                    m_strTreeVolValSumTable + " AS b " +
                                    "WHERE a.biosum_cond_id = b.biosum_cond_id AND " +
                                    "a.rxpackage = b.rxpackage AND " +
                                    "a.rx=b.rx AND " +
                                    "a.rxcycle=b.rxcycle LIMIT 1)";

                    if ((int)p_dataMgr.getRecordCount(workConn, p_dataMgr.m_strSQL, "temp") > 0)
                    {
                        p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondRxAuditTableName + " AS a " +
                            "SET processor_tree_vol_val_yn = 'Y' " +
                            "WHERE EXISTS (SELECT 1 FROM " + m_strTreeVolValSumTable + " AS b " +
                            "WHERE a.biosum_cond_id = b.biosum_cond_id AND a.rxpackage = b.rxpackage " +
                            "AND a.rx = b.rx AND a.rxcycle = b.rxcycle)";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet processor_tree_vol_val_yn=Y if plot + rx + rxpackage + rxcycle record exists in " + m_strTreeVolValSumTable + " table\r\n");
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                        p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                    p_dataMgr.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondRxAuditTableName +
                        " SET processor_tree_vol_val_yn = 'N' " +
                        "WHERE processor_tree_vol_val_yn IS NULL OR LENGTH(TRIM(processor_tree_vol_val_yn))=0";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet processor_tree_vol_val_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;
                }
            }

            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", false);
            frmMain.g_oDelegate.SetControlPropertyValue((Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
            _uc_scenario_run.uc_filesize_monitor4.EndMonitoringFile();
            _uc_scenario_run.uc_filesize_monitor3.EndMonitoringFile();
            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
            }

            p_dataMgr = null;
        }

        /// <summary>
        /// evaluate the effectiveness of fvs treatment data 
        /// by loading the effective table with 
        /// results from user defined expressions 
        /// </summary>
        private void Effective()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Effective\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();
            int x, y;
            string strPreTable = "";
            string strPreColumn = "";
            string strPostTable = "";
            string strPostColumn = "";
            string[] strEffectiveColumnArray;
            string[] strBetterIsNotNull = new string[uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES];
            string[] strWorseIsNotNull = new string[uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES];
            string[] strEffectiveIsNotNull = new string[uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES];
            string strOverallEffectiveIsNotNull = "";
            string strBetterSql = "";
            string strWorseSql = "";
            string strEffectiveSql = "";
            int intListViewIndex = -1;

            string strVariableNumber = "";
            FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_effective.Variables oFvsVar =
                ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nEffective Treatments\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "---------------------------\r\n");
            }

            intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                     ReferenceUserControlScenarioRun.listViewEx1, "Identify Effective Treatments For Each Stand");

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 8;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);

            using (System.Data.SQLite.SQLiteConnection resultsConn = new System.Data.SQLite.SQLiteConnection(p_dataMgr.GetConnectionString(this.m_strSystemResultsDbPathAndFile)))
            {
                resultsConn.Open();

                //attach FVS prepost weighted database
                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile + "' AS fvs_weighted";
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                //attach FVS prepost
                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutPrePostDbFile + "' AS fvs_out";
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                //get all the column names in the effective table
                string strEffectiveTableName = ReferenceOptimizerScenarioForm.OutputTablePrefix + Tables.OptimizerScenarioResults.DefaultScenarioResultsEffectiveTableSuffix;
                strEffectiveColumnArray = p_dataMgr.getFieldNamesArray(resultsConn, "SELECT * FROM " + strEffectiveTableName);

                /********************************************
			     **delete all records in the effective table
			     ********************************************/
                p_dataMgr.m_strSQL = "DELETE FROM " + strEffectiveTableName;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //insert the valid combos into the effective table
                p_dataMgr.m_strSQL = "INSERT INTO " + strEffectiveTableName + " (biosum_cond_id,rxpackage,rx,rxcycle) SELECT biosum_cond_id,rxpackage,rx,rxcycle FROM validcombos WHERE rxcycle='1'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //insert revenue filter field into the effective table
                if (this.m_oOptimizationVariable.bUseFilter == true)
                {
                    uc_optimizer_scenario_calculated_variables.VariableItem oItem = null;
                    foreach (uc_optimizer_scenario_calculated_variables.VariableItem oNextItem in this.ReferenceOptimizerScenarioForm.m_oWeightedVariableCollection)
                    {
                        if (oNextItem.strVariableName.Equals(this.m_oOptimizationVariable.strRevenueAttribute))
                        {
                            oItem = oNextItem;
                            break;
                        }
                    }
                    if (oItem != null)
                    {
                        if (oItem.strVariableSource.IndexOf(".") > -1)
                        {
                            string[] strDatabase = frmMain.g_oUtils.ConvertListToArray(oItem.strVariableSource, ".");
                            p_dataMgr.m_strSQL = "UPDATE  " + strEffectiveTableName + " AS e " +
                                "SET " + this.m_oOptimizationVariable.strRevenueAttribute + " = CASE WHEN p." + strDatabase[1] + " IS NOT NULL THEN p." + strDatabase[1] + " ELSE 0 END " +
                                "FROM " + strDatabase[0] + " AS p " +
                                "WHERE e.biosum_cond_id = p.biosum_cond_id AND e.rxpackage = p.rxpackage";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                            p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);
                        }
                    }
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //populate the variable table.column name and its value to the effective table
                for (x = 0; x <= uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES - 1; x++)
                {

                    p_dataMgr.m_strSQL = "";
                    strPreTable = "";
                    strPreColumn = "";
                    strPostTable = "";
                    strPostColumn = "";
                    strVariableNumber = Convert.ToString(x + 1).Trim();
                    if (oFvsVar.TableName(oFvsVar.m_strPreVarArray[x].Trim().ToUpper()) != "NOT DEFINED")
                    {

                        strPreTable = oFvsVar.TableName(oFvsVar.m_strPreVarArray[x].Trim());
                        strPreColumn = oFvsVar.ColumnName(oFvsVar.m_strPreVarArray[x].Trim());
                    }
                    if (oFvsVar.m_strPostVarArray[x].Trim().ToUpper() != "NOT DEFINED")
                    {

                        strPostTable = oFvsVar.TableName(oFvsVar.m_strPostVarArray[x].Trim());
                        strPostColumn = oFvsVar.ColumnName(oFvsVar.m_strPostVarArray[x].Trim());
                    }
                    if (strPreTable.Trim().Length > 0)
                    {
                        p_dataMgr.m_strSQL = "pre_variable" + strVariableNumber + "_name='" + strPreTable + "." + strPreColumn + "',";
                        p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + "pre_variable" + strVariableNumber + "_value= (SELECT pre." + strPreColumn + " FROM " + strPreTable + 
                            " AS pre WHERE pre.biosum_cond_id = e.biosum_cond_id AND pre.rxpackage = e.rxpackage AND pre.rx = e.rx AND pre.rxcycle = e.rxcycle)";
                    }
                    else
                    {
                        p_dataMgr.m_strSQL = "pre_variable" + strVariableNumber + "_name=null,";
                        p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + "pre_variable" + strVariableNumber + "_value=null";
                    }

                    if (strPostTable.Trim().Length > 0)
                    {
                        p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + ",post_variable" + strVariableNumber + "_name='" + strPostTable + "." + strPostColumn + "',";
                        p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + "post_variable" + strVariableNumber + "_value= (SELECT post." + strPostColumn + " FROM " + strPostTable +
                            " AS post WHERE post.biosum_cond_id = e.biosum_cond_id AND post.rxpackage = e.rxpackage AND post.rx = e.rx AND post.rxcycle = e.rxcycle)";
                    }
                    else
                    {
                        p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + ",post_variable" + strVariableNumber + "_name=null,";
                        p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + "post_variable" + strVariableNumber + "_value=null";
                    }
                    if (strPreTable.Trim().Length > 0 && strPostTable.Trim().Length > 0)
                    {
                        p_dataMgr.m_strSQL = "UPDATE " + strEffectiveTableName + " AS e " +
                            "SET " + p_dataMgr.m_strSQL;
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                        p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);
                    }
                }
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //populate the change column by subtracting pre value from post value
                p_dataMgr.m_strSQL = "";
                for (x = 0; x <= uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES - 1; x++)
                {
                    strVariableNumber = Convert.ToString(x + 1).Trim();
                    p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + "variable" + strVariableNumber + "_change = CASE WHEN pre_variable" + strVariableNumber + "_value IS NOT NULL AND post_variable" +
                        strVariableNumber + "_value IS NOT NULL THEN post_variable" + strVariableNumber + "_value - pre_variable" + strVariableNumber + "_value ELSE NULL END,";
                }
                p_dataMgr.m_strSQL = p_dataMgr.m_strSQL.Substring(0, p_dataMgr.m_strSQL.Length - 1);

                p_dataMgr.m_strSQL = "UPDATE " + strEffectiveTableName + " SET " + p_dataMgr.m_strSQL;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //see what variables are referenced in the sql expression and make sure they are not null
                strOverallEffectiveIsNotNull = "";
                for (x = 0; x <= uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES - 1; x++)
                {
                    strBetterIsNotNull[x] = "";
                    strWorseIsNotNull[x] = "";
                    strEffectiveIsNotNull[x] = "";

                    for (y = 0; y <= strEffectiveColumnArray.Length - 1; y++)
                    {
                        if (oFvsVar.m_strBetterExpr[x].Trim().Length > 0)
                        {
                            if (oFvsVar.m_strBetterExpr[x].Trim().ToUpper().IndexOf(strEffectiveColumnArray[y].ToUpper(), 0) >= 0)
                            {
                                strBetterIsNotNull[x] = strBetterIsNotNull[x] + strEffectiveColumnArray[y] + " IS NOT NULL AND ";
                            }
                        }
                        if (oFvsVar.m_strWorseExpr[x].Trim().Length > 0)
                        {
                            if (oFvsVar.m_strWorseExpr[x].Trim().ToUpper().IndexOf(strEffectiveColumnArray[y].ToUpper(), 0) >= 0)
                            {
                                strWorseIsNotNull[x] = strWorseIsNotNull[x] + strEffectiveColumnArray[y] + " IS NOT NULL AND ";
                            }
                        }
                        if (oFvsVar.m_strEffectiveExpr[x].Trim().Length > 0)
                        {
                            if (oFvsVar.m_strEffectiveExpr[x].Trim().ToUpper().IndexOf(strEffectiveColumnArray[y].ToUpper(), 0) >= 0)
                            {
                                strEffectiveIsNotNull[x] = strEffectiveIsNotNull[x] + strEffectiveColumnArray[y] + " IS NOT NULL AND ";
                            }
                        }
                    }

                }

                if (oFvsVar.m_strOverallEffectiveExpr.Trim().Length > 0)
                {
                    for (y = 0; y <= strEffectiveColumnArray.Length - 1; y++)
                    {

                        if (oFvsVar.m_strOverallEffectiveExpr.Trim().ToUpper().IndexOf(strEffectiveColumnArray[y].ToUpper(), 0) >= 0)
                        {
                            strOverallEffectiveIsNotNull = strOverallEffectiveIsNotNull + strEffectiveColumnArray[y] + " IS NOT NULL AND ";
                        }

                    }
                }

                //remove the last AND
                for (x = 0; x <= uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES - 1; x++)
                {
                    if (strBetterIsNotNull[x].Trim().Length > 0)
                    {
                        strBetterIsNotNull[x] = strBetterIsNotNull[x].Substring(0, strBetterIsNotNull[x].Length - 5);
                    }
                    if (strWorseIsNotNull[x].Trim().Length > 0)
                    {
                        strWorseIsNotNull[x] = strWorseIsNotNull[x].Substring(0, strWorseIsNotNull[x].Length - 5);
                    }
                    if (strEffectiveIsNotNull[x].Trim().Length > 0)
                    {
                        strEffectiveIsNotNull[x] = strEffectiveIsNotNull[x].Substring(0, strEffectiveIsNotNull[x].Length - 5);
                    }
                }

                if (strOverallEffectiveIsNotNull.Trim().Length > 0)
                {
                    strOverallEffectiveIsNotNull = strOverallEffectiveIsNotNull.Substring(0, strOverallEffectiveIsNotNull.Length - 5);
                }

                //populate the better,worse,effective, and overall effective columns
                p_dataMgr.m_strSQL = "";
                strBetterSql = "";
                strWorseSql = "";
                strEffectiveSql = "";
                for (x = 0; x <= uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES - 1; x++)
                {
                    strVariableNumber = Convert.ToString(x + 1).Trim();
                    if (oFvsVar.m_strBetterExpr[x].Trim().Length > 0)
                    {
                        strBetterSql = strBetterSql + "variable" + strVariableNumber + "_better_yn = CASE WHEN " + strBetterIsNotNull[x].Trim() + " THEN CASE WHEN " + oFvsVar.m_strBetterExpr[x].Trim() + " THEN 'Y' ELSE 'N' END ELSE NULL END,";
                    }
                    if (oFvsVar.m_strWorseExpr[x].Trim().Length > 0)
                    {
                        strWorseSql = strWorseSql + "variable" + strVariableNumber + "_worse_yn = CASE WHEN " + strWorseIsNotNull[x].Trim() + " THEN CASE WHEN " + oFvsVar.m_strWorseExpr[x].Trim() + " THEN 'Y' ELSE 'N' END ELSE NULL END,";
                    }
                    if (oFvsVar.m_strEffectiveExpr[x].Trim().Length > 0)
                    {
                        strEffectiveSql = strEffectiveSql + "variable" + strVariableNumber + "_effective_yn = CASE WHEN " + strEffectiveIsNotNull[x].Trim() + " THEN CASE WHEN " + oFvsVar.m_strEffectiveExpr[x].Trim() + " THEN 'Y' ELSE 'N' END ELSE NULL END,";
                    }
                }

                // Mark effective treatments with a 'Y'
                p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + "overall_effective_yn = CASE WHEN " + strOverallEffectiveIsNotNull.Trim() + " THEN CASE WHEN " + oFvsVar.m_strOverallEffectiveExpr.Trim() + " THEN 'Y' ELSE 'N' END ELSE NULL END,";

                //better
                if (strBetterSql.Trim().Length > 0)
                {
                    strBetterSql = strBetterSql.Substring(0, strBetterSql.Length - 1);
                    strBetterSql = "UPDATE " + strEffectiveTableName + " SET " + strBetterSql;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "--Improvement--\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strBetterSql + "\r\n");
                    p_dataMgr.SqlNonQuery(resultsConn, strBetterSql);

                }

                //worse
                if (strWorseSql.Trim().Length > 0)
                {
                    strWorseSql = strWorseSql.Substring(0, strWorseSql.Length - 1);
                    strWorseSql = "UPDATE " + strEffectiveTableName + " SET " + strWorseSql;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "--Disimprovement--\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strWorseSql + "\r\n");
                    p_dataMgr.SqlNonQuery(resultsConn, strWorseSql);

                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //effective
                strEffectiveSql = strEffectiveSql.Substring(0, strEffectiveSql.Length - 1);
                strEffectiveSql = "UPDATE " + strEffectiveTableName + " SET " + strEffectiveSql;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--Variable Effective--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strEffectiveSql + "\r\n");
                p_dataMgr.SqlNonQuery(resultsConn, strEffectiveSql);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //overall effective
                p_dataMgr.m_strSQL = p_dataMgr.m_strSQL.Substring(0, p_dataMgr.m_strSQL.Length - 1);
                p_dataMgr.m_strSQL = "UPDATE " + strEffectiveTableName + " SET " + p_dataMgr.m_strSQL;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--Overall Effective--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug) frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (Convert.ToInt32(p_dataMgr.getRecordCount(resultsConn, "SELECT COUNT(*) FROM " + strEffectiveTableName + " WHERE overall_effective_yn='Y'", "temp")) == 0)
                {

                    MessageBox.Show("No overall effective treatments found. Processing is cancelled");
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Cancelled!!");
                    ReferenceUserControlScenarioRun.m_bUserCancel = true;
                    return;

                }
            }

            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
            }
            p_dataMgr = null;
        }


        private void Optimization()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Optimization\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();
            string strPreTable = "";
            string strPreColumn = "";
            string strPostTable = "";
            string strPostColumn = "";


            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nOptimization\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "--------------------\r\n");
            }

            intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                    ReferenceUserControlScenarioRun.listViewEx1, "Optimize the Effective Treatments For Each Stand");

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 3;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);


            using (System.Data.SQLite.SQLiteConnection resultsConn = new System.Data.SQLite.SQLiteConnection(p_dataMgr.GetConnectionString(this.m_strSystemResultsDbPathAndFile)))
            {
                resultsConn.Open();

                /********************************************
                 **delete all records in the optimization table
                 ********************************************/
                string strOptimizationTableName = ReferenceOptimizerScenarioForm.OutputTablePrefix +
                    Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix;

                p_dataMgr.m_strSQL = "DELETE FROM " + strOptimizationTableName;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //insert the valid combos into the optimization table
                p_dataMgr.m_strSQL = "INSERT INTO " + strOptimizationTableName + " (biosum_cond_id,rxpackage,rx,rxcycle,affordable_YN";
                if (this.m_oOptimizationVariable.bUseFilter == true)
                    p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + "," + this.m_oOptimizationVariable.strRevenueAttribute;
                p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + ") SELECT biosum_cond_id,rxpackage,rx,rxcycle,'Y' ";
                if (this.m_oOptimizationVariable.bUseFilter == true)
                    p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + "," + this.m_oOptimizationVariable.strRevenueAttribute;
                p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + " FROM " + ReferenceOptimizerScenarioForm.OutputTablePrefix +
                    Tables.OptimizerScenarioResults.DefaultScenarioResultsEffectiveTableSuffix +
                    " WHERE overall_effective_yn='Y'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nOptimization Type: " + this.m_oOptimizationVariable.strOptimizedVariable.ToUpper() + "\r\n");

                //populate the variable table.column name and its value to the optimization table
                if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "REVENUE")
                {
                    p_dataMgr.m_strSQL = "UPDATE " + strOptimizationTableName +
                        " SET pre_variable_name = 'p.max_nr_dpa', post_variable_name = 'p.max_nr_dpa', " +
                        "pre_variable_value = CASE WHEN p.max_nr_dpa IS NOT NULL THEN p.max_nr_dpa ELSE 0 END, " +
                        "post_variable_value = CASE WHEN p.max_nr_dpa IS NOT NULL THEN p.max_nr_dpa ELSE 0 END, " +
                        "change_value = 0 " +
                        "FROM " + strOptimizationTableName + " AS e LEFT JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName + " AS p " +
                        "ON e.biosum_cond_id = p.biosum_cond_id AND e.rxpackage = p.rxpackage AND e.rx = p.rx AND e.rxcycle AND p.rxcycle";

                }
                else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "MERCHANTABLE VOLUME")
                {
                    p_dataMgr.m_strSQL = "UPDATE " + strOptimizationTableName +
                        " SET pre_variable_name = 'p.merch_vol_cf', post_variable_name = 'p.merch_vol_cf', " +
                        "pre_variable_value = CASE WHEN p.merch_vol_cf IS NOT NULL THEN p.merch_vol_cf ELSE 0 END, " +
                        "post_variable_value = CASE WHEN p.merch_vol_cf IS NOT NULL THEN p.merch_vol_cf ELSE 0 END, " +
                        "change_value = 0 " +
                        "FROM " + strOptimizationTableName + " AS e LEFT JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName + " AS p " +
                        "ON e.biosum_cond_id = p.biosum_cond_id AND e.rxpackage = p.rxpackage AND e.rx = p.rx AND e.rxcycle AND p.rxcycle";
                }
                else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "ECONOMIC ATTRIBUTE")
                {
                    p_dataMgr.m_strSQL = getEconomicOptimizationSql();

                }
                else
                {

                    string[] strCol = frmMain.g_oUtils.ConvertListToArray(this.m_oOptimizationVariable.strFVSVariableName, ".");
                    strPreTable = "PRE_" + strCol[0].Trim();
                    strPreColumn = strCol[1].Trim();
                    strPostTable = "POST_" + strCol[0].Trim();
                    strPostColumn = strCol[1].Trim();

                    p_dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutPrePostDbFile + "' AS prepost_fvsout";
                    p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);
                    p_dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile + "' AS fvs_weighted";
                    p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);
                    p_dataMgr.m_strSQL = "UPDATE " + strOptimizationTableName + " AS e " +
                        "SET change_value = CASE WHEN a." + strPostColumn + " IS NOT NULL AND b." + strPreColumn + " IS NOT NULL " +
                        "THEN a." + strPostColumn + " - b." + strPreColumn +
                        " WHEN a." + strPostColumn + " IS NOT NULL THEN a." + strPostColumn +
                        " WHEN b." + strPreColumn + " IS NOT NULL THEN 0 - b." + strPreColumn + " ELSE 0 END, " +
                        "pre_variable_name = 'Pre_" + this.m_oOptimizationVariable.strFVSVariableName.Trim() + "'," +
                        "post_variable_name = 'Post_" + this.m_oOptimizationVariable.strFVSVariableName.Trim() + "'," +
                        "pre_variable_value = CASE WHEN b." + strPreColumn + " IS NOT NULL THEN b." + strPreColumn + " ELSE 0 END, " +
                        "post_variable_value = CASE WHEN a." + strPostColumn + " IS NOT NULL THEN a." + strPostColumn + " ELSE 0 END " +
                        "FROM " + strPostTable + " AS a JOIN " + strPreTable + " AS b " +
                        "ON a.biosum_cond_id = b.biosum_cond_id AND a.rxpackage = b.rxpackage AND a.rx = b.rx AND a.rxcycle = b.rxcycle " +
                        "WHERE e.biosum_cond_id = a.biosum_cond_id AND e.rxpackage = a.rxpackage AND e.rx = a.rx AND e.rxcycle = a.rxcycle";
                }
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                // Update affordable flag for revenue filter
                if (this.m_oOptimizationVariable.bUseFilter == true)
                {
                    p_dataMgr.m_strSQL = "UPDATE " + strOptimizationTableName +
                        " SET affordable_YN = CASE WHEN " + this.m_oOptimizationVariable.strRevenueAttribute + " " + this.m_oOptimizationVariable.strFilterOperator + " " + this.m_oOptimizationVariable.dblFilterValue +
                        " THEN 'Y' ELSE 'N' END";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);
                }
            }
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");

            p_dataMgr = null;
        }

        private string getEconomicOptimizationSqlAccess()
        {
            string strSql = "";
            string strOptimizationTableName = ReferenceOptimizerScenarioForm.OutputTablePrefix +
                Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix;
            string[] strCol = frmMain.g_oUtils.ConvertListToArray(this.m_oOptimizationVariable.strFVSVariableName, "_");
            uc_optimizer_scenario_calculated_variables.VariableItem oItem = null;
            foreach (uc_optimizer_scenario_calculated_variables.VariableItem oNextItem in this.ReferenceOptimizerScenarioForm.m_oWeightedVariableCollection)
            {
                if (oNextItem.strVariableName.Equals(this.m_oOptimizationVariable.strFVSVariableName))
                {
                    oItem = oNextItem;
                    break;
                }
            }
            if (strCol.Length > 1)
            {
                // This is a default economic variable; They always end in _1
                if (strCol[strCol.Length - 1] == "1")
                {
                    // We are storing the table and field name in the database for most variables
                    if (oItem.strVariableSource.IndexOf(".") > -1)
                    {
                        string[] strDatabase = frmMain.g_oUtils.ConvertListToArray(oItem.strVariableSource, ".");
                        strSql = "UPDATE " + strOptimizationTableName + " e " +
                                 "INNER JOIN " + strDatabase[0] + " p " +
                                 "ON e.biosum_cond_id=p.biosum_cond_id AND " +
                                 "e.rxpackage=p.rxpackage " +
                                 "SET e.pre_variable_name = '" + oItem.strVariableName + "'," +
                                 "e.post_variable_name = '" + oItem.strVariableName + "'," +
                                 "e.pre_variable_value = IIF(p." + strDatabase[1] + " IS NOT NULL,p." + strDatabase[1] + ",0)," +
                                 "e.post_variable_value = IIF(p." + strDatabase[1] + " IS NOT NULL,p." + strDatabase[1] + ",0)," +
                                 "e.change_value = 0";
                    }
                    // We specify a calculation for the total volume
                    else if (oItem.strVariableName.Equals("total_volume_1"))
                    {
                        strSql = "UPDATE " + strOptimizationTableName + " e " +
                             "INNER JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName + " p " +
                             "ON e.biosum_cond_id=p.biosum_cond_id AND " +
                             "e.rxpackage=p.rxpackage " +
                             "SET e.pre_variable_name = '" + oItem.strVariableName + "'," +
                             "e.post_variable_name = '" + oItem.strVariableName + "'," +
                             "e.pre_variable_value = IIF(p.chip_vol_cf_utilized + p.merch_vol_cf IS NOT NULL,p.chip_vol_cf_utilized + p.merch_vol_cf,0)," +
                             "e.post_variable_value = IIF(p.chip_vol_cf_utilized + p.merch_vol_cf IS NOT NULL,p.chip_vol_cf_utilized + p.merch_vol_cf,0)," +
                             "e.change_value = 0";
                    }
                    else if (oItem.strVariableName.Equals("treatment_haul_costs_1"))
                    {
                        strSql = "UPDATE " + strOptimizationTableName + " e " +
                             "INNER JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName + " p " +
                             "ON e.biosum_cond_id=p.biosum_cond_id AND " +
                             "e.rxpackage=p.rxpackage " +
                             "SET e.pre_variable_name = '" + oItem.strVariableName + "'," +
                             "e.post_variable_name = '" + oItem.strVariableName + "'," +
                             "e.pre_variable_value = HARVEST_ONSITE_COST_DPA + MERCH_HAUL_COST_DPA + IIF (MERCH_CHIP_NR_DPA < MAX_NR_DPA,0, CHIP_HAUL_COST_DPA_utilized)," +
                             "e.post_variable_value = HARVEST_ONSITE_COST_DPA + MERCH_HAUL_COST_DPA + IIF (MERCH_CHIP_NR_DPA < MAX_NR_DPA,0, CHIP_HAUL_COST_DPA_utilized)," +
                             "e.change_value = 0";
                    }
                }
                // This is a custom-weighted economic variable
                else
                {
                    if (oItem.strVariableSource.IndexOf(".") > -1)
                    {
                        string[] strDatabase = frmMain.g_oUtils.ConvertListToArray(oItem.strVariableSource, ".");
                        strSql = "UPDATE " + strOptimizationTableName + " e " +
                                 "INNER JOIN " + strDatabase[0] + " p " +
                                 "ON e.biosum_cond_id=p.biosum_cond_id AND " +
                                 "e.rxpackage=p.rxpackage " +
                                 "SET e.pre_variable_name = '" + oItem.strVariableName + "'," +
                                 "e.post_variable_name = '" + oItem.strVariableName + "'," +
                                 "e.pre_variable_value = IIF(p." + strDatabase[1] + " IS NOT NULL,p." + strDatabase[1] + ",0)," +
                                 "e.post_variable_value = IIF(p." + strDatabase[1] + " IS NOT NULL,p." + strDatabase[1] + ",0)," +
                                 "e.change_value = 0";
                    }
                }
            }
            return strSql;
        }
        private string getEconomicOptimizationSql()
        {
            string strSql = "";
            string strOptimizationTableName = ReferenceOptimizerScenarioForm.OutputTablePrefix +
                Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix;
            string[] strCol = frmMain.g_oUtils.ConvertListToArray(this.m_oOptimizationVariable.strFVSVariableName, "_");
            uc_optimizer_scenario_calculated_variables.VariableItem oItem = null;

            foreach (uc_optimizer_scenario_calculated_variables.VariableItem oNextItem in this.ReferenceOptimizerScenarioForm.m_oWeightedVariableCollection)
            {
                if (oNextItem.strVariableName.Equals(this.m_oOptimizationVariable.strFVSVariableName))
                {
                    oItem = oNextItem;
                    break;
                }
            }

            if (strCol.Length > 1)
            {
                // This is a default economic variable; They always end in _1
                if (strCol[strCol.Length - 1] == "1")
                {
                    // We are storing the table and field name in the database for most variables
                    if (oItem.strVariableSource.IndexOf(".") > -1)
                    {
                        string[] strDatabase = frmMain.g_oUtils.ConvertListToArray(oItem.strVariableSource, ".");
                        strSql = "UPDATE " + strOptimizationTableName + " AS e " +
                            "SET pre_variable_name = '" + oItem.strVariableName + "', " +
                            "post_variable_name = '" + oItem.strVariableName + "', " +
                            "pre_variable_value = CASE WHEN p." + strDatabase[1] + " IS NOT NULL THEN p." + strDatabase[1] + " ELSE 0 END, " +
                            "post_variable_value = CASE WHEN p." + strDatabase[1] + " IS NOT NULL THEN p." + strDatabase[1] + " ELSE 0 END, " +
                            "change_value = 0 " +
                            "FROM " + strDatabase[0] + " AS p WHERE e.biosum_cond_id = p.biosum_cond_id AND p.rxpackage = e.rxpackage";
                    }
                    // We specify a calculation for the total volume
                    else if (oItem.strVariableName.Equals("total_volume_1"))
                    {
                        strSql = "UPDATE " + strOptimizationTableName + " AS e " +
                            "SET pre_variable_name = '" + oItem.strVariableName + "', " +
                            "post_variable_name = '" + oItem.strVariableName + "', " +
                            "pre_variable_value = CASE WHEN p.chip_vol_cf_utilized + p.merch_vol_cf IS NOT NULL THEN p.chip_vol_cf_utilized + p.merch_vol_cf ELSE 0 END, " +
                            "post_variable_value = CASE WHEN p.chip_vol_cf_utilized + p.merch_vol_cf IS NOT NULL THEN p.chip_vol_cf_utilized + p.merch_vol_cf ELSE 0 END, " +
                            "change_value = 0 " +
                            "FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName + " AS p WHERE e.biosum_cond_id = p.biosum_cond_id AND e.rxpackage = p.rxpackage";
                    }
                    else if (oItem.strVariableName.Equals("treatment_haul_costs_1"))
                    {
                        strSql = "UPDATE " + strOptimizationTableName + " AS e " +
                           "SET pre_variable_name = '" + oItem.strVariableName + "', " +
                           "post_variable_name = '" + oItem.strVariableName + "', " +
                           "pre_variable_name = HARVEST_ONSITE_COST_DPA + MERCH_HAUL_COST_DPA + CASE WHEN MERCH_CHIP_NR_DPA < MAX_NR_DPA THEN 0 ELSE CHIP_HAUL_COST_DPA_utilized, " +
                           "post_variable_name = HARVEST_ONSITE_COST_DPA + MERCH_HAUL_COST_DPA + CASE WHEN MERCH_CHIP_NR_DPA < MAX_NR_DPA THEN 0 ELSE CHIP_HAUL_COST_DPA_utilized, " +
                           "change_value = 0 " +
                           "FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName + " AS p WHERE e.biosum_cond_id = p.biosum_cond_id AND e.rxpackage = p.rxpackage";
                    }
                }
                // This is a custom-weighted economic variable
                else
                {
                    if (oItem.strVariableSource.IndexOf(".") > -1)
                    {
                        string[] strDatabase = frmMain.g_oUtils.ConvertListToArray(oItem.strVariableSource, ".");
                        strSql = "UPDATE " + strOptimizationTableName + " AS e " +
                            "SET pre_variable_name = '" + oItem.strVariableName + "', " +
                            "post_variable_name = '" + oItem.strVariableName + "', " +
                            "pre_variable_value = CASE WHEN p." + strDatabase[1] + " IS NOT NULL THEN p." + strDatabase[1] + " ELSE 0 END, " +
                            "post_variable_value = CASE WHEN p." + strDatabase[1] + " IS NOT NULL THEN p." + strDatabase[1] + " ELSE 0 END, " +
                            "change_value = 0 " +
                            "FROM " + strDatabase[0] + " AS p WHERE e.biosum_cond_id = p.biosum_cond_id AND e.rxpackage = p.rxpackage";
                    }
                }
            }
            return strSql;
        }


        private void tiebreaker()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//tiebreaker\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();
            string strPreTable = "";
            string strPreColumn = "";
            string strPostTable = "";
            string strPostColumn = "";
            string[] strArray = null;
            string strVariableNumber = "";

            FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection oTieBreakerCollection =
                ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.m_oSavTieBreakerCollection;
            FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker.TieBreakerItem oItem;


            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nTie Breaker\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "--------------------\r\n");
            }


            intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                   ReferenceUserControlScenarioRun.listViewEx1, "Load Tie Breaker Tables");

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 4;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);

            using (System.Data.SQLite.SQLiteConnection resultsConn = new System.Data.SQLite.SQLiteConnection(p_dataMgr.GetConnectionString(this.m_strSystemResultsDbPathAndFile)))
            {
                resultsConn.Open();

                /********************************************
			     **delete all records in the tie breaker table
			     ********************************************/
                p_dataMgr.m_strSQL = "DELETE FROM tiebreaker";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //insert the valid combos into the tiebreaker table
                string strOptimizationTableName = ReferenceOptimizerScenarioForm.OutputTablePrefix +
                    Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix;
                p_dataMgr.m_strSQL = "INSERT INTO tiebreaker (biosum_cond_id,rxpackage,rx,rxcycle) " +
                              "SELECT biosum_cond_id,rxpackage,rx,rxcycle FROM " + strOptimizationTableName + " WHERE affordable_YN='Y'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //attach FVS prepost weighted database
                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile + "' AS fvs_weighted";
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                //attach FVS out database
                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutPrePostDbFile + "' AS prepost_fvsout";
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                //populate the variable table.column name and its value to the tiebreaker table
                oItem = oTieBreakerCollection.Item(0);  // STAND ATTRIBUTE
                if (oItem.bSelected)
                {
                    strArray = frmMain.g_oUtils.ConvertListToArray(oItem.strFVSVariableName, ".");

                    strPreTable = "PRE_" + strArray[0].Trim();
                    strPreColumn = strArray[1].Trim();
                    strPostTable = "POST_" + strArray[0].Trim();
                    strPostColumn = strArray[1].Trim();
                    strVariableNumber = "1";
                    if (strPreTable.Trim().Length > 0)
                    {
                        p_dataMgr.m_strSQL = "pre_variable" + strVariableNumber + "_name='" + strPreTable + "." + strPreColumn + "',";
                        p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + "pre_variable" + strVariableNumber + "_value=pre." + strPreColumn;
                    }
                    else
                    {
                        p_dataMgr.m_strSQL = "pre_variable" + strVariableNumber + "_name=null,";
                        p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + "pre_variable" + strVariableNumber + "_value=null";
                    }

                    if (strPostTable.Trim().Length > 0)
                    {
                        p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + ",post_variable" + strVariableNumber + "_name='" + strPostTable + "." + strPostColumn + "',";
                        p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + "post_variable" + strVariableNumber + "_value=post." + strPostColumn;
                    }
                    else
                    {
                        p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + ",post_variable" + strVariableNumber + "_name=null,";
                        p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + "post_variable" + strVariableNumber + "_value=null";
                    }
                    if (strPreTable.Trim().Length > 0 && strPostTable.Trim().Length > 0)
                    {
                        p_dataMgr.m_strSQL = "UPDATE tiebreaker AS e " +
                            "SET " + p_dataMgr.m_strSQL +
                            " FROM " + strPostTable + " AS post JOIN " + strPreTable + " AS pre " +
                            "ON post.biosum_cond_id = pre.biosum_cond_id AND post.rxpackage = pre.rxpackage AND post.rx = pre.rx AND post.rxcycle = pre.rxcycle " +
                            "WHERE e.biosum_cond_id = post.biosum_cond_id AND e.rxpackage = post.rxpackage AND e.rx = post.rx AND e.rxcycle = post.rxcycle";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                        p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                    }

                    //populate the change column by subtracting pre value from post value
                    p_dataMgr.m_strSQL = "";
                    p_dataMgr.m_strSQL += "variable" + strVariableNumber + "_change = CASE WHEN pre_variable" + strVariableNumber + "_value IS NOT NULL AND post_variable" + strVariableNumber + "_value IS NOT NULL " +
                        "THEN post_variable" + strVariableNumber + "_value - pre_variable" + strVariableNumber + "_value ELSE NULL END,";
                    p_dataMgr.m_strSQL = p_dataMgr.m_strSQL.Substring(0, p_dataMgr.m_strSQL.Length - 1);

                    p_dataMgr.m_strSQL = "UPDATE tiebreaker SET " + p_dataMgr.m_strSQL;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);
                }

                oItem = oTieBreakerCollection.Item(1);  // ECONOMIC ATTRIBUTE
                if (oItem.bSelected)
                {
                    string[] strCol = frmMain.g_oUtils.ConvertListToArray(oItem.strFVSVariableName, "_");
                    uc_optimizer_scenario_calculated_variables.VariableItem oWeightItem = null;
                    foreach (uc_optimizer_scenario_calculated_variables.VariableItem oNextItem in this.ReferenceOptimizerScenarioForm.m_oWeightedVariableCollection)
                    {
                        if (oNextItem.strVariableName.Equals(oItem.strFVSVariableName))
                        {
                            oWeightItem = oNextItem;
                            break;
                        }
                    }

                    if (strCol.Length > 1)
                    {
                        // This is a default economic variable; They always end in _1
                        if (strCol[strCol.Length - 1] == "1")
                        {
                            // We are storing the table and field name in the database for most variables
                            if (oWeightItem.strVariableSource.IndexOf(".") > -1)
                            {
                                string[] strDatabase = frmMain.g_oUtils.ConvertListToArray(oWeightItem.strVariableSource, ".");
                                p_dataMgr.m_strSQL = "UPDATE tiebreaker AS e " +
                                    "SET pre_variable1_name = '" + oItem.strFVSVariableName + "', " +
                                    "post_variable1_name = '" + oItem.strFVSVariableName + "', " +
                                    "pre_variable1_value = CASE WHEN p." + strDatabase[1] + " IS NOT NULL THEN p." + strDatabase[1] + " ELSE 0 END, " +
                                    "post_variable1_value = CASE WHEN p." + strDatabase[1] + " IS NOT NULL THEN p." + strDatabase[1] + " ELSE 0 END, " +
                                    "variable1_change = 0 " +
                                    "FROM " + strDatabase[0] + " AS p WHERE e.biosum_cond_id = p.biosum_cond_id AND e.rxpackage = p.rxpackage";
                            }
                            // We specify a calculation for the total volume
                            else if (oItem.strFVSVariableName.Equals("total_volume_1"))
                            {
                                p_dataMgr.m_strSQL = "UPDATE tiebreaker AS e " +
                                    "SET pre_variable1_name = '" + oItem.strFVSVariableName + "', " +
                                    "post_variable1_name = '" + oItem.strFVSVariableName + "', " +
                                    "pre_variable1_value = CASE WHEN p.chip_vol_cf_utilized + p.merch_vol_cf IS NOT NULL THEN p.chip_vol_cf_utilized + p.merch_vol_cf ELSE 0 END, " +
                                    "post_variable1_value = CASE WHEN p.chip_vol_cf_utilized + p.merch_vol_cf IS NOT NULL THEN p.chip_vol_cf_utilized + p.merch_vol_cf ELSE 0 END, " +
                                    "variable1_change = 0 " +
                                    "FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName + " AS p " +
                                    "WHERE e.biosum_cond_id = p.biosum_cond_id AND e.rxpackage = p.rxpackage";
                            }
                            else if (oItem.strFVSVariableName.Equals("treatment_haul_costs_1"))
                            {
                                p_dataMgr.m_strSQL = "UPDATE tiebreaker AS e " +
                                    "SET pre_variable1_name = '" + oItem.strFVSVariableName + "', " +
                                    "post_variable1_name = '" + oItem.strFVSVariableName + "', " +
                                    "pre_variable1_value = HARVEST_ONSITE_COST_DPA + MERCH_HAUL_COST_DPA + CASE WHEN MERCH_CHIP_NR_DPA < MAX_NR_DPA THEN 0 ELSE CHIP_HAUL_COST_DPA_utilized END, " +
                                    "post_variable1_value = HARVEST_ONSITE_COST_DPA + MERCH_HAUL_COST_DPA + CASE WHEN MERCH_CHIP_NR_DPA < MAX_NR_DPA THEN 0 ELSE CHIP_HAUL_COST_DPA_utilized END, " +
                                    "variable1_change = 0 " +
                                    "FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName + " AS p " +
                                    "WHERE e.biosum_cond_id = p.biosum_cond_id AND e.rxpackage = p.rxpackage";
                            }
                        }
                        // This is a custom-weighted economic variable
                        // This is the same SQL used for built-in economic variables where the table/field are stored in database (see above)
                        else
                        {
                            string[] strDatabase = frmMain.g_oUtils.ConvertListToArray(oWeightItem.strVariableSource, ".");
                            p_dataMgr.m_strSQL = "UPDATE tiebreaker AS e " +
                                "SET pre_variable1_name = '" + oItem.strFVSVariableName + "', " +
                                "post_variable1_name = '" + oItem.strFVSVariableName + "', " +
                                "pre_variable1_value = CASE WHEN p." + strDatabase[1] + " IS NOT NULL THEN p." + strDatabase[1] + " ELSE 0 END, " +
                                "post_variable1_value = CASE WHEN p." + strDatabase[1] + " IS NOT NULL THEN p." + strDatabase[1] + " ELSE 0 END, " +
                                "variable1_change = 0 " +
                                "FROM " + strDatabase[0] + " AS p WHERE e.biosum_cond_id = p.biosum_cond_id AND e.rxpackage = p.rxpackage";
                        }

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                        p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);
                    }
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                oItem = oTieBreakerCollection.Item(2);  // LAST TIEBREAK RANK
                if (oItem.bSelected)
                {
                    p_dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
                    Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile + "' AS rule_defs";
                    p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);

                    p_dataMgr.m_strSQL = "UPDATE tiebreaker AS a " +
                        "SET last_tiebreak_rank = b.last_tiebreak_rank " +
                        "FROM scenario_last_tiebreak_rank AS b WHERE a.rxpackage = b.rxpackage " +
                        "AND TRIM(UPPER(b.scenario_id)) = '" + ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "'";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);
                }
            }
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");

            p_dataMgr = null;
        }


        /// <summary>
		/// get the wood product yields,
		/// revenue, and costs of an applied
		/// treatment on a plot 
		/// </summary>
        private void econ_by_rx_cycle()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//econ_by_rx_cycle\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nWood Product Yields,Revenue, And Costs Table By Treatment\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "-----------------------------------------------\r\n");
            }
            int intListViewIndex = -1;

            string[] strUpdateSQL = new string[25];

            /**********************************************
			 **complete harvest cost per acre
			 **********************************************/
            intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                        ReferenceUserControlScenarioRun.listViewEx1, "Summarize Wood Product Volume Yields, Costs, And Net Revenue For A Stand And Treatment");

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 5;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);

            /*****************************************************
			 **get the Chips market value per green ton
			 *****************************************************/
            DataMgr p_dataMgr = new DataMgr();

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                string strScenarioDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + strScenarioDb + "' AS scenario";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strSystemResultsDbPathAndFile + "' AS results";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strCondPathAndFile + "' AS master";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strProcessorResultsPathAndFile + "' AS processor_results";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                /********************************************
			     **delete all records in the table
			     ********************************************/
                p_dataMgr.m_strSQL = "DELETE FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert Records\r\n");

                if (p_dataMgr.TableExist(workConn, m_strEconByRxWorkTableName))
                {
                    p_dataMgr.SqlNonQuery(workConn, "DROP TABLE " + m_strEconByRxWorkTableName);
                }

                // Create econ_by_rx_cycle_work_table in temp work tables db
                // Have to parse by column source
                // Think I need to break this up into several queries
                string strSelectSQL = "SELECT vc.biosum_cond_id, vc.rxpackage, vc.rx, vc.rxcycle, " +
                    "tvvs.merch_vol_cf, tvvs.chip_vol_cf, tvvs.chip_wt_gt, tvvs.chip_val_dpa, tvvs.merch_wt_gt, tvvs.merch_val_dpa, hc.complete_cpa AS harvest_onsite_cost_dpa, " +
                    "CASE WHEN psa.merch_haul_cost_dpgt IS NOT NULL THEN " +
                    "CASE WHEN tvvs.rxcycle = '2' THEN psa.merch_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle2 + " ELSE " +
                    "CASE WHEN tvvs.rxcycle = '3' THEN psa.merch_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle3 + " ELSE " +
                    "CASE WHEN tvvs.rxcycle= '4' THEN psa.merch_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle4 + " ELSE " +
                    "psa.merch_haul_cost_dpgt END END END ELSE 0 END AS escalator_merch_haul_cpa_pt, " +
                    "escalator_merch_haul_cpa_pt * tvvs.merch_wt_gt AS merch_haul_cost_dpa, " +
                    "CASE WHEN psa.chip_haul_cost_dpgt IS NOT NULL THEN " +
                    "CASE WHEN tvvs.rxcycle = '2' THEN psa.chip_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle2 + " ELSE " +
                    "CASE WHEN tvvs.rxcycle = '3' THEN psa.chip_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle3 + " ELSE " +
                    "CASE WHEN tvvs.rxcycle = '4' THEN psa.chip_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle4 + " ELSE " +
                    "psa.chip_haul_cost_dpgt END END END ELSE 0 END AS escalator_chip_haul_cpa_pt, " +
                    "escalator_chip_haul_cpa_pt * tvvs.chip_wt_gt AS chip_haul_cost_dpa, " +
                    "merch_val_dpa + chip_val_dpa - harvest_onsite_cost_dpa - (chip_haul_cost_dpa + merch_haul_cost_dpa) AS merch_chip_nr_dpa, " +
                    "merch_val_dpa - harvest_onsite_cost_dpa - merch_haul_cost_dpa AS merch_nr_dpa, " +
                    "CASE WHEN psa.chip_haul_psite IS NULL THEN 'N' ELSE 'Y' END AS usebiomass_yn, " +
                    "0.0 AS max_nr_dpa, c.acres, c.owngrpcd, " +
                    "CASE WHEN usebiomass_yn = 'Y' THEN merch_haul_cost_dpa + chip_haul_cost_dpa ELSE merch_haul_cost_dpa END AS haul_costs_dpa " +
                    "FROM validcombos AS vc, " + m_strCondTable + " AS c, " + m_strHvstCostsTable.Trim() + " AS hc, " +
                    Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS psa, " + m_strTreeVolValSumTable.Trim() + " AS tvvs " +
                    "WHERE vc.biosum_cond_id = c.biosum_cond_id AND c.biosum_cond_id = psa.biosum_cond_id AND " +
                    "vc.biosum_cond_id = tvvs.biosum_cond_id AND vc.rxpackage = tvvs.rxpackage AND vc.rx = tvvs.rx AND " +
                    "vc.rxcycle = tvvs.rxcycle AND vc.biosum_cond_id = hc.biosum_cond_id AND vc.rxpackage = hc.rxpackage AND " +
                    "vc.rx = hc.rx = vc.rxcycle = hc.rx_cycle AND tvvs.place_holder = 'N'";

                string fieldsAndDataTypes = "";

                string field = "";
                string dataType = "";
                string[] columnsFromValidCombos = { "biosum_cond_id", "rxpackage", "rx", "rxcycle" };
                foreach (string column in columnsFromValidCombos)
                {
                    field = "";
                    dataType = "";
                    p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT " + column + " FROM validcombos", ref field, ref dataType);
                    dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                    fieldsAndDataTypes += field + " " + dataType + ", ";
                }

                string[] columnsFromTreeVolValSum = { "merch_vol_cf", "chip_vol_cf", "chip_wt_gt", "chip_val_dpa", "merch_wt_gt", "merch_val_dpa" };
                foreach (string column in columnsFromTreeVolValSum)
                {
                    field = "";
                    dataType = "";
                    p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT " + column + " FROM " + m_strTreeVolValSumTable.Trim(), ref field, ref dataType);
                    dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                    fieldsAndDataTypes += field + " " + dataType + ", ";
                }

                field = "";
                dataType = "";
                p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT complete_cpa FROM " + m_strHvstCostsTable.Trim(), ref field, ref dataType);
                dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                fieldsAndDataTypes += "harvest_onsite_cost_dpa " + dataType + ", ";

                fieldsAndDataTypes += "escalator_merch_haul_cpa_pt DOUBLE, merch_haul_cost_dpa DOUBLE, escalator_chip_haul_cpa_pt DOUBLE," +
                    "chip_haul_cost_dpa DOUBLE, merch_chip_nr_dpa DOUBLE, merch_nr_dpa DOUBLE, usebiomass_yn CHAR(1), max_nr_dpa DOUBLE, ";

                string[] columnsFromCondTable = { "acres", "owngrpcd" };
                foreach (string column in columnsFromCondTable)
                {
                    field = "";
                    dataType = "";
                    p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT " + column + " FROM " + m_strCondTable.Trim(), ref field, ref dataType);
                    dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                    fieldsAndDataTypes += field + " " + dataType + ", ";
                }

                fieldsAndDataTypes += "haul_costs_dpa DOUBLE, ";

                p_dataMgr.m_strSQL = "CREATE TABLE " + m_strEconByRxWorkTableName + " (" + fieldsAndDataTypes + "PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "INSERT INTO " + m_strEconByRxWorkTableName +
                    " (biosum_cond_id, rxpackage, rx, rxcycle, merch_vol_cf, chip_vol_cf, chip_wt_gt, " +
                    "chip_val_dpa, merch_wt_gt, merch_val_dpa, harvest_onsite_cost_dpa, escalator_merch_haul_cpa_pt, " +
                    "escalator_chip_haul_cpa_pt, usebiomass_yn, max_nr_dpa, acres, owngrpcd) " +
                    "SELECT vc.biosum_cond_id, vc.rxpackage, vc.rx, vc.rxcycle, " +
                    "tvvs.merch_vol_cf, tvvs.chip_vol_cf, tvvs.chip_wt_gt, tvvs.chip_val_dpa, tvvs.merch_wt_gt, tvvs.merch_val_dpa, hc.complete_cpa AS harvest_onsite_cost_dpa, " +
                    "CASE WHEN psa.merch_haul_cost_dpgt IS NOT NULL THEN " +
                    "CASE WHEN tvvs.rxcycle = '2' THEN psa.merch_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle2 + " ELSE " +
                    "CASE WHEN tvvs.rxcycle = '3' THEN psa.merch_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle3 + " ELSE " +
                    "CASE WHEN tvvs.rxcycle= '4' THEN psa.merch_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle4 + " ELSE " +
                    "psa.merch_haul_cost_dpgt END END END ELSE 0 END AS escalator_merch_haul_cpa_pt, " +
                    "CASE WHEN psa.chip_haul_cost_dpgt IS NOT NULL THEN " +
                    "CASE WHEN tvvs.rxcycle = '2' THEN psa.chip_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle2 + " ELSE " +
                    "CASE WHEN tvvs.rxcycle = '3' THEN psa.chip_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle3 + " ELSE " +
                    "CASE WHEN tvvs.rxcycle = '4' THEN psa.chip_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle4 + " ELSE " +
                    "psa.chip_haul_cost_dpgt END END END ELSE 0 END AS escalator_chip_haul_cpa_pt, " +
                    "CASE WHEN psa.chip_haul_psite IS NULL THEN 'N' ELSE 'Y' END AS usebiomass_yn, " +
                    "0.0 AS max_nr_dpa, c.acres, c.owngrpcd " +
                    "FROM validcombos AS vc, " + m_strCondTable + " AS c, " + m_strHvstCostsTable.Trim() + " AS hc, " +
                    Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS psa, " + m_strTreeVolValSumTable.Trim() + " AS tvvs " +
                    "WHERE vc.biosum_cond_id = c.biosum_cond_id AND c.biosum_cond_id = psa.biosum_cond_id AND " +
                    "vc.biosum_cond_id = tvvs.biosum_cond_id AND vc.rxpackage = tvvs.rxpackage AND vc.rx = tvvs.rx AND " +
                    "vc.rxcycle = tvvs.rxcycle AND vc.biosum_cond_id = hc.biosum_cond_id AND vc.rxpackage = hc.rxpackage AND " +
                    "vc.rx = hc.rx AND vc.rxcycle = hc.rxcycle AND tvvs.place_holder = 'N'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "UPDATE " + m_strEconByRxWorkTableName +
                    " SET merch_haul_cost_dpa = escalator_merch_haul_cpa_pt * merch_wt_gt, " +
                    "chip_haul_cost_dpa = escalator_chip_haul_cpa_pt * chip_wt_gt";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "UPDATE " + m_strEconByRxWorkTableName +
                    " SET merch_chip_nr_dpa = merch_val_dpa + chip_val_dpa - harvest_onsite_cost_dpa - (chip_haul_cost_dpa + merch_haul_cost_dpa), " +
                    "merch_nr_dpa = merch_val_dpa - harvest_onsite_cost_dpa - merch_haul_cost_dpa, " +
                    "haul_costs_dpa = CASE WHEN usebiomass_yn = 'Y' THEN merch_haul_cost_dpa + chip_haul_cost_dpa ELSE merch_haul_cost_dpa END";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);


                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdates to " + m_strEconByRxWorkTableName + "\r\n");

                // Get the highest value for chips out of the Processor item. They should all be the same, but just in case ...            
                string strChipValue = "-1";
                for (int x = 0; x < this.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Count - 1; x++)
                {
                    ProcessorScenarioItem.TreeSpeciesAndDbhDollarValuesItem oItem =
                      this.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x);
                    if (Convert.ToDouble(strChipValue) < Convert.ToDouble(oItem.ChipsDollarPerCubicFootValue.Trim()))
                    {
                        strChipValue = oItem.ChipsDollarPerCubicFootValue.Trim();
                    }
                }

                // Only use Biomass if chip revenue is higher than the cost of hauling them
                p_dataMgr.m_strSQL = "UPDATE " + m_strEconByRxWorkTableName + " AS w " +
                    "SET usebiomass_yn = CASE WHEN " + strChipValue + " < psa.chip_haul_cost_dpgt THEN 'N' ELSE 'Y' END " +
                    "FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS psa " +
                    "WHERE w.biosum_cond_id = psa.biosum_cond_id AND usebiomass_yn = 'Y' AND rxcycle = '1'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                p_dataMgr.m_strSQL = "UPDATE " + m_strEconByRxWorkTableName + " AS w " +
                    "SET usebiomass_yn = CASE WHEN " + strChipValue + " * " + m_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle2 +
                    " < psa.chip_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle2 +
                    " THEN 'N' ELSE 'Y' END " +
                    "FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS psa " +
                    "WHERE w.biosum_cond_id = psa.biosum_cond_id AND usebiomass_yn = 'Y' AND rxcycle = '2'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                p_dataMgr.m_strSQL = "UPDATE " + m_strEconByRxWorkTableName + " AS w " +
                    "SET usebiomass_yn = CASE WHEN " + strChipValue + " * " + m_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle3 +
                    " < psa.chip_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle3 +
                    " THEN 'N' ELSE 'Y' END " +
                    "FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS psa " +
                    "WHERE w.biosum_cond_id = psa.biosum_cond_id AND usebiomass_yn = 'Y' AND rxcycle = '3'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                p_dataMgr.m_strSQL = "UPDATE " + m_strEconByRxWorkTableName + " AS w " +
                    "SET usebiomass_yn = CASE WHEN " + strChipValue + " * " + m_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle4 +
                    " < psa.chip_haul_cost_dpgt * " + m_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle4 +
                    " THEN 'N' ELSE 'Y' END " +
                    "FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS psa " +
                    "WHERE w.biosum_cond_id = psa.biosum_cond_id AND usebiomass_yn = 'Y' AND rxcycle = '4'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");

                    return;
                }

                // Don't use Biomass if no chip weight
                p_dataMgr.m_strSQL = "UPDATE " + m_strEconByRxWorkTableName +
                    " SET usebiomass_yn = CASE WHEN chip_wt_gt = 0 THEN 'N' ELSE 'Y' END " +
                    "WHERE usebiomass_yn = 'Y'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                // Update fields based on USEBIOMASS_YN
                p_dataMgr.m_strSQL = "UPDATE " + m_strEconByRxWorkTableName +
                    " SET haul_costs_dpa = CASE WHEN usebiomass_yn = 'N' THEN merch_haul_cost_dpa ELSE merch_haul_cost_dpa + chip_haul_cost_dpa END, " +
                    "max_nr_dpa = CASE WHEN usebiomass_yn = 'Y' THEN merch_chip_nr_dpa ELSE merch_nr_dpa END";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                // Create temp table with inactive stands from harvest_costs table
                // This should only contain the inactive stands missing from the econ_by_rx_cycle worktable, but we will
                // do an outer join when appending them to be sure

                string strInactiveStandsWorkTable = "InactiveStandsWorkTable";
                strSelectSQL = "SELECT vc.biosum_cond_id, vc.rxpackage, vc.rx, vc.rxcycle, " +
                            "0 AS merch_vol_cf, 0 as chip_vol_cf, 0 as chip_wt_gt, " +
                            "0 as chip_val_dpa, 0 as merch_wt_gt, 0 AS merch_val_dpa, " +
                            "hc.complete_cpa AS harvest_onsite_cost_dpa, 0 AS escalator_merch_haul_cpa_pt ," +
                            "0 AS merch_haul_cost_dpa, 0 AS escalator_chip_haul_cpa_pt, 0 AS chip_haul_cost_dpa, " +
                            "0 AS merch_chip_nr_dpa, 0 AS merch_nr_dpa, 0 AS max_nr_dpa, " +
                            "c.acres, c.owngrpcd, 0 AS haul_costs_dpa " +
                            "FROM validcombos AS vc, " + m_strCondTable + " AS c, " + m_strHvstCostsTable.Trim() + " AS hc " +
                            "WHERE vc.biosum_cond_id = c.biosum_cond_id AND vc.biosum_cond_id = hc.biosum_cond_id AND " +
                            "vc.rxpackage = hc.rxpackage AND vc.rx = hc.rx AND vc.rxcycle = hc.rxcycle AND " +
                            "hc.harvest_cpa = 0 AND hc.complete_cpa <> 0";

                fieldsAndDataTypes = "";

                foreach (string column in columnsFromValidCombos)
                {
                    field = "";
                    dataType = "";
                    p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT " + column + " FROM validcombos", ref field, ref dataType);
                    dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                    fieldsAndDataTypes += field + " " + dataType + ", ";
                }

                fieldsAndDataTypes += "merch_vol_cf INTEGER, chip_vol_cf INTEGER, chip_wt_gt INTEGER, chip_val_dpa INTEGER, merch_wt_gt INTEGER, merch_val_dpa INTEGER, ";

                field = "";
                dataType = "";
                p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT complete_cpa FROM " + m_strHvstCostsTable.Trim(), ref field, ref dataType);
                dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                fieldsAndDataTypes += "harvest_onsite_cost_dpa " + dataType + ", ";

                fieldsAndDataTypes += "escalator_merch_haul_cpa_pt INTEGER, merch_haul_cost_dpa INTEGER, escalator_chip_haul_cpa_pt INTEGER, chip_haul_cost_dpa INTEGER, " +
                    "merch_chip_nr_dpa INTEGER, merch_nr_dpa INTEGER, max_nr_dpa INTEGER, ";

                foreach (string column in columnsFromCondTable)
                {
                    field = "";
                    dataType = "";
                    p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT " + column + " FROM " + m_strCondTable.Trim(), ref field, ref dataType);
                    dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                    fieldsAndDataTypes += field + " " + dataType + ", ";
                }

                fieldsAndDataTypes += "haul_costs_dpa INTEGER,";

                p_dataMgr.m_strSQL = "CREATE TABLE " + strInactiveStandsWorkTable + " (" + fieldsAndDataTypes + " PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "INSERT INTO " + strInactiveStandsWorkTable + " " + strSelectSQL;
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                p_dataMgr.m_strSQL = "INSERT INTO " + m_strEconByRxWorkTableName +
                    " SELECT isw.biosum_cond_id, isw.rxpackage, isw.rx, isw.rxcycle, isw.merch_vol_cf, " +
                    "isw.chip_vol_cf, isw.chip_wt_gt, isw.chip_val_dpa, isw.merch_wt_gt, isw.merch_val_dpa, " +
                    "isw.harvest_onsite_cost_dpa, isw.escalator_merch_haul_cpa_pt, isw.merch_haul_cost_dpa, " +
                    "isw.escalator_chip_haul_cpa_pt, isw.chip_haul_cost_dpa, isw.merch_chip_nr_dpa, " +
                    "isw.merch_nr_dpa, 'N' AS usebiomass_yn, isw.max_nr_dpa, isw.acres, isw.owngrpcd, isw.haul_costs_dpa " +
                    "FROM " + strInactiveStandsWorkTable + " AS isw " +
                    "LEFT OUTER JOIN " + m_strEconByRxWorkTableName + " AS w " +
                    "ON w.biosum_cond_id = isw.biosum_cond_id AND w.rxpackage = isw.rxpackage AND w.rx = isw.rx AND " +
                    "w.rxcycle = isw.rxcycle " +
                    "WHERE w.biosum_cond_id IS NULL";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName +
                                    "(biosum_cond_id,rxpackage,rx,rxcycle," +
                                    "merch_vol_cf,chip_vol_cf," +
                                    "chip_wt_gt,chip_val_dpa," +
                                    "merch_wt_gt,merch_val_dpa,harvest_onsite_cost_dpa," +
                                    "merch_haul_cost_dpa,chip_haul_cost_dpa,merch_chip_nr_dpa," +
                                    "merch_nr_dpa,usebiomass_yn,max_nr_dpa,acres,owngrpcd,haul_costs_dpa) " +
                                "SELECT biosum_cond_id,rxpackage,rx,rxcycle," +
                                    "merch_vol_cf,chip_vol_cf," +
                                    "chip_wt_gt,chip_val_dpa," +
                                    "merch_wt_gt,merch_val_dpa,harvest_onsite_cost_dpa," +
                                    "merch_haul_cost_dpa,chip_haul_cost_dpa,merch_chip_nr_dpa," +
                                    "merch_nr_dpa,usebiomass_yn,max_nr_dpa,acres,owngrpcd,haul_costs_dpa " +
                                "FROM " + m_strEconByRxWorkTableName;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            }

            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
            }

            p_dataMgr = null;
        }


        private void calculate_weighted_econ_variables()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//calculate_weighted_econ_variables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                   ReferenceUserControlScenarioRun.listViewEx1, "Calculate Weighted Economic Variables For Each Stand And Treatment Package");

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 5;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);

            DataMgr p_dataMgr = new DataMgr();

            // Query optimization and tiebreaker settings to see if there are weighted variables to calculate
            System.Collections.Generic.IList<string> lstFieldNames =
                new System.Collections.Generic.List<string>();
            // Optimization variable
            if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "ECONOMIC ATTRIBUTE")
            {
                string[] strCol = frmMain.g_oUtils.ConvertListToArray(this.m_oOptimizationVariable.strFVSVariableName, "_");
                if (strCol.Length > 1)
                {
                    // This is not a default economic variable; They always end in _1
                    if (strCol[strCol.Length - 1] != "1")
                    {
                        lstFieldNames.Add(this.m_oOptimizationVariable.strFVSVariableName.Trim());
                    }
                }
            }

            // Dollars per acre filter
            if (this.m_oOptimizationVariable.bUseFilter == true)
            {
                string[] strCol = frmMain.g_oUtils.ConvertListToArray(this.m_oOptimizationVariable.strRevenueAttribute, "_");
                if (strCol.Length > 1)
                {
                    // This is not a default economic variable; They always end in _1
                    if (strCol[strCol.Length - 1] != "1")
                    {
                        if (!lstFieldNames.Contains(this.m_oOptimizationVariable.strRevenueAttribute.Trim()))
                        {
                            lstFieldNames.Add(this.m_oOptimizationVariable.strRevenueAttribute.Trim());
                        }
                    }
                }
            }

            // Tiebreaker
            if (this.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.m_oSavTieBreakerCollection.Item(1).bSelected == true)
            {
                string strFieldName = this.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.m_oSavTieBreakerCollection.Item(1).strFVSVariableName.Trim();
                string[] strCol = frmMain.g_oUtils.ConvertListToArray(strFieldName, "_");
                if (strCol.Length > 1)
                {
                    // This is not a default economic variable; They always end in _1
                    if (strCol[strCol.Length - 1] != "1")
                    {
                        string strFieldType = uc_optimizer_scenario_calculated_variables.getEconVariableType(strFieldName);
                        if (!String.IsNullOrEmpty(strFieldType))
                        {
                            // This is a valid economic variable type
                            if (!lstFieldNames.Contains(strFieldName))
                            {
                                lstFieldNames.Add(strFieldName);
                            }
                        }
                    }
                }
            }

            System.Collections.Generic.IList<uc_optimizer_scenario_calculated_variables.VariableItem> lstVariableItems =
               new System.Collections.Generic.List<uc_optimizer_scenario_calculated_variables.VariableItem>();  //Parallel list to lstFieldNames; Holds variable definitions

            if (lstFieldNames.Count > 0)
            {
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nCalculating these weighted economic variables\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "---------------------------\r\n");
                    foreach (string strFieldName in lstFieldNames)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, strFieldName + "\r\n");
                    }
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "---------------------------\r\n");
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                // Populate economic variable information from configuration database
                FIA_Biosum_Manager.uc_optimizer_scenario_calculated_variables.Variable_Collection oWeightedVariableCollection =
                    new FIA_Biosum_Manager.uc_optimizer_scenario_calculated_variables.Variable_Collection();
                FIA_Biosum_Manager.OptimizerScenarioTools oOptimizerScenarioTools = new OptimizerScenarioTools();
                oOptimizerScenarioTools.LoadWeightedVariables(oWeightedVariableCollection);
                foreach (string strVariableName in lstFieldNames)
                {
                    foreach (uc_optimizer_scenario_calculated_variables.VariableItem oVariableItem in oWeightedVariableCollection)
                    {
                        if (oVariableItem.strVariableType.Equals("ECON") && oVariableItem.strVariableName.Equals(strVariableName))
                        {
                            oOptimizerScenarioTools.loadEconomicVariableWeights(oVariableItem);
                            lstVariableItems.Add(oVariableItem);
                            break;
                        }
                    }
                }

                string strEconConn = p_dataMgr.GetConnectionString(m_strSystemResultsDbPathAndFile);

                try
                {
                    using (System.Data.SQLite.SQLiteConnection econConn = new System.Data.SQLite.SQLiteConnection(strEconConn))
                    {
                        econConn.Open();

                        //Add columns to post_economic_weighted table to receive the data
                        foreach (string strFieldName in lstFieldNames)
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Adding columns for: " + strFieldName + "\r\n");
                            p_dataMgr.AddColumn(econConn, Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName,
                                "c1_" + strFieldName, "DOUBLE", "");
                            p_dataMgr.AddColumn(econConn, Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName,
                                "c2_" + strFieldName, "DOUBLE", "");
                            p_dataMgr.AddColumn(econConn, Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName,
                                "c3_" + strFieldName, "DOUBLE", "");
                            p_dataMgr.AddColumn(econConn, Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName,
                                "c4_" + strFieldName, "DOUBLE", "");
                            p_dataMgr.AddColumn(econConn, Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName,
                                strFieldName, "DOUBLE", "");
                        }

                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                        System.Collections.Generic.IDictionary<string, ProductYields> dictProductYields =
                        new System.Collections.Generic.Dictionary<string, ProductYields>();
                        p_dataMgr.m_strSQL = "select * from " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                        p_dataMgr.SqlQueryReader(econConn, p_dataMgr.m_strSQL);
                        ProductYields oProductYields = null;

                        while (p_dataMgr.m_DataReader.Read())
                        {
                            string strCondId = p_dataMgr.m_DataReader["biosum_cond_id"].ToString().Trim();
                            string strRxPackage = p_dataMgr.m_DataReader["rxpackage"].ToString().Trim();
                            string strKey = strCondId + "_" + strRxPackage;
                            if (dictProductYields.ContainsKey(strKey))
                            {
                                oProductYields = dictProductYields[strKey];
                            }
                            else
                            {
                                oProductYields = new ProductYields(strCondId, strRxPackage);
                            }
                            string strRxCycle = p_dataMgr.m_DataReader["rxcycle"].ToString().Trim();
                            double dblChipYieldCf = Convert.ToDouble(p_dataMgr.m_DataReader["chip_vol_cf"]);
                            double dblMerchYieldCf = Convert.ToDouble(p_dataMgr.m_DataReader["merch_vol_cf"]);
                            double dblHarvestOnsiteCpa = Convert.ToDouble(p_dataMgr.m_DataReader["harvest_onsite_cost_dpa"]);
                            double dblMaxNrDpa = Convert.ToDouble(p_dataMgr.m_DataReader["max_nr_dpa"]);
                            double dblHaulMerchCpa = Convert.ToDouble(p_dataMgr.m_DataReader["merch_haul_cost_dpa"]);
                            double dblMerchChipNrDpa = Convert.ToDouble(p_dataMgr.m_DataReader["merch_chip_nr_dpa"]);
                            double dblHaulChipCpa = Convert.ToDouble(p_dataMgr.m_DataReader["chip_haul_cost_dpa"]);

                            switch (strRxCycle)
                            {
                                case "1":
                                    oProductYields.UpdateCycle1Yields(dblChipYieldCf, dblMerchYieldCf, dblHarvestOnsiteCpa,
                                        dblMaxNrDpa, dblHaulMerchCpa, dblMerchChipNrDpa, dblHaulChipCpa);
                                    break;
                                case "2":
                                    oProductYields.UpdateCycle2Yields(dblChipYieldCf, dblMerchYieldCf, dblHarvestOnsiteCpa,
                                        dblMaxNrDpa, dblHaulMerchCpa, dblMerchChipNrDpa, dblHaulChipCpa);
                                    break;
                                case "3":
                                    oProductYields.UpdateCycle3Yields(dblChipYieldCf, dblMerchYieldCf, dblHarvestOnsiteCpa,
                                        dblMaxNrDpa, dblHaulMerchCpa, dblMerchChipNrDpa, dblHaulChipCpa);
                                    break;
                                case "4":
                                    oProductYields.UpdateCycle4Yields(dblChipYieldCf, dblMerchYieldCf, dblHarvestOnsiteCpa,
                                        dblMaxNrDpa, dblHaulMerchCpa, dblMerchChipNrDpa, dblHaulChipCpa);
                                    break;
                            }
                            dictProductYields[strKey] = oProductYields;
                        }

                        if (dictProductYields.Keys.Count > 0)
                        {
                            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                            string strSqlPrefix = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName +
                                        " (biosum_cond_id, rxpackage, ";
                            foreach (string strFieldName in lstFieldNames)
                            {
                                strSqlPrefix = strSqlPrefix + "c1_" + strFieldName + " , c2_" +
                                               strFieldName + ", c3_" + strFieldName + ", c4_" +
                                               strFieldName + ", " + strFieldName + ",";
                            }
                            strSqlPrefix = strSqlPrefix.TrimEnd(strSqlPrefix[strSqlPrefix.Length - 1]); //trim trailing comma
                            strSqlPrefix = strSqlPrefix + " ) VALUES ( '";

                            foreach (string strKey in dictProductYields.Keys)
                            {
                                ProductYields oSavedProductYields = dictProductYields[strKey];
                                string strSql = strSqlPrefix + oSavedProductYields.CondId() + "', '" +
                                    oSavedProductYields.RxPackage() + "',";

                                System.Collections.Generic.IList<double> lstFieldValues = new System.Collections.Generic.List<double>();
                                int i = 0;
                                foreach (string strFieldName in lstFieldNames)
                                {
                                    string strFieldType = uc_optimizer_scenario_calculated_variables.getEconVariableType(strFieldName);
                                    uc_optimizer_scenario_calculated_variables.VariableItem oVariableItem = lstVariableItems[i];
                                    System.Collections.Generic.IList<double> lstWeights = oVariableItem.lstWeights;
                                    switch (strFieldType)
                                    {
                                        case uc_optimizer_scenario_calculated_variables.PREFIX_MERCH_VOLUME:
                                            lstFieldValues.Add(oSavedProductYields.MerchYieldCfCycle1() * lstWeights[0]);
                                            lstFieldValues.Add(oSavedProductYields.MerchYieldCfCycle2() * lstWeights[1]);
                                            lstFieldValues.Add(oSavedProductYields.MerchYieldCfCycle3() * lstWeights[2]);
                                            lstFieldValues.Add(oSavedProductYields.MerchYieldCfCycle4() * lstWeights[3]);
                                            lstFieldValues.Add(oSavedProductYields.MerchYieldCfCycle1() * lstWeights[0] +
                                                               oSavedProductYields.MerchYieldCfCycle2() * lstWeights[1] +
                                                               oSavedProductYields.MerchYieldCfCycle3() * lstWeights[2] +
                                                               oSavedProductYields.MerchYieldCfCycle4() * lstWeights[3]);
                                            break;
                                        case uc_optimizer_scenario_calculated_variables.PREFIX_CHIP_VOLUME:
                                            lstFieldValues.Add(oSavedProductYields.ChipYieldCfCycle1() * lstWeights[0]);
                                            lstFieldValues.Add(oSavedProductYields.ChipYieldCfCycle2() * lstWeights[1]);
                                            lstFieldValues.Add(oSavedProductYields.ChipYieldCfCycle3() * lstWeights[2]);
                                            lstFieldValues.Add(oSavedProductYields.ChipYieldCfCycle4() * lstWeights[3]);
                                            lstFieldValues.Add(oSavedProductYields.ChipYieldCfCycle1() * lstWeights[0] +
                                                               oSavedProductYields.ChipYieldCfCycle2() * lstWeights[1] +
                                                               oSavedProductYields.ChipYieldCfCycle3() * lstWeights[2] +
                                                               oSavedProductYields.ChipYieldCfCycle4() * lstWeights[3]);
                                            break;
                                        case uc_optimizer_scenario_calculated_variables.PREFIX_TOTAL_VOLUME:
                                            lstFieldValues.Add(oSavedProductYields.TotalYieldCfCycle1() * lstWeights[0]);
                                            lstFieldValues.Add(oSavedProductYields.TotalYieldCfCycle2() * lstWeights[1]);
                                            lstFieldValues.Add(oSavedProductYields.TotalYieldCfCycle3() * lstWeights[2]);
                                            lstFieldValues.Add(oSavedProductYields.TotalYieldCfCycle4() * lstWeights[3]);
                                            lstFieldValues.Add(oSavedProductYields.TotalYieldCfCycle1() * lstWeights[0] +
                                                               oSavedProductYields.TotalYieldCfCycle2() * lstWeights[1] +
                                                               oSavedProductYields.TotalYieldCfCycle3() * lstWeights[2] +
                                                               oSavedProductYields.TotalYieldCfCycle4() * lstWeights[3]);
                                            break;
                                        case uc_optimizer_scenario_calculated_variables.PREFIX_NET_REVENUE:
                                            lstFieldValues.Add(oSavedProductYields.MaxNrDpaCycle1() * lstWeights[0]);
                                            lstFieldValues.Add(oSavedProductYields.MaxNrDpaCycle2() * lstWeights[1]);
                                            lstFieldValues.Add(oSavedProductYields.MaxNrDpaCycle3() * lstWeights[2]);
                                            lstFieldValues.Add(oSavedProductYields.MaxNrDpaCycle4() * lstWeights[3]);
                                            lstFieldValues.Add(oSavedProductYields.MaxNrDpaCycle1() * lstWeights[0] +
                                                               oSavedProductYields.MaxNrDpaCycle2() * lstWeights[1] +
                                                               oSavedProductYields.MaxNrDpaCycle3() * lstWeights[2] +
                                                               oSavedProductYields.MaxNrDpaCycle4() * lstWeights[3]);
                                            break;
                                        case uc_optimizer_scenario_calculated_variables.PREFIX_ONSITE_TREATMENT_COSTS:
                                            lstFieldValues.Add(oSavedProductYields.HarvestOnsiteCpaCycle1() * lstWeights[0]);
                                            lstFieldValues.Add(oSavedProductYields.HarvestOnsiteCpaCycle2() * lstWeights[1]);
                                            lstFieldValues.Add(oSavedProductYields.HarvestOnsiteCpaCycle3() * lstWeights[2]);
                                            lstFieldValues.Add(oSavedProductYields.HarvestOnsiteCpaCycle4() * lstWeights[3]);
                                            lstFieldValues.Add(oSavedProductYields.HarvestOnsiteCpaCycle1() * lstWeights[0] +
                                                               oSavedProductYields.HarvestOnsiteCpaCycle2() * lstWeights[1] +
                                                               oSavedProductYields.HarvestOnsiteCpaCycle3() * lstWeights[2] +
                                                               oSavedProductYields.HarvestOnsiteCpaCycle4() * lstWeights[3]);
                                            break;
                                        case uc_optimizer_scenario_calculated_variables.PREFIX_TREATMENT_HAUL_COSTS:
                                            lstFieldValues.Add(oSavedProductYields.TreatmentHaulCostsCycle1() * lstWeights[0]);
                                            lstFieldValues.Add(oSavedProductYields.TreatmentHaulCostsCycle2() * lstWeights[1]);
                                            lstFieldValues.Add(oSavedProductYields.TreatmentHaulCostsCycle3() * lstWeights[2]);
                                            lstFieldValues.Add(oSavedProductYields.TreatmentHaulCostsCycle4() * lstWeights[3]);
                                            lstFieldValues.Add(oSavedProductYields.TreatmentHaulCostsCycle1() * lstWeights[0] +
                                                               oSavedProductYields.TreatmentHaulCostsCycle2() * lstWeights[1] +
                                                               oSavedProductYields.TreatmentHaulCostsCycle3() * lstWeights[2] +
                                                               oSavedProductYields.TreatmentHaulCostsCycle4() * lstWeights[3]);
                                            break;
                                        default:
                                            lstFieldValues.Add(-1.0);
                                            lstFieldValues.Add(-1.0);
                                            lstFieldValues.Add(-1.0);
                                            lstFieldValues.Add(-1.0);
                                            lstFieldValues.Add(-1.0);
                                            break;
                                    }
                                    i++;
                                }

                                foreach (double dblFieldValue in lstFieldValues)
                                {
                                    strSql = strSql + dblFieldValue + " ,";
                                }
                                strSql = strSql.TrimEnd(strSql[strSql.Length - 1]); //trim trailing comma
                                strSql = strSql + " ) ";
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strSql + "\r\n");

                                p_dataMgr.SqlNonQuery(econConn, strSql);
                            }
                        }
                    }

                }
                catch (Exception err)
                {
                    p_dataMgr.m_intError = -1;
                    p_dataMgr.m_strError = "Error calculating weighted economic variables: " + err.Message;
                    MessageBox.Show("!! " + p_dataMgr.m_strError + " !!", "FIA Biosum");
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug) frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }
            }
            else
            {
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nWeighted economic variables are not used in this scenario\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "---------------------------\r\n");
                }
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 2;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            }
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
            }
            p_dataMgr = null;
        }

        /// <summary>
        /// get the wood product yields,
        /// revenue, and costs of an applied
        /// treatment on a plot 
        /// </summary>
        private void econ_by_rx_utilized_sum()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//econ_by_rx_utilized_sum\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nWood Product Yields,Revenue, And Costs Table By Treatment Package\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "-------------------------------------------------------\r\n");
            }

            intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                      ReferenceUserControlScenarioRun.listViewEx1, "Summarize Wood Product Volume Yields, Costs, And Net Revenue For A Stand And Treatment Package");

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 6;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);

            DataMgr p_dataMgr = new DataMgr();

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strSystemResultsDbPathAndFile + "' AS results";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                /********************************************
                 **delete all records in the table
                 ********************************************/
                p_dataMgr.m_strSQL = "DELETE FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName;
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                // We manipulate econ_by_rx_worktable to zero out some chip fields if use_biomass_yn='N'
                // Then we sum from the worktable to get the correct numbers in the summary fields
                p_dataMgr.m_strSQL = "UPDATE " + m_strEconByRxWorkTableName +
                    " SET chip_vol_cf = CASE WHEN usebiomass_yn = 'N' THEN 0 ELSE chip_vol_cf END, " +
                    "chip_wt_gt = CASE WHEN usebiomass_yn = 'N' THEN 0 ELSE chip_wt_gt END, " +
                    "chip_val_dpa = CASE WHEN usebiomass_yn = 'N' THEN 0 ELSE chip_val_dpa END, " +
                    "chip_haul_cost_dpa = CASE WHEN usebiomass_yn = 'N' THEN 0 ELSE chip_haul_cost_dpa END";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate " + m_strEconByRxWorkTableName + " based on use_biomass_yn values\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");

                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert Records\r\n");

                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName +
                    " (biosum_cond_id, rxpackage, chip_vol_cf_utilized, merch_vol_cf, chip_wt_gt_utilized, merch_wt_gt, chip_val_dpa_utilized, " +
                    "merch_val_dpa, harvest_onsite_cost_dpa, chip_haul_cost_dpa_utilized, merch_haul_cost_dpa, merch_chip_nr_dpa, " +
                    "merch_nr_dpa, max_nr_dpa, haul_costs_dpa, treated_acres,acres, owngrpcd) " +
                    "SELECT a.biosum_cond_id, a.rxpackage, a.sum_chip_yield_cf AS chip_vol_cf_utilized, a.sum_merch_yield_cf AS merch_vol_cf, " +
                    "a.sum_chip_yield_gt AS chip_wt_gt_utilized, a.sum_merch_yield_gt AS merch_wt_gt, a.sum_chip_val_dpa AS chip_val_dpa_utilized, " +
                    "a.sum_merch_val_dpa AS merch_val_dpa, a.sum_harvest_onsite_cpa AS harvest_onsite_cost_dpa, a.sum_haul_chip_cpa AS chip_haul_cost_dpa_utilized, " +
                    "a.sum_haul_merch_cpa AS merch_haul_cost_dpa, a.sum_merch_chip_nr_dpa AS merch_chip_nr_dpa, a.sum_merch_nr_dpa AS merch_nr_dpa, " +
                    "a.sum_max_nr_dpa AS max_nr_dpa, a.sum_haul_costs_dpa AS haul_costs_dpa, a.sum_treated_acres AS treated_acres, a.acres, a.owngrpcd " +
                    "FROM (SELECT biosum_cond_id, rxpackage, SUM(IFNULL(chip_vol_cf, 0)) AS sum_chip_yield_cf, SUM(IFNULL(merch_vol_cf, 0)) AS sum_merch_yield_cf, " +
                    "SUM(IFNULL(chip_wt_gt, 0)) AS sum_chip_yield_gt, SUM(IFNULL(merch_wt_gt, 0)) AS sum_merch_yield_gt, SUM(IFNULL(chip_val_dpa, 0)) AS sum_chip_val_dpa, " +
                    "SUM(IFNULL(merch_val_dpa, 0)) AS sum_merch_val_dpa, SUM(IFNULL(harvest_onsite_cost_dpa, 0)) AS sum_harvest_onsite_cpa, SUM(IFNULL(chip_haul_cost_dpa, 0)) AS sum_haul_chip_cpa, " +
                    "SUM(IFNULL(merch_haul_cost_dpa, 0)) AS sum_haul_merch_cpa, SUM(IFNULL(merch_chip_nr_dpa, 0)) AS sum_merch_chip_nr_dpa, SUM(IFNULL(merch_nr_dpa, 0)) AS sum_merch_nr_dpa, " +
                    "SUM(IFNULL(max_nr_dpa, 0)) AS sum_max_nr_dpa, SUM(IFNULL(haul_costs_dpa, 0)) AS sum_haul_costs_dpa, SUM(IFNULL(acres, 0)) AS sum_treated_acres, acres, owngrpcd " +
                    "FROM " + m_strEconByRxWorkTableName + " GROUP BY biosum_cond_id, rxpackage, acres, owngrpcd) AS a";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");

                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                // Create worktable for HVST_TYPE_BY_CYCLE field
                string strTreeSumWorktableName = "ECON_SUM_WORKTABLE";
                if (p_dataMgr.TableExist(workConn, strTreeSumWorktableName))
                {
                    p_dataMgr.SqlNonQuery(workConn, "DROP TABLE " + strTreeSumWorktableName);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nDelete existing econ_sum_work table\r\n");

                }

                p_dataMgr.m_strSQL = "CREATE TABLE " + strTreeSumWorktableName + " (" +
                                "biosum_cond_id CHAR(25), " +
                                "rxpackage CHAR(3), " +
                                "CYCLE_1 CHAR(1), " +
                                "CYCLE_2 CHAR(1), " +
                                "CYCLE_3 CHAR(1), " +
                                "CYCLE_4 CHAR(1))";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nCreate econ_sum_worktable\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                if (p_dataMgr.m_intError != 0)
                {
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    this.m_intError = p_dataMgr.m_intError;
                    return;
                }

                //Add rows to worktable
                p_dataMgr.m_strSQL = "INSERT INTO " + strTreeSumWorktableName +
                                " SELECT biosum_cond_id, rxpackage, '0' as cycle_1, '0' AS cycle_2, '0' AS cycle_3, '0' AS cycle_4" +
                                " FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nAdd condition/rxpackages to econ_sum_worktable\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    this.m_intError = p_dataMgr.m_intError;
                    return;
                }

                // Update worktable from econ_by_rx_cycle
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate cycle values on econ_sum_worktable\r\n");

                string[] arrFieldToUpdate = { "CYCLE_1", "CYCLE_2", "CYCLE_3", "CYCLE_4" };
                string[] arrRxCycle = { "1", "2", "3", "4" };
                for (int arrIdx = 0; arrIdx < 4; arrIdx++)
                {
                    p_dataMgr.m_strSQL = "UPDATE " + strTreeSumWorktableName +
                        " SET " + arrFieldToUpdate[arrIdx] + " = " +
                        "CASE WHEN a.chip_vol_cf > 0 AND a.merch_vol_cf > 0 THEN '3' WHEN a.chip_vol_cf = 0 AND a.merch_vol_cf = 0 THEN '0' " +
                        "WHEN a.chip_vol_cf > 0 THEN '2' ELSE '1' END " +
                        "FROM (SELECT biosum_cond_id, rxpackage, rxcycle, chip_vol_cf, merch_vol_cf FROM econ_by_rx_cycle) AS a " +
                        "WHERE " + strTreeSumWorktableName + ".biosum_cond_id = a.biosum_cond_id AND " + strTreeSumWorktableName + ".rxpackage = " +
                        "a.rxpackage AND a.rxcycle = '" + arrRxCycle[arrIdx] + "'";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                }
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    this.m_intError = p_dataMgr.m_intError;
                    return;
                }

                // Populate hvst_type_by_cycle from worktable
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate hvst_type_by_cycle from econ_sum_worktable\r\n");

                p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName +
                    " SET hvst_type_by_cycle = (SELECT cycle_1 || cycle_2 || cycle_3 || cycle_4 " +
                    "FROM " + strTreeSumWorktableName + " WHERE " + strTreeSumWorktableName + ".biosum_cond_id = " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName + ".biosum_cond_id " +
                    "AND " + strTreeSumWorktableName + ".rxpackage = " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName + ".rxpackage)";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.m_intError != 0)
                {
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    this.m_intError = p_dataMgr.m_intError;
                    return;
                }
            }

            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
            }
            p_dataMgr = null;
        }

        /// <summary>
        /// create a temporary work table for summing harvest costs
        /// </summary>
        private void CreateTableStructureOfHarvestCosts()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureOfHarvestCosts\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();

            /*********************************************
            * set the application version in the database
            * *******************************************/
            string strConn = p_dataMgr.GetConnectionString(this.m_strSystemResultsDbPathAndFile);
            using (System.Data.SQLite.SQLiteConnection resultsConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                resultsConn.Open();
                p_dataMgr.m_strSQL = "INSERT INTO VERSION (APPLICATION_VERSION)" +
                                " VALUES ('" + frmMain.g_strAppVer + "') ";
                p_dataMgr.SqlNonQuery(resultsConn, p_dataMgr.m_strSQL);
            }
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                /**************************************************
				 **create harvest_cost_sum table structure 
				 **in the temp work table db file
				 **************************************************/
                if (p_dataMgr.TableExist(workConn, "harvest_costs_sum"))
                {
                    p_dataMgr.m_strSQL = "DROP TABLE harvest_costs_sum";
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    if (p_dataMgr.m_intError != 0)
                    {
                        if (frmMain.g_bDebug)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "!! Error Deleting harvest_costs_sum Table!!\r\n");
                        this.m_intError = p_dataMgr.m_intError;
                        return;
                    }
                }

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Copy table structure harest_costs to harvest_costs_sum\r\n");

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + this.m_strProcessorResultsPathAndFile + "' AS processor_results";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                string[] arrFields = p_dataMgr.getFieldNamesArray(workConn, "SELECT biosum_cond_id,rxpackage,rx,rxcycle, complete_cpa FROM harvest_costs");

                string fieldsAndDataTypes = "";

                foreach (string column in arrFields)
                {
                    string field = "";
                    string dataType = "";
                    p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT " + column + " FROM harvest_costs", ref field, ref dataType);
                    dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                    fieldsAndDataTypes = fieldsAndDataTypes + field + " " + dataType + ", ";
                }
                p_dataMgr.m_strSQL = "CREATE TABLE harvest_costs_sum (" + fieldsAndDataTypes + "PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
            }
        }

        private void CreateTableStructureForScenarioProcessingSites()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForScenarioProcessingSites\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();

            /***********************************************************************
            **make a copy of the scenario_psites table and give it the
            **name scenario_psites_work_table
            ***********************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Delete table scenario_psites_work_table\r\n");

            this.m_strPSiteWorkTable = "scenario_psites_work_table";

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                if (p_dataMgr.TableExist(workConn, this.m_strPSiteWorkTable))
                {
                    p_dataMgr.m_strSQL = "DROP TABLE " + this.m_strPSiteWorkTable;
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    if (p_dataMgr.m_intError != 0)
                    {
                        if (frmMain.g_bDebug)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "!! Error Deleting scenario_psites_work_table Table!!\r\n");
                        this.m_intError = p_dataMgr.m_intError;
                        p_dataMgr = null;
                        return;
                    }
                }

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Copy table structure scenario_psites to scenario_psites_work_tabler\r\n");

                string strRuleDefinitionsDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + strRuleDefinitionsDb + "' AS rule_defs";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                string[] arrFields = p_dataMgr.getFieldNamesArray(workConn, "SELECT psite_id,trancd,biocd FROM scenario_psites");

                string fieldsAndDataTypes = "";

                foreach (string column in arrFields)
                {
                    string field = "";
                    string dataType = "";
                    p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT " + column + " FROM scenario_psites", ref field, ref dataType);
                    dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                    fieldsAndDataTypes = fieldsAndDataTypes + field + " " + dataType + ", ";
                }

                p_dataMgr.m_strSQL = "CREATE TABLE " + this.m_strPSiteWorkTable + " (" + fieldsAndDataTypes.Substring(0, fieldsAndDataTypes.Length - 2) + ")";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
            }
            p_dataMgr = null;
        }

        /// <summary>
        /// create temporary work tables for getting 
        /// a plot's cheapest merch and chip haul cost
        /// </summary>
        private void CreateTableStructureForHaulCosts()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForHaulCosts\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();

            /*****************************************************************
            **create the table structures in the temp work tables db file
            **and give them the name OF all_road_merch_haul_costs_work_table and 
            **                          all_road_chip_haul_costs_work_table
            **                          cheapest_road_merch_haul_costs_work_table
            **                          cheapest_road_chip_haul_costs_work_table
            **                          cheapest_rail_merch_haul_costs_work_table
            **                          cheapest_rail_chip_haul_costs_work_table
            **                          merch_plot_to_rh_to_collector_haul_costs_work_table
            **                          chip_plot_to_rh_to_collector_haul_costs_work_table
            **                          combine_chip_rail_road_haul_costs_work_table
            **                          combine_merch_rail_road_haul_costs_work_table
            **                          psite_accessible_work_table
            *****************************************************************/

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--Create psite_accessible_work_table\r\n");

                if (p_dataMgr.TableExist(workConn, Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName))
                    p_dataMgr.SqlNonQuery(workConn, "DROP TABLE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName);

                if (p_dataMgr.m_intError == 0)
                    frmMain.g_oTables.m_oOptimizerScenarioResults.CreatePSitesWorktableSqlite(
                        p_dataMgr, workConn, Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName);

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--Create haul costs work tables from the haul_costs table SQL\r\n");

                string[] arrHaulCostWorkTables = { "all_road_chip_haul_costs_work_table", "all_road_merch_haul_costs_work_table", "cheapest_road_merch_haul_costs_work_table",
                "cheapest_road_chip_haul_costs_work_table", "cheapest_rail_merch_haul_costs_work_table", "cheapest_rail_chip_haul_costs_work_table", "cheapest_merch_haul_costs_work_table",
                "cheapest_chip_haul_costs_work_table", "chip_plot_to_rh_to_collector_haul_costs_work_table", "merch_plot_to_rh_to_collector_haul_costs_work_table",
                "combine_chip_rail_road_haul_costs_work_table", "combine_merch_rail_road_haul_costs_work_table"};

                foreach (string table in arrHaulCostWorkTables)
                {
                    if (p_dataMgr.m_intError == 0)
                    {
                        if (p_dataMgr.TableExist(workConn, table))
                        {
                            p_dataMgr.SqlNonQuery(workConn, "DROP TABLE " + table);
                        }

                        frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteHaulCostTable(
                            p_dataMgr, workConn, table);
                    }
                }

                /*****************************************************************
                **create the railroad table structures in the temp work tables file
                **and give them the name OF merch_rh_to_collector_haul_costs_work_table
                **                          chip_rh_to_collector_haul_costs_work_table
                *****************************************************************/

                string[] arrHaulCostRailroadTables = { "merch_rh_to_collector_haul_costs_work_table", "chip_rh_to_collector_haul_costs_work_table" };

                foreach (string table in arrHaulCostRailroadTables)
                {
                    if (p_dataMgr.m_intError == 0)
                    {
                        if (p_dataMgr.TableExist(workConn, table))
                        {
                            p_dataMgr.SqlNonQuery(workConn, "DROP TABLE " + table);
                        }

                        frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteHaulCostRailroadTable(
                            p_dataMgr, workConn, table);
                    }
                }

                if (p_dataMgr.m_intError != 0)
                {

                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    return;
                }
            }
            p_dataMgr = null;
        }

        /// <summary>
        /// create a temporary work tables for finding best treatments
        /// </summary>
        private void CreateTableStructureForIntensity()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForIntensity\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                /*****************************************************************
                   **create the table structure in the temp work tables db file
                   **and give it the name of rx_intensity_work_table
                   *****************************************************************/
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--Create rx_intensity_work_table Schema--\r\n");

                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateIntensityWorkTableSqlite(p_dataMgr, workConn, "rx_intensity_duplicates_work_table");
                if (p_dataMgr.m_intError == 0)
                {
                    frmMain.g_oTables.m_oOptimizerScenarioResults.CreateIntensityWorkTableSqlite(p_dataMgr, workConn, "rx_intensity_duplicates_work_table2");
                }
                if (p_dataMgr.m_intError == 0)
                {
                    frmMain.g_oTables.m_oOptimizerScenarioResults.CreateIntensityWorkTableSqlite(p_dataMgr, workConn, "rx_intensity_unique_work_table");
                }

                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    return;
                }
            }
        }

        /// <summary>
        /// create a temporary work tables for identifying inaccessible plots and conditions
        /// </summary>
        private void CreateTableStructureForPlotCondAccessiblity()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForPlotCondAccessiblity\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Delete table plot_cond_accessible_work_table\r\n");

                if (p_dataMgr.TableExist(workConn, "plot_cond_accessible_work_table"))
                {
                    p_dataMgr.m_strSQL = "DROP TABLE plot_cond_accessible_work_table";
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    if (p_dataMgr.m_intError != 0)
                    {
                        if (frmMain.g_bDebug)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "!! Error Deleting plot_cond_accessible_work_table Table!!\r\n");
                        this.m_intError = p_dataMgr.m_intError;
                        p_dataMgr = null;
                        return;
                    }
                }

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Copy table structure plot to plot_cond_accessible_work_table\r\n");

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strPlotPathAndFile + "' AS master";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                string[] arrFields = p_dataMgr.getFieldNamesArray(workConn, "SELECT biosum_plot_id, num_cond FROM " + this.m_strPlotTable);

                string fieldsAndDataTypes = "";
                string numCondDataType = "";

                foreach (string column in arrFields)
                {
                    string field = "";
                    string dataType = "";
                    p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT " + column + " FROM " + this.m_strPlotTable, ref field, ref dataType);
                    dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                    if (field == "num_cond")
                    {
                        numCondDataType = dataType;
                    }
                    fieldsAndDataTypes = fieldsAndDataTypes + field + " " + dataType + ", ";
                }

                p_dataMgr.m_strSQL = "CREATE TABLE plot_cond_accessible_work_table (" + fieldsAndDataTypes + "num_cond_not_accessible " + numCondDataType + ")";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "CREATE TABLE plot_cond_accessible_work_table2 (" + fieldsAndDataTypes + "num_cond_not_accessible " + numCondDataType + ")";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
            }

            p_dataMgr = null;
        }

        private void CreateTableStructureForUserDefinedConditionTable()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureOfUserDefinedCondSQL\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                if (p_dataMgr.TableExist(workConn, "userdefinedcondfilter_work"))
                {
                    p_dataMgr.m_strSQL = "DROP TABLE userdefinedcondfilter_work";
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    if (p_dataMgr.m_intError != 0)
                    {
                        if (frmMain.g_bDebug)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "!! Error Deleting userdefinedcondfilter_work Table!!\r\n");
                        this.m_intError = p_dataMgr.m_intError;
                        p_dataMgr = null;
                        return;
                    }
                }

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Create table structure userdefinedcondfilter_work\r\n");

                // Attach all databases available for use for cond filter
                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strCondPathAndFile + "' AS master";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strPSitePathAndFile + "' AS gis";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strHarvestMethodsPathAndFile + "' AS ref";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                string[] arrFields = p_dataMgr.getFieldNamesArray(workConn, this.m_strUserDefinedCondSQL);

                string fieldsAndDataTypes = "";

                foreach (string column in arrFields)
                {
                    string field = "";
                    string dataType = "";
                    p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT " + column + " FROM cond", ref field, ref dataType);
                    dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                    fieldsAndDataTypes = fieldsAndDataTypes + field + " " + dataType + ", ";
                }
                p_dataMgr.m_strSQL = "CREATE TABLE userdefinedcondfilter_work (" + fieldsAndDataTypes + "PRIMARY KEY (biosum_cond_id))";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                /***********************************************************************
			     **create ruledefinitionscondfilter table
			     ***********************************************************************/
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Delete table ruledefinitionscondfilter\r\n");

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strSystemResultsDbPathAndFile + "' AS results";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                if (p_dataMgr.AttachedTableExist(workConn, "ruledefinitionscondfilter"))
                {
                    p_dataMgr.m_strSQL = "DROP TABLE results.ruledefinitionscondfilter";
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    if (p_dataMgr.m_intError != 0)
                    {
                        if (frmMain.g_bDebug)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "!! Error Deleting ruledefinitionscondfilter Table!!\r\n");
                        this.m_intError = p_dataMgr.m_intError;
                        p_dataMgr = null;
                        return;
                    }
                }

                arrFields = p_dataMgr.getFieldNamesArray(workConn, "SELECT * FROM " + this.m_strCondTable);

                fieldsAndDataTypes = "";

                foreach (string column in arrFields)
                {
                    string field = "";
                    string dataType = "";
                    p_dataMgr.getFieldNamesAndDataTypes(workConn, "SELECT " + column + " FROM " + this.m_strCondTable, ref field, ref dataType);
                    dataType = utils.DataTypeConvert(dataType.ToUpper(), true);
                    fieldsAndDataTypes = fieldsAndDataTypes + field + " " + dataType + ", ";
                }
                p_dataMgr.m_strSQL = "CREATE TABLE results.ruledefinitionscondfilter (" + fieldsAndDataTypes + "PRIMARY KEY (biosum_cond_id))";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
            }

            p_dataMgr = null;
        }

		/// <summary>
		/// find the best treatment by these categories: 
		/// maximum net revenue; merchantable wood removal;
		/// and optimization variable
		/// </summary>
		
        private void Best_rx_summary()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Best_rx_summary\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            string strTable = "";
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Text", "Finding Best Treatments: Maximum Net Revenue");
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Visible", true);
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nBest Rx Summary\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "--------------------\r\n");

            intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                   ReferenceUserControlScenarioRun.listViewEx1, "Identify The Best Effective Treatment For Each Stand");

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 5;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;
            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);


            FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection oTieBreakerCollection =
                ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.m_oSavTieBreakerCollection;

            DataMgr p_dataMgr = new DataMgr();

            string strScenarioId = this.ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim();
            string strTieBreakerAggregate = "MAX";

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strPlotPathAndFile + "' AS master";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);


                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strSystemResultsDbPathAndFile + "' AS results";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                //
                //CREATE WORK TABLES
                //
                //best_rx_summary_work_table
                strTable = "cycle1_best_rx_summary_work_table";
                if (p_dataMgr.TableExist(workConn, strTable))
                {
                    p_dataMgr.SqlNonQuery(workConn, "DROP TABLE " + strTable);
                }
                p_dataMgr.SqlNonQuery(workConn, Tables.OptimizerScenarioResults.CreateSqliteBestRxSummaryTableSQL(strTable));
                
                //best_rx_summury_optimization_and_tiebreaker_work_table
                if (p_dataMgr.TableExist(workConn, "cycle1_best_rx_summary_optimization_and_tiebreaker_work_table"))
                {
                    p_dataMgr.SqlNonQuery(workConn, "DROP TABLE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table");
                }
                p_dataMgr.SqlNonQuery(workConn, Tables.OptimizerScenarioResults.CreateSqliteBestRxSummaryCycle1TieBreakerTableSQL("cycle1_best_rx_summary_optimization_and_tiebreaker_work_table"));
                
                //best_rx_summury_optimization_and_tiebreaker_work_table2
                if (p_dataMgr.TableExist(workConn, "cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2"))
                {
                    p_dataMgr.SqlNonQuery(workConn, "DROP TABLE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2");
                }
                p_dataMgr.SqlNonQuery(workConn, Tables.OptimizerScenarioResults.CreateSqliteBestRxSummaryCycle1TieBreakerTableSQL("cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2"));
                
                //best_rx_summury_optimization_and_tiebreaker_work_table3
                if (p_dataMgr.TableExist(workConn, "cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3"))
                {
                    p_dataMgr.SqlNonQuery(workConn, "DROP TABLE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3");
                }
                p_dataMgr.SqlNonQuery(workConn, Tables.OptimizerScenarioResults.CreateSqliteBestRxSummaryCycle1TieBreakerTableSQL("cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3"));

                /**********************************************
			     **insert unique biosum_cond_id's into the
			     **best_rx_summary table so we dont have
			     **to worry about whether the biosum_cond_id 
			     **record is in the table or not
			     **********************************************/
                p_dataMgr.m_strSQL = "INSERT INTO " + strTable +
                    " (biosum_cond_id, acres, owngrpcd) " +
                    "SELECT DISTINCT c.biosum_cond_id, c.acres, c.owngrpcd " +
                    "FROM " + m_strCondTable.Trim() + " AS c, " +
                    Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS p, " +
                    ReferenceOptimizerScenarioForm.OutputTablePrefix + Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix + " AS e " +
                    "WHERE c.biosum_cond_id = p.biosum_cond_id AND e.biosum_cond_id = c.biosum_cond_id AND " +
                    "e.affordable_YN = 'Y' AND e.rxcycle = '1' AND " +
                    "(p.merch_haul_cost_id IS NOT NULL OR p.chip_haul_cost_id IS NOT NULL)";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--insert condition records that have MERCH or CHIP haul costs into best_rx_summary--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                string strWorkTable = "cycle1_effective_" + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;
                if (p_dataMgr.TableExist(workConn, strWorkTable))
                {
                    p_dataMgr.SqlNonQuery(workConn, "DROP TABLE " + strWorkTable);
                }

                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteProductYieldsTable(p_dataMgr, workConn, strWorkTable);

                p_dataMgr.m_strSQL = "INSERT INTO " + strWorkTable +
                    " SELECT p.* FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName + " AS p, " +
                    ReferenceOptimizerScenarioForm.OutputTablePrefix + Tables.OptimizerScenarioResults.DefaultScenarioResultsEffectiveTableSuffix + " AS e " +
                    "WHERE p.biosum_cond_id = e.biosum_cond_id AND " +
                    "p.rxpackage=e.rxpackage AND p.rx=e.rx AND p.rxcycle=e.rxcycle AND e.overall_effective_yn='Y'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--write overall effective treatments to the " + strWorkTable + "--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n\r\n");

                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (p_dataMgr.TableExist(workConn, "cycle1_effective_optimization_treatments"))
                {
                    p_dataMgr.SqlNonQuery(workConn, "DROP TABLE cycle1_effective_optimization_treatments");
                }
                string strColumnFilterName = "";
                if (this.m_oOptimizationVariable.bUseFilter)
                {
                    strColumnFilterName = this.m_oOptimizationVariable.strRevenueAttribute;
                }
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteEffectiveTable(p_dataMgr, workConn, "cycle1_effective_optimization_treatments", strColumnFilterName);

                p_dataMgr.m_strSQL = "INSERT INTO cycle1_effective_optimization_treatments " +
                   "SELECT e.* FROM " + ReferenceOptimizerScenarioForm.OutputTablePrefix +
                   Tables.OptimizerScenarioResults.DefaultScenarioResultsEffectiveTableSuffix + " AS e " +
                   "WHERE e.overall_effective_yn = 'Y'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--write overall effective treatments to the cycle1_effective_optimization_treatments--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "REVENUE")
                {
                    Best_rx_summary(oTieBreakerCollection, strTieBreakerAggregate, false);

                }
                else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "MERCHANTABLE VOLUME")
                {
                    Best_rx_summary(oTieBreakerCollection, strTieBreakerAggregate, false);
                }
                else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "ECONOMIC ATTRIBUTE")
                {
                    Best_rx_summary(oTieBreakerCollection, strTieBreakerAggregate, false);
                }
                else
                {
                    Best_rx_summary(oTieBreakerCollection, strTieBreakerAggregate, true);
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                string strBestRxSummaryTableName = this.ReferenceOptimizerScenarioForm.OutputTablePrefix +
                    Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryTableSuffix;

                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                /****************************************************************************
                 **finished with minimum merchantable wood removal with positive net revenue
                 ****************************************************************************/
            }

            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
            }

            p_dataMgr = null;
        }

        private void Best_rx_summary(FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection oTieBreakerCollection,
            string strTieBreakerAggregate,
            bool bFVSVariable)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//" + ReferenceOptimizerScenarioForm.OutputTablePrefix +
                    Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix + "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n\r\n");

                frmMain.g_oUtils.WriteText(m_strDebugFile, "Parameters\r\n-------------\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "TieBreaker_Collection object\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "TieBreakerAggregate=" + strTieBreakerAggregate + "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "FVS Variable as Tie Breaker = ");
                if (bFVSVariable)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Yes\r\n");
                else
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "No\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + this.m_strSystemResultsDbPathAndFile + "' AS results";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                string strOptimizationTableName = ReferenceOptimizerScenarioForm.OutputTablePrefix +
                Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix;
                p_dataMgr.m_strSQL = "";
                if (bFVSVariable == false)
                {
                    //find the treatment for each plot that produces the MAX/MIN revenue value
                    p_dataMgr.m_strSQL = "SELECT a.biosum_cond_id, a.rxpackage, a.rx, a." + m_strOptimizationColumnNameSql + " AS optimization_value " +
                        "FROM " + strOptimizationTableName + " AS a, " +
                        "(SELECT " + m_strOptimizationAggregateSql + "(" + m_strOptimizationColumnNameSql + ") AS " + m_strOptimizationAggregateColumnName + ", biosum_cond_id " +
                        "FROM " + strOptimizationTableName + " WHERE affordable_YN = 'Y' " +
                        "GROUP BY biosum_cond_id) AS b " +
                        "WHERE a.biosum_cond_id = b.biosum_cond_id AND a." + m_strOptimizationColumnNameSql + " = b. " + m_strOptimizationAggregateColumnName +
                        " AND a.affordable_YN = 'Y'";
                }
                else
                {
                    p_dataMgr.m_strSQL = "SELECT a.biosum_cond_id, a.rxpackage, a.rx, a." + m_strOptimizationColumnNameSql + " AS optimization_value " +
                        "FROM " + strOptimizationTableName + " AS a, " +
                        "(SELECT " + m_strOptimizationAggregateSql + "(" + m_strOptimizationColumnNameSql + ") AS " + m_strOptimizationAggregateColumnName + ", biosum_cond_id " +
                        "FROM " + strOptimizationTableName + " WHERE affordable_YN = 'Y' " +
                        "GROUP BY biosum_cond_id) AS b " +
                        "WHERE a.biosum_cond_id = b.biosum_cond_id AND a." + m_strOptimizationColumnNameSql + " = b." + m_strOptimizationAggregateColumnName +
                        " AND a.affordable_YN = 'Y'";
                }

                p_dataMgr.m_strSQL = "INSERT INTO cycle1_best_rx_summary_optimization_and_tiebreaker_work_table (biosum_cond_id, rxpackage, rx, optimization_value) " +
                    p_dataMgr.m_strSQL;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--filter effective treatments to find " + this.m_strOptimizationAggregateSql + " " + this.m_oOptimizationVariable.strOptimizedVariable + "--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                if (p_dataMgr.m_intError != 0)
                {
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                    this.m_intError = p_dataMgr.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    return;
                }

                p_dataMgr.m_strSQL = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table AS a " +
                    "SET acres = b.acres, owngrpcd = b.owngrpcd " +
                    "FROM cycle1_best_rx_summary_work_table AS b " +
                    "WHERE a.biosum_cond_id = b.biosum_cond_id";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                //Stand OR Economic Attribute selected AND Last Tie-Break Rank

                if ((oTieBreakerCollection.Item(0).bSelected || oTieBreakerCollection.Item(1).bSelected) &&
                    oTieBreakerCollection.Item(2).bSelected)
                {
                    string strTiebreakerValueField = "post_variable1_value";    //Economic attributes will always write the post value
                    if (oTieBreakerCollection.Item(0).bSelected)    //FVS attribute selected
                    {
                        if (oTieBreakerCollection.Item(0).strValueSource == "POST-PRE")
                        {
                            strTiebreakerValueField = "variable1_change";
                        }
                    }

                    //update the tiebreaker and rx intensity fields for each plot
                    p_dataMgr.m_strSQL = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table AS a " +
                        "SET tiebreaker_value = b." + strTiebreakerValueField + ", last_tiebreak_rank = b.last_tiebreak_rank " +
                        "FROM tiebreaker AS b " +
                        "WHERE a.biosum_cond_id = b.biosum_cond_id AND a.rxpackage = b.rxpackage AND a.rx = b.rx";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    p_dataMgr.m_strSQL = "INSERT INTO " + ReferenceOptimizerScenarioForm.OutputTablePrefix +
                        Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryBeforeTiebreaksTableSuffix +
                        " SELECT DISTINCT * FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    //find the treatment for each plot that produces the MAX/MIN tiebreaker value
                    p_dataMgr.m_strSQL = "SELECT a.biosum_cond_id,a.rxpackage,a.rx,a.acres,a.owngrpcd,a.optimization_value,a.tiebreaker_value,a.last_tiebreak_rank " +
                        "FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table AS a," +
                        "(SELECT biosum_cond_id," + strTieBreakerAggregate + "(tiebreaker_value) AS tiebreaker " +
                        "FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table " +
                        "GROUP BY biosum_cond_id) AS c " +
                        "WHERE a.biosum_cond_id=c.biosum_cond_id AND a.tiebreaker_value=c.tiebreaker";

                    p_dataMgr.m_strSQL = "INSERT INTO cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 " + p_dataMgr.m_strSQL;

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "--break any ties by finding the " + strTieBreakerAggregate + " tie breaker value--\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                    if (p_dataMgr.m_intError != 0)
                    {
                        if (frmMain.g_bDebug)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                        this.m_intError = p_dataMgr.m_intError;
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    p_dataMgr.m_strSQL = "SELECT a.biosum_cond_id,a.rxpackage,a.rx,a.acres,a.owngrpcd,a.optimization_value," +
                        "a.tiebreaker_value,a.last_tiebreak_rank " +
                        "FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 AS a," +
                        "(SELECT biosum_cond_id,MIN(last_tiebreak_rank) AS min_intensity " +
                        "FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 " +
                        "GROUP BY biosum_cond_id) AS c " +
                        "WHERE a.biosum_cond_id=c.biosum_cond_id AND a.last_tiebreak_rank=c.min_intensity";

                    p_dataMgr.m_strSQL = "INSERT INTO cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3 " + p_dataMgr.m_strSQL;

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "--break any additional ties by finding the least intense treatment--\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                    if (p_dataMgr.m_intError != 0)
                    {
                        if (frmMain.g_bDebug)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                        this.m_intError = p_dataMgr.m_intError;
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    p_dataMgr.m_strSQL = "UPDATE cycle1_best_rx_summary_work_table AS a " +
                        "SET optimization_value = b.optimization_value, tiebreaker_value = b.tiebreaker_value, " +
                        "rxpackage = b.rxpackage, rx = b.rx, last_tiebreak_rank = b.last_tiebreak_rank " +
                        "FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3 AS b " +
                        "WHERE a.biosum_cond_id = b.biosum_cond_id";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    p_dataMgr.m_strSQL = "INSERT INTO " + ReferenceOptimizerScenarioForm.OutputTablePrefix +
                        Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryTableSuffix +
                        " SELECT * FROM cycle1_best_rx_summary_work_table";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "--insert the work table records into the best_rx_summary table--\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                    if (p_dataMgr.m_intError != 0)
                    {
                        if (frmMain.g_bDebug)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                        this.m_intError = p_dataMgr.m_intError;
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }
                }
                // Last tie-break rank ONLY
                else if (oTieBreakerCollection.Item(2).bSelected)
                {
                    //update the rx intensity fields for each plot
                    p_dataMgr.m_strSQL = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table AS a " +
                        "SET last_tiebreak_rank = b.last_tiebreak_rank " +
                        "FROM tiebreaker AS b " +
                        "WHERE a.biosum_cond_id = b.biosum_cond_id AND a.rx = b.rx AND a.rxpackage = b.rxpackage";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    p_dataMgr.m_strSQL = "INSERT INTO " +
                        ReferenceOptimizerScenarioForm.OutputTablePrefix +
                        Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryBeforeTiebreaksTableSuffix +
                        " SELECT DISTINCT * FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    p_dataMgr.m_strSQL = "SELECT a.biosum_cond_id,a.rxpackage,a.rx,a.acres,a.owngrpcd,a.optimization_value," +
                        "a.tiebreaker_value,a.last_tiebreak_rank " +
                        "FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table AS a," +
                        "(SELECT biosum_cond_id,MIN(last_tiebreak_rank) AS min_intensity " +
                        "FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table " +
                        "GROUP BY biosum_cond_id) AS c " +
                        "WHERE a.biosum_cond_id=c.biosum_cond_id AND a.last_tiebreak_rank=c.min_intensity";

                    p_dataMgr.m_strSQL = "INSERT INTO cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 " + p_dataMgr.m_strSQL;

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "--break any additional ties by finding the least intense treatment--\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                    if (p_dataMgr.m_intError != 0)
                    {
                        if (frmMain.g_bDebug)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                        this.m_intError = p_dataMgr.m_intError;
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    p_dataMgr.m_strSQL = "UPDATE cycle1_best_rx_summary_work_table AS a " +
                        "SET optimization_value = b.optimization_value, tiebreaker_value = b.tiebreaker_value, " +
                        "rxpackage = b.rxpackage, rx = b.rx, last_tiebreak_rank = b.last_tiebreak_rank " +
                        "FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 AS b " +
                        "WHERE a.biosum_cond_id = b.biosum_cond_id";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                    p_dataMgr.m_strSQL = "INSERT INTO " + ReferenceOptimizerScenarioForm.OutputTablePrefix +
                        Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryTableSuffix +
                        " SELECT * FROM cycle1_best_rx_summary_work_table";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "--insert the work table records into the best_rx_summary table--\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + p_dataMgr.m_strSQL + "\r\n\r\n");
                    p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

                    if (p_dataMgr.m_intError != 0)
                    {
                        if (frmMain.g_bDebug)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                        this.m_intError = p_dataMgr.m_intError;
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }
                }
            }
        }

		private void CreateHtml()
		{
			System.IO.FileStream oTxtFileStream;
			System.IO.StreamWriter oTxtStreamWriter;

			oTxtFileStream = new System.IO.FileStream(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\runstats.htm", System.IO.FileMode.Create, 
				System.IO.FileAccess.Write);
			oTxtStreamWriter = new System.IO.StreamWriter(oTxtFileStream);
			oTxtStreamWriter.WriteLine("<html>\r\n");
			oTxtStreamWriter.WriteLine("<head>\r\n");
			oTxtStreamWriter.WriteLine("<title>\r\n");
            oTxtStreamWriter.WriteLine("FIA Biosum Optimizer Scenario Run Summary Report\r\n");
			oTxtStreamWriter.WriteLine("</title>\r\n");
			oTxtStreamWriter.WriteLine("<body bgcolor='#ffffff' link='#33339a' vlink='#33339a' alink='#33339a'>\r\n");
			oTxtStreamWriter.WriteLine(System.DateTime.Now.ToString() + "\r\n");
			oTxtStreamWriter.WriteLine("<A NAME='GO TOP'></A>\r\n");
			oTxtStreamWriter.WriteLine("<BR> <BR>\r\n");
            oTxtStreamWriter.WriteLine("<b><CENTER><FONT SIZE='+2' >FIA Biosum Optimizer Scenario Run Summary Report</FONT></b></center><br>\r\n");
			oTxtStreamWriter.WriteLine("<CENTER>\r\n");
			oTxtStreamWriter.WriteLine("<!--REPORT: CONDITION SAMPLE STATUS-->\r\n");
			oTxtStreamWriter.WriteLine("<TABLE COLSPAN='4' border='1' WIDTH='70%' HEIGHT='2%' cellpadding='0' cellspacing='0'\r\n>");
			//
			//condition table land class code counts
			//
			oTxtStreamWriter.WriteLine("<TR>\r\n");
			oTxtStreamWriter.WriteLine("   <td colspan='4'  bgcolor='lightgray' VAlign=middle height='2%' width='100%' align='center'>\r\n");
			oTxtStreamWriter.WriteLine("      <b>CONDITION SAMPLE STATUS</b>\r\n");
			oTxtStreamWriter.WriteLine("   </TD>\r\n");
			oTxtStreamWriter.WriteLine("</TR>\r\n");
			oTxtStreamWriter.WriteLine("<!--column headers-->\r\n");
			oTxtStreamWriter.WriteLine("<TR>\r\n");
			oTxtStreamWriter.WriteLine("	<td colspan='1'  bgcolor='lightyellow' VAlign=middle height='2%' width='25%' align='center'>\r\n");
			oTxtStreamWriter.WriteLine("		<b>Description</b>\r\n");
			oTxtStreamWriter.WriteLine("	</TD>\r\n");
			oTxtStreamWriter.WriteLine("	<td colspan='1'  bgcolor='lightyellow' VAlign=middle height='2%' width='25%' align='center'>\r\n");
			oTxtStreamWriter.WriteLine("		<b>Count</b>\r\n");
			oTxtStreamWriter.WriteLine("	</TD>\r\n");
			oTxtStreamWriter.WriteLine("	<td colspan='1'  bgcolor='lightyellow' VAlign=middle height='2%' width='25%' align='center'>\r\n");
			oTxtStreamWriter.WriteLine("		<b>With Trees</b>\r\n");
			oTxtStreamWriter.WriteLine("	</TD>\r\n");
			oTxtStreamWriter.WriteLine("	<td colspan='1'  bgcolor='lightyellow' VAlign=middle height='2%' width='25%' align='center'>\r\n");
			oTxtStreamWriter.WriteLine("		<b>Without Trees</b>\r\n");
			oTxtStreamWriter.WriteLine("	</TD>\r\n");
			oTxtStreamWriter.WriteLine("</TR>\r\n");
			oTxtStreamWriter.WriteLine("<!--count data-->\r\n");
			oTxtStreamWriter.WriteLine("<TR>\r\n");



			oTxtStreamWriter.WriteLine("</TABLE>\r\n");
			oTxtStreamWriter.WriteLine("</CENTER>\r\n");
			oTxtStreamWriter.WriteLine("</BODY>\r\n");
			oTxtStreamWriter.WriteLine("</HEAD>\r\n");
			oTxtStreamWriter.WriteLine("</HTML>\r\n");


			
			oTxtStreamWriter.Close();
			oTxtFileStream.Close();
			//oTxtFileStream=null;
			//oTxtStreamWriter=null;

		}

		private void WriteText(string p_strTextFile,string p_strText)
		{
			System.IO.FileStream oTextFileStream;
			System.IO.StreamWriter oTextStreamWriter;

			if (!System.IO.File.Exists(p_strTextFile))
			{
				oTextFileStream = new System.IO.FileStream(p_strTextFile, System.IO.FileMode.Create, 
					System.IO.FileAccess.Write);
			}
			else
			{
				oTextFileStream = new System.IO.FileStream(p_strTextFile, System.IO.FileMode.Append, 
					System.IO.FileAccess.Write);
			}
			
			oTextStreamWriter = new System.IO.StreamWriter(oTextFileStream);
			oTextStreamWriter.Write(p_strText);
			oTextStreamWriter.Close();
			oTextFileStream.Close();
		}

		/// <summary>
		/// check and see if the user pressed the cancel button
		/// </summary>
		/// <param name="p_oLabel"></param>
		/// <returns></returns>
		private bool UserCancel(System.Windows.Forms.Label p_oLabel)
		{
			//System.Windows.Forms.Application.DoEvents();
			if (ReferenceUserControlScenarioRun.m_bUserCancel == true)
			{
				p_oLabel.ForeColor = System.Drawing.Color.Red;
				p_oLabel.Text = "Cancelled";
				return true;
			}
			return false;

		}
        private bool UserCancel(ProgressBarBasic.ProgressBarBasic p_oPb)
        {
            //System.Windows.Forms.Application.DoEvents();
            if (ReferenceUserControlScenarioRun.m_bUserCancel == true)
            {
                p_oPb.TextColor = Color.Red;
                frmMain.g_oDelegate.SetControlPropertyValue(p_oPb, "Text", "!!Cancelled!!");
                return true;
            }
            return false;

        }
		public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
		
		public FIA_Biosum_Manager.uc_optimizer_scenario_run ReferenceUserControlScenarioRun
		{
			get {return _uc_scenario_run;}
			set {_uc_scenario_run=value;}
		}

        private string getFileNamePrefix()
        {
            string strPrefix = "cycle_1";
            // Check the effective variables for a weighted variable
            foreach (string strPreVariable in ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.m_strPreVarArray)
            {
                string[] strPieces = strPreVariable.Split('.');
                if (strPieces.Length == 2 && !String.IsNullOrEmpty(strPieces[0]))
                {
                    if (strPieces[0].ToUpper().Contains("_WEIGHTED"))
                    {
                        strPrefix = "all_cycles";
                        break;
                    }
                }
            }
            // Check the optimization variable
            if (strPrefix == "cycle_1")
            {
                if (this.m_oOptimizationVariable.strOptimizedVariable == "Economic Attribute")
                {
                    // economic attributes are always weighted
                    strPrefix = "all_cycles";
                }
                else if (this.m_oOptimizationVariable.strOptimizedVariable == "Stand Attribute")
                {
                    string[] strPieces = this.m_oOptimizationVariable.strFVSVariableName.Split('.');
                    if (strPieces[0].ToUpper().Contains("_WEIGHTED"))
                    {
                        strPrefix = "all_cycles";
                    }
                }
            }
            // Check for a revenue filter (they are always weighted)
            if (strPrefix == "cycle_1")
            {
                if (this.m_oOptimizationVariable.bUseFilter == true)
                { strPrefix = "all_cycles"; }
            }
            // Check the tiebreaker filter
            if (strPrefix == "cycle_1")
            {
                FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection oTieBreakerCollection =
                    ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_tiebreaker1.m_oSavTieBreakerCollection;
                foreach (FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker.TieBreakerItem oItem in
                    oTieBreakerCollection)
                {
                    if (oItem.bSelected == true)
                    {
                        if (oItem.strMethod == "Economic Attribute")
                        {
                            // economic attributes are always weighted
                            strPrefix = "all_cycles";
                            break;
                        }
                        else if (oItem.strMethod == "Stand Attribute")
                        {
                            string[] strPieces = oItem.strFVSVariableName.Split('.');
                            if (strPieces[0].ToUpper().Contains("_WEIGHTED"))
                            {
                                strPrefix = "all_cycles";
                            }
                        }
                    }
                }
            }
            
            return strPrefix;
        }

        /// <summary>
        /// create table that lists all conditions in valid combos tables with psite information
        /// </summary>
        private void CondPsiteTable()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CondPsiteTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 2;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;

            intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                ReferenceUserControlScenarioRun.listViewEx1, "Create Condition - Processing Site Table");

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);


            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nPopulate cond_psite table\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "----------------------------------------\r\n");
            }

            DataMgr p_dataMgr = new DataMgr();

            using (System.Data.SQLite.SQLiteConnection workConn = new System.Data.SQLite.SQLiteConnection(m_strWorkTablesConn))
            {
                workConn.Open();

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strSystemResultsDbPathAndFile + "' AS results";
                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsCondPsiteTableName +
                " (BIOSUM_COND_ID,MERCH_PSITE_NUM,MERCH_PSITE_NAME,CHIP_PSITE_NUM,CHIP_PSITE_NAME) " +
                "SELECT a.biosum_cond_id, a.merch_haul_psite, a.merch_haul_psite_name, a.chip_haul_psite, a.chip_haul_psite_name " +
                "FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " AS a " +
                "JOIN validcombos AS b ON a.biosum_cond_id = b.biosum_cond_id " +
                "GROUP BY a.biosum_cond_id";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into cond_psite \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");

                p_dataMgr.SqlNonQuery(workConn, p_dataMgr.m_strSQL);
            }
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (p_dataMgr.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");

                this.m_intError = p_dataMgr.m_intError;

                return;
            }

            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");

            }
        }

        /// <summary>
        /// Populate context reference tables
        /// </summary>
        
        private void ContextReferenceTables()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//ContextReferenceTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                ReferenceUserControlScenarioRun.listViewEx1, "Populate Context Database");

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);


            DataMgr p_dataMgr = new DataMgr();

            using (System.Data.SQLite.SQLiteConnection contextConn = new System.Data.SQLite.SQLiteConnection(p_dataMgr.GetConnectionString(m_strContextDbPathAndFile)))
            {
                contextConn.Open();

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerDefinitions.DefaultSqliteDbFile + "' AS optimizer_defs";
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strProcessorResultsPathAndFile + "' AS processor_results";
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile + "' AS processor_rules";
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile + "' AS optimizer_rules";
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strHarvestMethodsPathAndFile + "' AS biosum_ref";
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + m_strPlotPathAndFile + "' AS master";
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nPopulate HARVEST_METHOD_REF table\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "----------------------------------------\r\n");
                }

                ProcessorScenarioItem.HarvestMethod oHarvestMethod = this.m_oProcessorScenarioItem.m_oHarvestMethod;
                string strRxHarvestMethod = "Y";
                if (!oHarvestMethod.SelectedHarvestMethod.Equals("RX"))
                {
                    strRxHarvestMethod = "N";
                }
                int intSteepSlopePct = -1;
                bool bSuccess = int.TryParse(oHarvestMethod.SteepSlopePercent, out intSteepSlopePct);

                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName +
                    " (rx, rx_harvest_method_low, rx_harvest_method_low_id, rx_harvest_method_low_category, rx_harvest_method_low_category_descr, " +
                    "rx_harvest_method_steep, use_rx_harvest_method_yn, steep_slope_pct) " +
                    "SELECT rx, HarvestMethodLowSlope As rx_harvest_method_low, m.HarvestMethodId AS rx_harvest_method_low_id, " +
                    "m.biosum_category AS rx_harvest_method_low_category, m.top_limb_slope_status AS rx_harvest_method_low_category_descr, " +
                    "HarvestMethodSteepSlope AS rx_harvest_method_sleep, '" + strRxHarvestMethod + "' AS use_rx_harvest_method_yn, " +
                    intSteepSlopePct + " AS steep_slope_pct FROM " + m_strRxTable + " AS r, " + m_strHarvestMethodsTable + " AS m " +
                    "WHERE TRIM(m.method) = TRIM(r.HarvestMethodLowSlope) AND steep_yn = 'N'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert rx harvest methods into HARVEST_METHOD_REF \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName + " AS ref " +
                    "SET rx_harvest_method_steep_id = HarvestMethodId, rx_harvest_method_steep_category = biosum_category, " +
                    "rx_harvest_method_steep_category_descr = top_limb_slope_status " +
                    "FROM " + m_strHarvestMethodsTable + " AS m " +
                    "WHERE TRIM(m.method) = TRIM(ref.rx_harvest_method_steep) AND steep_yn = 'Y'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate HARVEST_METHOD_REF with properties of rx steep slope methods\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                if (strRxHarvestMethod.Equals("N"))
                {
                    p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName +
                                     " SET scenario_harvest_method_low = '" + m_oProcessorScenarioItem.m_oHarvestMethod.HarvestMethodLowSlope.Trim() +
                                     "', scenario_harvest_method_steep = '" + m_oProcessorScenarioItem.m_oHarvestMethod.HarvestMethodSteepSlope.Trim() + "'";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate HARVEST_METHOD_REF for scenario harvest methods\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                    p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName + " AS ref " +
                        "SET scenario_harvest_method_low_id = HarvestMethodId, scenario_harvest_method_low_category = biosum_category, " +
                        "scenario_harvest_method_low_category_descr = top_limb_slope_status " +
                        "FROM " + m_strHarvestMethodsTable + " AS m " +
                        "WHERE TRIM(m.method) = TRIM(ref.scenario_harvest_method_low) AND steep_yn = 'N'";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate HARVEST_METHOD_REF with properties of scenario low slope methods\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                    p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName + " AS ref " +
                        "SET scenario_harvest_method_steep_id = HarvestMethodId, scenario_harvest_method_steep_category = biosum_category, " +
                        "scenario_harvest_method_steep_category_descr = top_limb_slope_status " +
                        "FROM " + m_strHarvestMethodsTable + " AS m " +
                        "WHERE TRIM(m.method) = TRIM(ref.scenario_harvest_method_steep) AND steep_yn = 'Y'";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate HARVEST_METHOD_REF with properties of scenario steep slope methods\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);
                }

                p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName +
                                " SET rx_harvest_method_low = 'NA', rx_harvest_method_low_id = 999, rx_harvest_method_low_category = 999, " +
                                "rx_harvest_method_low_category_descr = 'NA', " +
                                "rx_harvest_method_steep = 'NA', rx_harvest_method_steep_id = 999, rx_harvest_method_steep_category = 999, " +
                                "rx_harvest_method_steep_category_descr = 'NA', " +
                                "scenario_harvest_method_low = 'NA', scenario_harvest_method_low_id = 999, scenario_harvest_method_low_category = 999, " +
                                "scenario_harvest_method_low_category_descr = 'NA', " +
                                "scenario_harvest_method_steep = 'NA', scenario_harvest_method_steep_id = 999, scenario_harvest_method_steep_category = 999, " +
                                "scenario_harvest_method_steep_category_descr = 'NA' " +
                                "WHERE rx = '999'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate HARVEST_METHOD_REF for rx 999 (grow-only) \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsRxPackageRefTableName +
                                " (rxpackage, description, simyear1_rx, simyear2_rx, simyear3_rx, simyear4_rx) " +
                                "SELECT rxpackage, " + m_strRxPackageTable + ".description, simyear1_rx, simyear2_rx, simyear3_rx, simyear4_rx " +
                                "FROM " + m_strRxPackageTable;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert records into RXPACKAGE_REF \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                string[] arrRxCycle = new string[] { "simyear1_rx", "simyear2_rx", "simyear3_rx", "simyear4_rx" };
                foreach (string strRxCycle in arrRxCycle)
                {
                    p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsRxPackageRefTableName +
                        " SET " + strRxCycle + "_description = " + m_strRxTable + ".description " +
                        "FROM " + m_strRxTable +
                        " WHERE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsRxPackageRefTableName + "." + strRxCycle + " = " + m_strRxTable + ".rx";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate RXPACKAGE_REF for " + strRxCycle + " \r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();


                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.FVS.DefaultRxHarvestCostColumnsTableName + "_C" +
                                " SELECT * FROM " + Tables.FVS.DefaultRxHarvestCostColumnsTableName;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate rx_harvest_cost_columns_C table \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);


                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                foreach (ProcessorScenarioItem.TreeSpeciesAndDbhDollarValuesItem oItem in this.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection)
                {
                    bool bEnergyWood = Convert.ToBoolean(oItem.UseAsEnergyWood);
                    double dblChippedValue = Convert.ToDouble(oItem.ChipsDollarPerCubicFootValue);
                    string strEnergyWood = "Y";
                    if (bEnergyWood == false)
                    {
                        strEnergyWood = "N";
                        dblChippedValue = 0.0F;
                    }

                    p_dataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsDiameterSpeciesGroupRefTableName +
                                    " (dbh_class_num, dbh_range_inches, spp_grp_code, spp_grp, to_chips, merch_val_dpcf, value_if_chipped_dpgt)" +
                                    " VALUES (" + oItem.DiameterGroupId + ", '" + oItem.DbhGroup + "'," + oItem.SpeciesGroupId + ",'" + oItem.SpeciesGroup + "','" +
                                    strEnergyWood + "', " + oItem.MerchDollarPerCubicFootValue + ", " +
                                    dblChippedValue + ")";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate DIAMETER_SPP_GRP_REF_C table \r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsFvsWeightedVariablesRefTableName +
                    " SELECT variable_name, variable_description, baseline_rxpackage, variable_source, " +
                    "weight_1_pre, weight_1_post, weight_2_pre, weight_2_post, weight_3_pre, weight_3_post, weight_4_pre, weight_4_post " +
                    "FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + " AS o, " +
                    Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + " AS f " +
                    "WHERE f.calculated_variables_id = o.id";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate FVS_WEIGHTED_VARIABLES_REF table \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconWeightedVariablesRefTableName +
                    " (variable_name, variable_description, variable_source, cycle_1_weight) " +
                    "SELECT variable_name, variable_description, variable_source, weight AS cycle_1_weight " +
                    "FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + " AS o, " +
                    Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName + " AS e " +
                    "WHERE e.calculated_variables_id = o.id AND rxcycle = '1'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate ECON_WEIGHTED_VARIABLES_REF table \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                arrRxCycle = new string[] { "2", "3", "4" };

                foreach (string strRxCycle in arrRxCycle)
                {
                    string strFieldName = "cycle_" + strRxCycle + "_weight";

                    p_dataMgr.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconWeightedVariablesRefTableName + " AS ref " +
                        "SET " + strFieldName + " = e.weight " +
                        "FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + " AS o, " +
                        Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName + " AS e " +
                        "WHERE o.id = e.calculated_variables_id AND ref.variable_name = o.variable_name AND e.rxcycle = '" + strRxCycle + "'";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate " + strFieldName + " field \r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                    p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName + "_C" +
                            " SELECT * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName +
                            " WHERE scenario_id = '" + this.m_oProcessorScenarioItem.ScenarioId + "'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate SCENARIO_ADDITIONAL_ HARVEST_COSTS table \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsSpeciesGroupRefTableName +
                " SELECT species_group AS spp_grp_cd, common_name, spcd AS Ffia_spcd" +
                " FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName +
                " WHERE scenario_id = '" + this.m_oProcessorScenarioItem.ScenarioId + "'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate spp_grp_ref_C table \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + "_C " +
                    "SELECT * FROM " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate harvest_costs_C table \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + "_C " +
                    "SELECT * FROM " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate tree_vol_val_by_species_diam_groups_C table \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                p_dataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableName + "_C " +
                    "SELECT * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableName;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate scenario_psites_C table \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + p_dataMgr.m_strSQL + "\r\n");
                p_dataMgr.SqlNonQuery(contextConn, p_dataMgr.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            }

            if (p_dataMgr.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");

                this.m_intError = p_dataMgr.m_intError;

                return;
            }

            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;
            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
            }
        }

        public void ContextTextFiles(string strScenarioOutputFolder)
        {
            ProcessorScenarioTools oProcessorScenarioTools = new ProcessorScenarioTools();
            string strProperties = oProcessorScenarioTools.ScenarioProperties(this.m_oProcessorScenarioItem);
            string strPath = strScenarioOutputFolder + @"\db\processor_params_" + this.m_oProcessorScenarioItem.ScenarioId + ".txt";
            System.IO.File.WriteAllText(strPath, strProperties);

            OptimizerScenarioTools oOptimizerScenarioTools = new OptimizerScenarioTools();
            strProperties = oOptimizerScenarioTools.ScenarioProperties(this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0));
            strPath = strScenarioOutputFolder + @"\db\optimizer_params_" + this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).ScenarioId + ".txt";
            System.IO.File.WriteAllText(strPath, strProperties);
        }

	}

    public class ProductYields
    {
        string _strCondId = "";
        string _strRxPackage = "";
        double _dblChipYieldCfCycle1 = 0;
        double _dblMerchYieldCfCycle1 = 0;
        double _dblMaxNrDpaCycle1 = 0;
        double _dblHarvestOnsiteCpaCycle1 = 0;
        double _dblHaulMerchCpaCycle1 = 0;
        double _dblMerchChipNrDpaCycle1 = 0;
        double _dblHaulChipCpaCycle1 = 0;
        double _dblChipYieldCfCycle2 = 0;
        double _dblMerchYieldCfCycle2 = 0;
        double _dblMaxNrDpaCycle2 = 0;
        double _dblHarvestOnsiteCpaCycle2 = 0;
        double _dblHaulMerchCpaCycle2 = 0;
        double _dblMerchChipNrDpaCycle2 = 0;
        double _dblHaulChipCpaCycle2 = 0;
        double _dblChipYieldCfCycle3 = 0;
        double _dblMerchYieldCfCycle3 = 0;
        double _dblMaxNrDpaCycle3 = 0;
        double _dblHarvestOnsiteCpaCycle3 = 0;
        double _dblHaulMerchCpaCycle3 = 0;
        double _dblMerchChipNrDpaCycle3 = 0;
        double _dblHaulChipCpaCycle3 = 0;
        double _dblChipYieldCfCycle4 = 0;
        double _dblMerchYieldCfCycle4 = 0;
        double _dblMaxNrDpaCycle4 = 0;
        double _dblHarvestOnsiteCpaCycle4 = 0;
        double _dblHaulMerchCpaCycle4 = 0;
        double _dblMerchChipNrDpaCycle4 = 0;
        double _dblHaulChipCpaCycle4 = 0;

        public ProductYields(string strCondId, string strRxPackage)
        {
            _strCondId = strCondId;
            _strRxPackage = strRxPackage;
        }

        public void UpdateCycle1Yields(double dblChipYieldCf, double dblMerchYieldCf, double dblHarvestOnsiteCpa,
            double dblMaxNrDpaCycle, double dblHaulMerchCpa, double dblMerchChipNrDpa, double dblHaulChipCpa)
        {
            _dblChipYieldCfCycle1 = dblChipYieldCf;
            _dblMerchYieldCfCycle1 = dblMerchYieldCf;
            _dblHarvestOnsiteCpaCycle1 = dblHarvestOnsiteCpa;
            _dblMaxNrDpaCycle1 = dblMaxNrDpaCycle;
            _dblHaulMerchCpaCycle1 = dblHaulMerchCpa;
            _dblMerchChipNrDpaCycle1 = dblMerchChipNrDpa;
            _dblHaulChipCpaCycle1 = dblHaulChipCpa;
        }

        public void UpdateCycle2Yields(double dblChipYieldCf, double dblMerchYieldCf, double dblHarvestOnsiteCpa,
            double dblMaxNrDpaCycle, double dblHaulMerchCpa, double dblMerchChipNrDpa, double dblHaulChipCpa)
        {
            _dblChipYieldCfCycle2 = dblChipYieldCf;
            _dblMerchYieldCfCycle2 = dblMerchYieldCf;
            _dblHarvestOnsiteCpaCycle2 = dblHarvestOnsiteCpa;
            _dblMaxNrDpaCycle2 = dblMaxNrDpaCycle;
            _dblHaulMerchCpaCycle2 = dblHaulMerchCpa;
            _dblMerchChipNrDpaCycle2 = dblMerchChipNrDpa;
            _dblHaulChipCpaCycle2 = dblHaulChipCpa;
        }

        public void UpdateCycle3Yields(double dblChipYieldCf, double dblMerchYieldCf, double dblHarvestOnsiteCpa,
            double dblMaxNrDpaCycle, double dblHaulMerchCpa, double dblMerchChipNrDpa, double dblHaulChipCpa)
        {
            _dblChipYieldCfCycle3 = dblChipYieldCf;
            _dblMerchYieldCfCycle3 = dblMerchYieldCf;
            _dblHarvestOnsiteCpaCycle3 = dblHarvestOnsiteCpa;
            _dblMaxNrDpaCycle3 = dblMaxNrDpaCycle;
            _dblHaulMerchCpaCycle3 = dblHaulMerchCpa;
            _dblMerchChipNrDpaCycle3 = dblMerchChipNrDpa;
            _dblHaulChipCpaCycle3 = dblHaulChipCpa;
        }

        public void UpdateCycle4Yields(double dblChipYieldCf, double dblMerchYieldCf, double dblHarvestOnsiteCpa,
            double dblMaxNrDpaCycle, double dblHaulMerchCpa, double dblMerchChipNrDpa, double dblHaulChipCpa)
        {
            _dblChipYieldCfCycle4 = dblChipYieldCf;
            _dblMerchYieldCfCycle4 = dblMerchYieldCf;
            _dblHarvestOnsiteCpaCycle4 = dblHarvestOnsiteCpa;
            _dblMerchYieldCfCycle4 = dblMaxNrDpaCycle;
            _dblHaulMerchCpaCycle4 = dblHaulMerchCpa;
            _dblMerchChipNrDpaCycle4 = dblMerchChipNrDpa;
            _dblHaulChipCpaCycle4 = dblHaulChipCpa;
        }

        public string CondId()
        {
            return _strCondId;
        }
        public string RxPackage()
        {
            return _strRxPackage;
        }
        public double ChipYieldCfCycle1()
        {
            return _dblChipYieldCfCycle1;
        }
        public double ChipYieldCfCycle2()
        {
            return _dblChipYieldCfCycle2;
        }
        public double ChipYieldCfCycle3()
        {
            return _dblChipYieldCfCycle3;
        }
        public double ChipYieldCfCycle4()
        {
            return _dblChipYieldCfCycle4;
        }

        public double MerchYieldCfCycle1()
        {
            return _dblMerchYieldCfCycle1;
        }
        public double MerchYieldCfCycle2()
        {
            return _dblMerchYieldCfCycle2;
        }
        public double MerchYieldCfCycle3()
        {
            return _dblMerchYieldCfCycle3;
        }
        public double MerchYieldCfCycle4()
        {
            return _dblMerchYieldCfCycle4;
        }

        public double TotalYieldCfCycle1()
        {
            return _dblChipYieldCfCycle1 + _dblMerchYieldCfCycle1;
        }
        public double TotalYieldCfCycle2()
        {
            return _dblChipYieldCfCycle2 + _dblMerchYieldCfCycle2;
        }
        public double TotalYieldCfCycle3()
        {
            return _dblChipYieldCfCycle3 + _dblMerchYieldCfCycle3;
        }
        public double TotalYieldCfCycle4()
        {
            return _dblChipYieldCfCycle4 + _dblMerchYieldCfCycle4;
        }

        public double HarvestOnsiteCpaCycle1()
        {
            return _dblHarvestOnsiteCpaCycle1;
        }
        public double HarvestOnsiteCpaCycle2()
        {
            return _dblHarvestOnsiteCpaCycle2;
        }
        public double HarvestOnsiteCpaCycle3()
        {
            return _dblHarvestOnsiteCpaCycle3;
        }
        public double HarvestOnsiteCpaCycle4()
        {
            return _dblHarvestOnsiteCpaCycle4;
        }

        public double MaxNrDpaCycle1()
        {
            return _dblMaxNrDpaCycle1;
        }
        public double MaxNrDpaCycle2()
        {
            return _dblMaxNrDpaCycle2;
        }
        public double MaxNrDpaCycle3()
        {
            return _dblMaxNrDpaCycle3;
        }
        public double MaxNrDpaCycle4()
        {
            return _dblMaxNrDpaCycle4;
        }

        public double TreatmentHaulCostsCycle1()
        {
            double dblAddedChipCost = 0;
            if (_dblMerchChipNrDpaCycle1 >= _dblMaxNrDpaCycle1)
            {
                dblAddedChipCost = _dblHaulChipCpaCycle1;
            }
            return _dblHarvestOnsiteCpaCycle1 + _dblHaulMerchCpaCycle1 + dblAddedChipCost;
        }
        public double TreatmentHaulCostsCycle2()
        {
            double dblAddedChipCost = 0;
            if (_dblMerchChipNrDpaCycle2 >= _dblMaxNrDpaCycle2)
            {
                dblAddedChipCost = _dblHaulChipCpaCycle2;
            }
            return _dblHarvestOnsiteCpaCycle2 + _dblHaulMerchCpaCycle2 + dblAddedChipCost;
        }
        public double TreatmentHaulCostsCycle3()
        {
            double dblAddedChipCost = 0;
            if (_dblMerchChipNrDpaCycle3 >= _dblMaxNrDpaCycle3)
            {
                dblAddedChipCost = _dblHaulChipCpaCycle3;
            }
            return _dblHarvestOnsiteCpaCycle3 + _dblHaulMerchCpaCycle3 + dblAddedChipCost;
        }
        public double TreatmentHaulCostsCycle4()
        {
            double dblAddedChipCost = 0;
            if (_dblMerchChipNrDpaCycle4 >= _dblMaxNrDpaCycle4)
            {
                dblAddedChipCost = _dblHaulChipCpaCycle4;
            }
            return _dblHarvestOnsiteCpaCycle4 + _dblHaulMerchCpaCycle4 + dblAddedChipCost;
        }

    }



}

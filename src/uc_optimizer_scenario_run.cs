using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Text;
using System.Data.OleDb;

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
		public System.Windows.Forms.Button btnAccess;
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
            this.AddListViewRowItem("Populate Context Database", true, true);
            //
            //
            //
            this.AddListViewRowItem("Populate FVS PRE-POST Context Database", true, false);
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
            this.btnAccess = new System.Windows.Forms.Button();
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
            this.panel1.Controls.Add(this.btnAccess);
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
            // btnAccess
            // 
            this.btnAccess.Enabled = false;
            this.btnAccess.ForeColor = System.Drawing.Color.Black;
            this.btnAccess.Location = new System.Drawing.Point(349, 8);
            this.btnAccess.Name = "btnAccess";
            this.btnAccess.Size = new System.Drawing.Size(120, 30);
            this.btnAccess.TabIndex = 33;
            this.btnAccess.Text = "Microsoft Access";
            this.btnAccess.Click += new System.EventHandler(this.btnAccess_Click);
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
			
			((frmOptimizerScenario)this.ParentForm).Height = 0 ; //((frmScenario)this.ParentForm).grpboxMenu.Height * 2;
		}

		private void btnAccess_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process proc = new System.Diagnostics.Process();
			proc.StartInfo.UseShellExecute = true;
			try
			{
                proc.StartInfo.FileName = ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile;		
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

		private void btnViewAuditTables_Click(object sender, System.EventArgs e)
		{
            viewAuditTables();	
           

		}
        private void viewAuditTables()
        {
            string strConn = "";
            string strTable = "";
            int x;
            ado_data_access oAdo = new ado_data_access();
            this.m_frmGridView = new frmGridView();
            this.m_frmGridView.Text = "Treatment Optimizer: Audit";
            lblMsg.Text = "";
            lblMsg.Show();
            string strAccdbPathAndFile = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\" + Tables.Audit.DefaultCondAuditTableDbFile;
            if (System.IO.File.Exists(strAccdbPathAndFile) == true)
            {
                oAdo.OpenConnection(oAdo.getMDBConnString(strAccdbPathAndFile, "", ""));
                if (oAdo.TableExist(oAdo.m_OleDbConnection, Tables.Audit.DefaultCondAuditTableName) == true)
                {
                    strTable = Tables.Audit.DefaultCondAuditTableName;
                    strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" +
                        strAccdbPathAndFile + ";User Id=admin;Password=;";

                    this.lblMsg.Text = strTable;
                    this.lblMsg.Refresh();

                    this.m_frmGridView.LoadDataSet(strConn, "select * from " + strTable + " " + strTable, strTable);
                         
                }
                if (oAdo.TableExist(oAdo.m_OleDbConnection, Tables.Audit.DefaultCondRxAuditTableName) == true)
                {
                    strTable = Tables.Audit.DefaultCondRxAuditTableName;
                    strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" +
                               strAccdbPathAndFile + ";User Id=admin;Password=;";

                    this.lblMsg.Text = strTable;
                    this.lblMsg.Refresh();

                    this.m_frmGridView.LoadDataSet(strConn, "select * from " + strTable + " " + strTable, strTable);
                }
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
            }
            this.m_frmGridView.TileGridViews();
            this.m_frmGridView.Show();
            this.m_frmGridView.Focus();
            lblMsg.Text = "";
            lblMsg.Refresh();
            lblMsg.Hide();
            oAdo = null;

            
        }

        private void btnViewScenarioTables_Click(object sender, System.EventArgs e)
		{
			this.viewResultsTables();
		}

		/// <summary>
		/// every optimizer_results.accdb table is viewed in a uc_gridview control
		/// </summary>
		private void viewResultsTables()
		{
			string strMDBPathAndFile="";
			string strConn="";
			string strSQL="";
			string[] strTableNames;
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();


            strMDBPathAndFile = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile;

			int intCount = p_dao.getTableNames(strMDBPathAndFile,ref strTableNames);
			if (p_dao.m_intError==0)
			{
				if (intCount > 0)
				{
					
					this.lblMsg.Text = "";
					this.lblMsg.Visible=true;
					
						strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + 
							strMDBPathAndFile + ";User Id=admin;Password=;";
					
					this.m_frmGridView = new frmGridView();
                    this.m_frmGridView.Text = "Treatment Optimizer: Run Scenario Results (" + this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim() + ")";
					for (int x=0; x <= intCount-1;x++)
					{
						this.lblMsg.Text = strTableNames[x];
						this.lblMsg.Refresh();
						strSQL = "select * from " + strTableNames[x].Trim();
						this.m_frmGridView.LoadDataSet(strConn,strSQL,strTableNames[x].Trim());
						

					}
					
					this.lblMsg.Text="";
					this.lblMsg.Visible=false;
					if (intCount > 1) this.m_frmGridView.TileGridViews();
					this.m_frmGridView.Show();
					this.m_frmGridView.Focus();


				}
				else
				{
					MessageBox.Show("No Tables Found In " + strMDBPathAndFile);
				}

			}
			p_dao = null;
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
            this.btnAccess.Enabled = false;

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
				this.btnAccess.Enabled=false;
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
                this.btnAccess.Enabled = true;
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
		private void StopThread()
		{
			this.m_thread.Suspend();
            DialogResult result =  MessageBox.Show("Cancel Running The Scenario (Y/N)?","Cancel Process", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			switch (result) 
			{
				case DialogResult.Yes:
					this.m_bAbortThread=true;
					this.m_thread.Resume();
					this.m_thread.Abort();
					while (this.m_thread.IsAlive)
					{
						this.lblMsg.Text = "Attempting To Abort Process...Stand By";
						this.m_thread.Join(1000);
						if (this.m_thread.IsAlive) this.m_thread.Abort();
						System.Windows.Forms.Application.DoEvents();
					}
					return;
				case DialogResult.No:
					this.m_thread.Resume();
					return;
			}                
			
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
				this.m_intError= ReferenceOptimizerScenarioForm.uc_scenario_filter1.Val_PlotFilter(ReferenceOptimizerScenarioForm.uc_scenario_filter1.txtCurrentSQL.Text.Trim());
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

                /***************************************************************************
                     **make sure all the scenario datasource tables and files are available
                     **and ready for use
                     ***************************************************************************/
                UpdateThermPercent();
                if (ReferenceOptimizerScenarioForm.m_ldatasourcefirsttime == true)
                {
                    ReferenceOptimizerScenarioForm.uc_datasource1.populate_listview_grid();
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



        private void UpdateOptimizationGroupboxText()
        {
            OptimizerScenarioItem.OptimizationVariableItem_Collection oOptimizationVariableCollection = this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection;

            for (int x = 0; x <= oOptimizationVariableCollection.Count - 1; x++)
            {
                if (oOptimizationVariableCollection.Item(x).bSelected)
                {
                   
                    break;
                }
            }
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
		private string m_strSQL;
		private string m_strConn;
		public string m_strTempMDBFile;
		public ado_data_access m_ado;
        private dao_data_access m_oDao;
		public System.Data.OleDb.OleDbConnection m_TempMDBFileConn;
		public string m_strSystemResultsDbPathAndFile="";
        public string m_strContextDbPathAndFile = "";
        public string m_strFvsContextDbPathAndFile = "";
        public string m_strFVSPreValidComboDbPathAndFile = "";
        public string m_strFVSPostValidComboDbPathAndFile = "";
        
		private env m_oEnv;
		private utils m_oUtils;
		public string m_strPlotTable;
        public string m_strPlotPathAndFile;
		public string m_strRxTable;
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
        ODBCMgr odbcMgr = new ODBCMgr();
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
			
			this.m_intError=0;
			_uc_scenario_run = p_form;
            this.ReferenceOptimizerScenarioForm = p_form.ReferenceOptimizerScenarioForm;
			try
			{
                m_strDebugFile = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\runlog.txt";


                if (frmMain.g_bDebug)
                {
                    if (System.IO.File.Exists(m_strDebugFile)) System.IO.File.Delete(m_strDebugFile);
                    this.m_strLine = "START: Optimizer Run Log " + System.DateTime.Now.ToString();
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
				this.m_intError=-1;
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
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic =(ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, 2);
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = 2;
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 19;
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps=1;
                FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep=1;

                frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, 2);
                frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, 2, "Selected", true);
                frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, 2, "focused", true);

                /**************************************************************************
                 **first lets create a temp mdb with links to all the scenario Optimizer 
                 **and result tables
                 **************************************************************************/
                this.m_oUtils=new utils();
				this.m_oEnv = new env();

                //get the selected processor scenario item
                if (this.ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_processor_scenario_select1.m_oProcessorScenarioItem != null)
                    this.m_oProcessorScenarioItem = this.ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_processor_scenario_select1.m_oProcessorScenarioItem;

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                //load the treatment packages
                this.m_oRxTools.LoadAllRxPackageItems(m_oRxPackageItem_Collection);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                
				this.m_oVarSub.ReferenceSQLMacroSubstitutionVariableCollection = 
					frmMain.g_oSQLMacroSubstitutionVariable_Collection;

				string strScenarioOutputFolder = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim();
                this.m_strSystemResultsDbPathAndFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");
				this.CopyScenarioResultsTable(this.m_strSystemResultsDbPathAndFile,strScenarioOutputFolder + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile);

                this.m_strContextDbPathAndFile = "";
                intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                    ReferenceUserControlScenarioRun.listViewEx1, "Populate Context Database");
                oCheckBox = (CheckBox) ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(0, intListViewIndex);
                if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)oCheckBox, "Checked", false) == true)
                {
                    this.m_strContextDbPathAndFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
                    this.CopyScenarioResultsTable(this.m_strContextDbPathAndFile, strScenarioOutputFolder + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsContextDbFile);
                }

                this.m_strFvsContextDbPathAndFile = "";
                intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                    ReferenceUserControlScenarioRun.listViewEx1, "Populate FVS PRE-POST Context Database");
                oCheckBox = (CheckBox)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(0, intListViewIndex);
                if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)oCheckBox, "Checked", false) == true)
                {
                    this.m_strFvsContextDbPathAndFile = strScenarioOutputFolder + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsFvsContextDbFile;
                }
                
                this.m_strFVSPreValidComboDbPathAndFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "mdb");
                this.CopyScenarioResultsTable(this.m_strFVSPreValidComboDbPathAndFile, ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\validcombo.mdb");

                this.m_strFVSPostValidComboDbPathAndFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "mdb");
                this.CopyScenarioResultsTable(this.m_strFVSPostValidComboDbPathAndFile, ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\validcombo.mdb");


                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                string[] arr1 = new string[] { this.m_oEnv.strTempDir };
                object oValue = frmMain.g_oDelegate.GetValueExecuteControlMethodWithParam(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_datasource1,
                    "CreateMDBAndScenarioTableDataSourceLinks", arr1, true);
                if (oValue != null)
                {
                    string strValue = Convert.ToString(oValue);
                    if (strValue != "false")
                    {
                        this.m_strTempMDBFile = strValue;
                    }
                }
                
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

				this.m_strUserDefinedPlotSQL= 
					this.m_oVarSub.SQLTranslateVariableSubstitution(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_filter1.txtCurrentSQL.Text.Trim());

				this.m_strUserDefinedCondSQL= 
					this.m_oVarSub.SQLTranslateVariableSubstitution(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_cond_filter1.txtCurrentSQL.Text.Trim());

            
				if (this.m_strTempMDBFile.Trim().Length == 0)
				{
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!RunOptimizer: Error Creating MDB File Containing Links To All The Tables!!\r\n");
                    MessageBox.Show("RunOptimizer: Error Creating MDB File Containing Links To All The Optimizer Tables");
				}
				else
				{

                    if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.UserControl)_uc_scenario_run.uc_filesize_monitor1, "Visible", false) == false)
                    {
                        _uc_scenario_run.uc_filesize_monitor1.BeginMonitoringFile(
                            m_strTempMDBFile, 2000000000, "2GB");
                        _uc_scenario_run.uc_filesize_monitor1.Information = "Work table containing table links";

                        _uc_scenario_run.uc_filesize_monitor2.BeginMonitoringFile(
                            m_strSystemResultsDbPathAndFile, 2000000000, "2GB");
                        _uc_scenario_run.uc_filesize_monitor2.Information = "Scenario results DB file";

                    }

                    m_oDao = new dao_data_access();
                    //compact the optimizer_results.accdb file

                    CompactMDB(m_strSystemResultsDbPathAndFile, null);

                    //28-APR-2021
                    //This method should be obsolete; Commented out
     //               if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
     //                   frmMain.g_oUtils.WriteText(m_strDebugFile,"--Delete Scenario Result Table Records--\r\n");
					//this.DeleteScenarioResultRecords();

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile,"Links MDB File: " + this.m_strTempMDBFile + "\r\n");
					ReferenceUserControlScenarioRun.btnAccess.Enabled=false;

					getTableNames();
                    if (this.m_intError != 0)
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
                    OptimizerScenarioItem.OptimizationVariableItem_Collection oOptimizationVariableCollection = this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection;
	
					
                    for (int x=0;x<=oOptimizationVariableCollection.Count-1;x++)
					{
						if (oOptimizationVariableCollection.Item(x).bSelected)
						{
							m_oOptimizationVariable.Copy(oOptimizationVariableCollection.Item(x),ref m_oOptimizationVariable);
							break;
						}
					}

					if (m_oOptimizationVariable.strMaxYN=="Y")
					{
						this.m_strOptimizationTableName = "Max";
						this.m_strOptimizationAggregateSql="MAX";
						this.m_strOptimizationAggregateColumnName="max_optimization_value";
					}
					else
					{
						this.m_strOptimizationTableName = "Min";
						this.m_strOptimizationAggregateSql="MIN";
						this.m_strOptimizationAggregateColumnName="min_optimization_value";
					}

					if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "REVENUE")
					{
						this.m_strOptimizationTableName = this.m_strOptimizationTableName + "NR";
						this.m_strOptimizationSourceTableName=Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;
                        this.m_strOptimizationTableNameSql = "cycle1_effective_" + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;
						this.m_strOptimizationSourceColumnName="max_nr_dpa";
						this.m_strOptimizationColumnNameSql="post_variable_value";
							
					}
					else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "MERCHANTABLE VOLUME")
					{
                        this.m_strOptimizationSourceTableName = Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;
                        this.m_strOptimizationTableNameSql = "cycle1_effective_" + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;
						this.m_strOptimizationColumnNameSql="post_variable_value";
						this.m_strOptimizationSourceColumnName= "merch_vol_cf";
						this.m_strOptimizationTableName = this.m_strOptimizationTableName + "MerchVol";
						if (this.m_oOptimizationVariable.bUseFilter) this.m_strOptimizationTableName = this.m_strOptimizationTableName + "NR";
					}
                    else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "ECONOMIC ATTRIBUTE")
                    {
                        //This is used to name the output .accdb
                        this.m_strOptimizationColumnNameSql = "post_variable_value";
                        this.m_strOptimizationTableName = this.m_oOptimizationVariable.strFVSVariableName;
                        if (this.m_oOptimizationVariable.bUseFilter) this.m_strOptimizationTableName = this.m_strOptimizationTableName + "NR";
                    }
					else
					{
						
						string[] strCol = frmMain.g_oUtils.ConvertListToArray(this.m_oOptimizationVariable.strFVSVariableName,".");
						this.m_strOptimizationSourceTableName=strCol[0];
						this.m_strOptimizationSourceColumnName = strCol[1];
						this.m_strOptimizationTableNameSql="cycle1_effective_optimization_treatments";
						if (this.m_oOptimizationVariable.strValueSource=="POST")
						{
							this.m_strOptimizationColumnNameSql="post_variable_value";
						}
						else if (this.m_oOptimizationVariable.strValueSource=="POST-PRE")
						{
							this.m_strOptimizationColumnNameSql="change_value";
						}


						this.m_strOptimizationTableName = this.m_strOptimizationTableName + strCol[1].Trim();
						if (this.m_oOptimizationVariable.bUseFilter) this.m_strOptimizationTableName = this.m_strOptimizationTableName + "NR";
					}
					//revenue filter
					if (this.m_oOptimizationVariable.bUseFilter)
					{
						switch (this.m_oOptimizationVariable.strFilterOperator)
						{
							case "<":
								this.m_strOptimizationTableName = this.m_strOptimizationTableName + "lt";
								break;
							case ">":
								this.m_strOptimizationTableName = this.m_strOptimizationTableName + "gt";
								break;
							case "<=":
								this.m_strOptimizationTableName = this.m_strOptimizationTableName + "lte";
								break;
							case ">=":
								this.m_strOptimizationTableName = this.m_strOptimizationTableName + "gte";
								break;
							case "<>":
								this.m_strOptimizationTableName = this.m_strOptimizationTableName + "ne";
								break;
							case "=":
								this.m_strOptimizationTableName = this.m_strOptimizationTableName + "e";
								break;
						}
                        if (m_oOptimizationVariable.dblFilterValue>=0)
						    this.m_strOptimizationTableName = this.m_strOptimizationTableName + "P" + Convert.ToString(this.m_oOptimizationVariable.dblFilterValue).Trim();
                        else
                            this.m_strOptimizationTableName = this.m_strOptimizationTableName + "N" + Convert.ToString(this.m_oOptimizationVariable.dblFilterValue*-1).Trim();
					}

                    this.ReferenceOptimizerScenarioForm.OutputTablePrefix = this.getFileNamePrefix();
                    m_strOptimizationTableName = ReferenceOptimizerScenarioForm.OutputTablePrefix + "_" + m_strOptimizationTableName;

                    CreateAuditTables();
					CreateOptimizerResultTables();
                    if (!String.IsNullOrEmpty(m_strContextDbPathAndFile))
                        CreateContextTables();
                    CreateValidComboTables();


                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

					//CREATE TABLE LINKS
					CreateScenarioResultTableLinks();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    if (! String.IsNullOrEmpty(m_strContextDbPathAndFile))
                    {
                        this.CreateContextTableLinks();
                        if (this.m_intError != 0)
                        {
                            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                            return;
                        }
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    //CREATE PROCESSOR SCENARIO TABLE LINKS
                    CreateProcessorScenarioResultTableLinks();

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

					this.CreateAuditTableLinks();
					if (this.m_intError != 0) return;

					this.CreateScenarioTableLinks();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }
                    this.CreateValidComboTableLinks();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }
                    

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    m_oRxTools.CreateTableLinksToFVSPrePostTables(m_strTempMDBFile);
                    m_intError = m_oRxTools.m_intError;
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

					//CREATE WORK TABLES
					CreateTableStructureOfHarvestCosts();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

					CreateTableStructureForIntensity();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
					
                    this.CreateTableStructureForScenarioProcessingSites();

                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
					
					this.CreateTableStructureForHaulCosts();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

					this.CreateTableStructureForPlotCondAccessiblity();
                    if (this.m_intError != 0)
                    {
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                        return;
                    }


                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    

					this.m_ado = new ado_data_access();

					this.m_strConn=m_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");


			
					this.m_TempMDBFileConn = new System.Data.OleDb.OleDbConnection();
					this.m_ado.OpenConnection(this.m_strConn,ref this.m_TempMDBFileConn);
					if (this.m_ado.m_intError==0)
					{
						/******************************************************************
						 **create table structure of user defined sql plot filter statement
						 ******************************************************************/
						CreateTableStructureOfUserDefinedPlotSQL();

                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

						/********************************************************************
						 **create table structure for condition table filters
						 ********************************************************************/
						this.CreateTableStructureForUserDefinedConditionTable();

                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

						/********************************************************************
						 **filter scenario selected processing sites
						 ********************************************************************/
						FilterPSites();

                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

						


						//close the connection if it is open
						if (this.m_TempMDBFileConn.State == System.Data.ConnectionState.Open)
						{
							this.m_TempMDBFileConn.Close();
						}
						if (this.m_intError==0)
						{
							this.m_strConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
							this.m_TempMDBFileConn = new System.Data.OleDb.OleDbConnection();
							this.m_ado.OpenConnection(this.m_strConn,ref this.m_TempMDBFileConn);
							this.m_intError=this.m_ado.m_intError;
						}

                        if (this.m_intError == 0)
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
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);

						if (this.m_intError==0) 
						{
                           
                           
							this.CondAccessible();
				
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
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);

						if (this.m_intError == 0)
						{

                           

							this.getHaulCosts();

						}
						else
						{
							if (ReferenceUserControlScenarioRun.m_bUserCancel==false)
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
                        FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
                        frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);

						if (this.m_intError == 0) //&& (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)oCheckBox,"Checked",false)==true && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.sumTreeVolVal();

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
                       
                        
						if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.validcombos();

						}
                        /***************************************************************************
                         **cond psite table
                         ***************************************************************************/

                        if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                        {
                            this.CondPsiteTable();

                        }

                        /***************************************************************************
                         **context reference database
                         ***************************************************************************/

                        if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false &&
                            !String.IsNullOrEmpty(m_strContextDbPathAndFile))
                        {
                            this.ContextReferenceTables();
                            this.ContextTextFiles(strScenarioOutputFolder);
                        }

                        /***************************************************************************
                         **FVS context reference database
                         ***************************************************************************/

                        if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false &&
                            !String.IsNullOrEmpty(this.m_strFvsContextDbPathAndFile))
                        {
                            this.FvsContextReferenceTables();

                        }

						/**********************************************************************
						 **wood product yields net revenue and costs summary by treatment table
						 **********************************************************************/
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
                            this.econ_by_rx_cycle();

						}

                        /*******************************************************************************
						 **wood product yields net revenue and costs summary by treatment package table 
						 *******************************************************************************/
                        if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                        {
                            this.econ_by_rx_utilized_sum();

                        }
                        //compact
                        if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
                            CompactMDB(m_strSystemResultsDbPathAndFile,m_TempMDBFileConn);
                        }

                        /**********************************************************************
                         **Calculate custom economic variables if needed
                         **********************************************************************/

                        if (this.m_intError == 0 && ReferenceUserControlScenarioRun.m_bUserCancel == false)
                        {
                            this.calculate_weighted_econ_variables();

                        }

						/***************************************************************************
						 **effective treatments
						 ***************************************************************************/
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.Effective();

						}
						/**************************************************************************
						 **optimization
						 **************************************************************************/
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.Optimization();
						}
						/**************************************************************************
						 **tie breakers
						 **************************************************************************/
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.tiebreaker();
						}
						
						/*********************************************************************
						 **find the best treatments for revenue, torch/crown index improvement,
						 **and merch removal
						 *********************************************************************/ 
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
							this.Best_rx_summary();
						}

 
                        if (m_TempMDBFileConn != null) m_ado.CloseConnection(m_TempMDBFileConn);
                        m_TempMDBFileConn = null;

                        this.CreateReferenceResultsTableLinks();
                       
						if (this.m_intError==0 && ReferenceUserControlScenarioRun.m_bUserCancel==false)
						{
                            CompactMDB(m_strSystemResultsDbPathAndFile, null);
							System.DateTime oDate = System.DateTime.Now;
							string strDateFormat = "yyyy-MM-dd_HH-mm";
							string strFileDate = oDate.ToString(strDateFormat);
							strFileDate = strFileDate.Replace("/","_"); strFileDate=strFileDate.Replace(":","_");
							this.CreateHtml();
                            this.CopyScenarioResultsTable(strScenarioOutputFolder + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile, this.m_strSystemResultsDbPathAndFile);
                            // issue #255: Optimizer results requires repair
                            if (frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Compacting " + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile + " \r\n");
                            }
                            CompactMDB(strScenarioOutputFolder + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile, null);
                            if (! String.IsNullOrEmpty(this.m_strContextDbPathAndFile))
                                this.CopyScenarioResultsTable(strScenarioOutputFolder + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsContextDbFile, this.m_strContextDbPathAndFile);
                            this.m_strSystemResultsDbPathAndFile = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile;
                            this.m_strFVSPreValidComboDbPathAndFile = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\validcombo_fvspre.mdb";
                            
							this.CopyScenarioResultsTable(
								ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\optimizer_results_" + this.m_strOptimizationTableName + "_" + strFileDate.Trim() + ".accdb",
                                ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile);


						}

						

                        if (frmMain.g_bDebug)
                        {
                            this.m_strLine = "END: Optimizer Analysis Run Log " + System.DateTime.Now.ToString();
                            frmMain.g_oUtils.WriteText(m_strDebugFile, m_strLine + "\r\n\r\n");
                        }

						
						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Visible",false);
						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnCancel,"Text","Start");
						frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnViewResultsTables,"Enabled",true);
						if (this.m_intError == 0) frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnAccess,"Enabled",true);

                        

						if (this.m_intError != 0)
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
				}
				
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.btnViewLog,"Enabled",true);

                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
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
                if (m_TempMDBFileConn != null) m_ado.CloseConnection(m_TempMDBFileConn);
                m_TempMDBFileConn = null;
			}
            _uc_scenario_run.uc_filesize_monitor1.EndMonitoringFile();
            _uc_scenario_run.uc_filesize_monitor2.EndMonitoringFile();
            if ((bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.UserControl)_uc_scenario_run.uc_filesize_monitor3,"Visible",false))
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
            dao_data_access oDao = new dao_data_access();
            //
            //create an audit DB file in the scenario output directory
            //
            string strAccdbPathAndFile = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\" + Tables.Audit.DefaultCondAuditTableDbFile;
            if (System.IO.File.Exists(strAccdbPathAndFile) == false)
            {
                oDao.CreateMDB(strAccdbPathAndFile);
            }
            ado_data_access oAdo = new ado_data_access();
            string strConn = oAdo.getMDBConnString(strAccdbPathAndFile, "", "");
            using (var oConn = new OleDbConnection(strConn))
            {
                oConn.Open();
                if (oAdo.TableExist(oConn, Tables.Audit.DefaultCondAuditTableName) == false)
                {
                    frmMain.g_oTables.m_oAudit.CreateCondAuditTable(oAdo, oConn, Tables.Audit.DefaultCondAuditTableName);
                }
                if (oAdo.TableExist(oConn, Tables.Audit.DefaultCondRxAuditTableName) == false)
                {
                    frmMain.g_oTables.m_oAudit.CreatePlotCondRxAuditTable(oAdo, oConn, Tables.Audit.DefaultCondRxAuditTableName);
                }
            }
            oDao = null;
            oAdo = null;

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
			
			
			/********************************************************
			 **get the user defined PLOT filter sql
			 ********************************************************/
			this.m_strSQL = this.m_strUserDefinedPlotSQL;
			/****************************************************************
			 **get the table structure that results from executing the sql
			 ****************************************************************/
			System.Data.DataTable p_dt = this.m_ado.getTableSchema(this.m_TempMDBFileConn,this.m_strSQL);

			/*****************************************************************
			 **create the table structure in the temporary accdb file
			 **and give it the name of userdefinedplotfilter_work
			 *****************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Create userdefinedplotfilter_work Table Schema From User Defined Plot Filter SQL\r\n");
			dao_data_access p_dao = new dao_data_access();
			p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"userdefinedplotfilter_work",p_dt,true);
			p_dt.Dispose();
			this.m_ado.m_OleDbDataReader.Close();
			if (p_dao.m_intError!=0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!! Error Creating Table Schema!!\r\n");
				this.m_intError=p_dao.m_intError;
				p_dao=null;
				return;
			}

			/***********************************************************************
			 **make a copy of the userdefinedplot filter table and give it the
			 **name ruledefinitionsplotfilter. This will apply the owngrpcd
			 **filters and any other future filters to the userdefinedplotfilter_work table.
			 ***********************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Delete table ruledefinitionsplotfilter\r\n");
			p_dao.DeleteTableFromMDB(this.m_strSystemResultsDbPathAndFile,"ruledefinitionsplotfilter");
			if (p_dao.m_intError !=0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!! Error Deleting ruledefinitionsplotfilter Table!!\r\n");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Copy table structure userdefinedplotfilter_work to ruledefinitionsplotfilter\r\n");
            p_dao.MoveTableToMDB(this.m_strSystemResultsDbPathAndFile, "ruledefinitionsplotfilter", this.m_strTempMDBFile, "userdefinedplotfilter_work", false);
			if (p_dao.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!! Error Creating ruledefinitionsplotfilter Table!!\r\n");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
			/***********************************************************************
			 **create a table link to the ruledefinitionsplotfilter table in the
			 **temporary MDB file located on the hard drive of the user
			 ***********************************************************************/
			p_dao.CreateTableLink(this.m_strTempMDBFile,"ruledefinitionsplotfilter",m_strSystemResultsDbPathAndFile,"ruledefinitionsplotfilter",true);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Create link in " + this.m_strTempMDBFile + "\r\n");

                frmMain.g_oUtils.WriteText(m_strDebugFile,m_strSystemResultsDbPathAndFile + "\truledefinitionsplotfilter\r\n");
            }


			p_dao=null;




		}
		/// <summary>
		/// create a table structure that will hold
		/// the plot data that results when running the user 
		/// defined sql
		/// </summary>
		/// <param name="strUserDefinedSQL"></param>
		private void CreateTableStructureOfUserDefinedSQLOld(string strUserDefinedSQL)
		{
			
			
			/*******************************************************************
			 ** get optimizer_results.accdb path
			 *******************************************************************/
            string strMDBPathAndFile = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile;
			/********************************************************
			 **get the user defined PLOT filter sql
			 ********************************************************/
			this.m_strSQL = this.m_strUserDefinedPlotSQL;
			/****************************************************************
			 **get the table structure that results from executing the sql
			 ****************************************************************/
			System.Data.DataTable p_dt = this.m_ado.getTableSchema(this.m_TempMDBFileConn,this.m_strSQL);

			/*****************************************************************
			 **create the table structure in the optimizer_results.accdb file
			 **and give it the name of userdefinedplotfilter
			 *****************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Create userdefinedplotfilter Table Schema From User Defined Plot Filter SQL\r\n");
			dao_data_access p_dao = new dao_data_access();
            p_dao.CreateMDBTableFromDataSetTable(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile, "userdefinedplotfilter", p_dt, true);
			p_dt.Dispose();
			this.m_ado.m_OleDbDataReader.Close();
			if (p_dao.m_intError!=0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!! Error Creating Table Schema!!\r\n");
				this.m_intError=p_dao.m_intError;
				p_dao=null;
				return;
			}
			/***********************************************************************
			 **create a table link to the user defined plot filter table in the
			 **temporary MDB file located on the hard drive of the user
			 ***********************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Create link in " + this.m_strTempMDBFile + "\r\n");
			p_dao.CreateTableLink(this.m_strTempMDBFile,"userdefinedplotfilter",strMDBPathAndFile,"userdefinedplotfilter",true);
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,strMDBPathAndFile + "\tuserdefinedplotfilter\r\n");

			/***********************************************************************
			 **make a copy of the userdefinedplot filter table and give it the
			 **name ruledefinitionsplotfilter. This will apply the owngrpcd
			 **filters and any other future filters to the userdefinedplotfilter table.
			 ***********************************************************************/
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Delete table ruledefinitionsplotfilter\r\n");
			//first delete the table if it exists already
             p_dao.DeleteTableFromMDB(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile, "ruledefinitionsplotfilter");
			if (p_dao.m_intError !=0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!! Error Deleting alluserdefinedplotfilter Table!!\r\n");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Copy table structure userdefinedplotfilter to ruledefinitionsplotfilter\r\n");
             p_dao.MoveTableToMDB(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile, "ruledefinitionsplotfilter", 
                 ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile, "userdefinedplotfilter", false);
			if (p_dao.m_intError != 0)
			{
                 if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!! Error Creating ruledefinitionsplotfilter Table!!\r\n");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
			/***********************************************************************
			 **create a table link to the ruledefinitionsplotfilter table in the
			 **temporary MDB file located on the hard drive of the user
			 ***********************************************************************/
			p_dao.CreateTableLink(this.m_strTempMDBFile,"ruledefinitionsplotfilter",strMDBPathAndFile,"ruledefinitionsplotfilter",true);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Create link in " + this.m_strTempMDBFile + "\r\n");

                frmMain.g_oUtils.WriteText(m_strDebugFile,strMDBPathAndFile + "\truledefinitionsplotfilter\r\n");
            }


			p_dao=null;




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
			string[] strTableNames;
			strTableNames = new string[1];
			ado_data_access oAdo = new ado_data_access();			
			oAdo.OpenConnection(oAdo.getMDBConnString(m_strSystemResultsDbPathAndFile,"",""));

            strTableNames = oAdo.getTableNames(oAdo.m_OleDbConnection);
            for (int x = 0; x <= strTableNames.Length - 1; x++)
            {
                if (strTableNames[x] != null &&
                    strTableNames[x].Trim().Length > 0)
                {
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + strTableNames[x]);
                }
            }

            // Query the optimization variable for the selected revenue attribute so we can pass it to the table
            string strColumnFilterName = "";
            if (this.m_oOptimizationVariable.bUseFilter == true)
                strColumnFilterName = this.m_oOptimizationVariable.strRevenueAttribute;

            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateValidComboTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioResults.DefaultScenarioResultsValidCombosTableName);
			frmMain.g_oTables.m_oOptimizerScenarioResults.CreateValidComboFVSPrePostTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioResults.DefaultScenarioResultsValidCombosFVSPrePostTableName);
			frmMain.g_oTables.m_oOptimizerScenarioResults.CreateTieBreakerTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioResults.DefaultScenarioResultsTieBreakerTableName);
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateOptimizationTable(oAdo, oAdo.m_OleDbConnection, this.ReferenceOptimizerScenarioForm.OutputTablePrefix +
                Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix, strColumnFilterName);
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateEffectiveTable(oAdo, oAdo.m_OleDbConnection, this.ReferenceOptimizerScenarioForm.OutputTablePrefix, strColumnFilterName);
            string strScenarioResultsBestRxSummaryTableName = this.ReferenceOptimizerScenarioForm.OutputTablePrefix + Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryTableSuffix;
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateBestRxSummaryCycle1Table(oAdo, oAdo.m_OleDbConnection, strScenarioResultsBestRxSummaryTableName);
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateBestRxSummaryCycle1Table(oAdo, oAdo.m_OleDbConnection, strScenarioResultsBestRxSummaryTableName + "_before_tiebreaks");
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateProductYieldsTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName);
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateEconByRxUtilSumTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName);
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreatePostEconomicWeightedTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName);
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsHaulCostsTableName);
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateCondPsiteTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsCondPsiteTableName);
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateVersionTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsVersionTableName);
            
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
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
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(m_strContextDbPathAndFile, "", ""));

            strTableNames = oAdo.getTableNames(oAdo.m_OleDbConnection);
            for (int x = 0; x <= strTableNames.Length - 1; x++)
            {
                if (strTableNames[x] != null &&
                    strTableNames[x].Trim().Length > 0)
                {
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + strTableNames[x]);
                }
            }


            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHarvestMethodRefTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName);
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateRxPackageRefTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsRxPackageRefTableName);
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateDiameterSpeciesGroupRefTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsDiameterSpeciesGroupRefTableName);
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateFvsWeightedVariableRefTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsFvsWeightedVariablesRefTableName);
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateEconWeightedVariableRefTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsEconWeightedVariablesRefTableName);
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSpeciesGroupRefTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsSpeciesGroupRefTableName);
            frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioAdditionalHarvestCostsTable(oAdo, oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName + "_C");

            // Add the ad hoc additional harvest cost columns to table
            string[] strSourceColumnsArray = null;
            if (!ReferenceOptimizerScenarioForm.m_bProcessorUsingSqlite)
            {
                string strProcessorPath = ((frmMain)this._frmScenario.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsDbFile;
                oAdo.OpenConnection(oAdo.getMDBConnString(strProcessorPath, "", ""));
                string strSourceColumnsList = oAdo.getFieldNames(oAdo.m_OleDbConnection, "SELECT * FROM scenario_additional_harvest_costs");
                strSourceColumnsArray = frmMain.g_oUtils.ConvertListToArray(strSourceColumnsList, ",");
            }
            else
            {
                string strProcessorPath = ((frmMain)this._frmScenario.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
                string strConn = dataMgr.GetConnectionString(strProcessorPath);
                using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    oConn.Open();
                    strSourceColumnsArray = dataMgr.getFieldNamesArray(oConn, "SELECT * FROM scenario_additional_harvest_costs");
                }
                // Set up DSN for Processor config tables; We will need it later
                // Check to see if the input SQLite DSN exists and if so, delete so we can add
                if (odbcMgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName))
                {
                    odbcMgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName);
                }
                odbcMgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName, strProcessorPath);
            }

            oAdo.OpenConnection(oAdo.getMDBConnString(m_strContextDbPathAndFile, "", ""));
            foreach (string strColumn in strSourceColumnsArray)
            {
                if (! oAdo.ColumnExist(oAdo.m_OleDbConnection,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName + "_C", strColumn))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName + "_C",
                        strColumn, "DOUBLE", "");
                }
            }
            frmMain.g_oTables.m_oFvs.CreateRxHarvestCostColumnTable(oAdo, oAdo.m_OleDbConnection, Tables.FVS.DefaultRxHarvestCostColumnsTableName + "_C");
            
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
        }

		/// <summary>
		/// Copy the scenario results db file from the scenario?\db directory to the temp directory
        /// where the temp directory version is used during a single Optimizer run. Once
		/// the run successfully completes it is copied back to the scenario?\db directory.
		/// </summary>
		private void CopyScenarioResultsTable(string p_strDestPathAndDbFileName,string p_strSourcePathAndDbFileName)
		{

			dao_data_access oDao = new dao_data_access();


			if (System.IO.File.Exists(p_strSourcePathAndDbFileName))
			{
				System.IO.File.Copy(p_strSourcePathAndDbFileName,p_strDestPathAndDbFileName,true);	
			}
			else
			{
					oDao.CreateMDB(p_strDestPathAndDbFileName);
			}
			oDao.m_DaoWorkspace.Close();
			oDao=null;
		}

		/// <summary>
		/// create links to the tables located in the optimizer_results.accdb file
		/// </summary>
		private void CreateScenarioResultTableLinks()
		{
            
			string[] strTableNames;
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();
			p_dao.CreateTableLinks(this.m_strTempMDBFile,this.m_strSystemResultsDbPathAndFile);
			if (p_dao.m_intError==0)
			{

				int intCount = p_dao.getTableNames(m_strSystemResultsDbPathAndFile,ref strTableNames);
				if (p_dao.m_intError==0)
				{

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        if (intCount > 0)
                        {
                            for (int x = 0; x <= intCount - 1; x++)
                            {
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                     frmMain.g_oUtils.WriteText(m_strDebugFile,
                                         "scenario_results\t" + strTableNames[x] + "\r\n");
                            }
                        }
                    }
				}
				else
				{
					this.m_intError=p_dao.m_intError;
				}
			}
			else
			{
                if (frmMain.g_bDebug) frmMain.g_oUtils.WriteText(m_strDebugFile, p_dao.m_strError + "\r\n");
				this.m_intError = p_dao.m_intError;
			}

		
		}

        /// <summary>
        /// create links to the tables located in the context.accdb file
        /// </summary>
        private void CreateContextTableLinks()
        {

            string[] strTableNames;
            strTableNames = new string[1];
            dao_data_access p_dao = new dao_data_access();
            p_dao.CreateTableLinks(this.m_strTempMDBFile, this.m_strContextDbPathAndFile);
            if (p_dao.m_intError == 0)
            {

                int intCount = p_dao.getTableNames(m_strContextDbPathAndFile, ref strTableNames);
                if (p_dao.m_intError == 0)
                {

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        if (intCount > 0)
                        {
                            for (int x = 0; x <= intCount - 1; x++)
                            {
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile,
                                        "context\t" + strTableNames[x] + "\r\n");
                            }
                        }
                    }
                }
                else
                {
                    this.m_intError = p_dao.m_intError;
                }

            }
            else
            {
                if (frmMain.g_bDebug) frmMain.g_oUtils.WriteText(m_strDebugFile, p_dao.m_strError + "\r\n");
                this.m_intError = p_dao.m_intError;
            }

            string strRefMasterDir = ((frmMain)this._frmScenario.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\" + Tables.Reference.DefaultRxCategoryTableDbFile;
            p_dao.CreateTableLink(this.m_strTempMDBFile, Tables.Reference.DefaultRxCategoryTableName, strRefMasterDir, Tables.Reference.DefaultRxCategoryTableName);
            p_dao.CreateTableLink(this.m_strTempMDBFile, Tables.Reference.DefaultRxSubCategoryTableName, strRefMasterDir, Tables.Reference.DefaultRxSubCategoryTableName);
            string strOptimizerDir = ((frmMain)this._frmScenario.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\" + Tables.OptimizerDefinitions.DefaultDbFile;
            p_dao.CreateTableLink(this.m_strTempMDBFile, Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName, strOptimizerDir, Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName);
            p_dao.CreateTableLink(this.m_strTempMDBFile, Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName, strOptimizerDir, Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName);
            p_dao.CreateTableLink(this.m_strTempMDBFile, Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName, strOptimizerDir, Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName);
            if (!ReferenceOptimizerScenarioForm.m_bProcessorUsingSqlite)
            {
                string strProcessorDir = ((frmMain)this._frmScenario.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsDbFile;
                p_dao.CreateTableLink(this.m_strTempMDBFile, Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName, strProcessorDir, Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName);
                p_dao.CreateTableLink(this.m_strTempMDBFile, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName, strProcessorDir, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName);
            }
            else
            {
                string strProcessorDir = ((frmMain)this._frmScenario.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                p_dao.CreateSQLiteTableLink(this.m_strTempMDBFile, Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName, 
                    ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName, strProcessorDir);
                p_dao.CreateSQLiteTableLink(this.m_strTempMDBFile, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName,
                    ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName, strProcessorDir);
            }

            // We also need to create links to harvest_costs and tree vol val source tables in the context .accdb
            // so that we can create them directly with select into statement            
            p_dao.CreateTableLink(this.m_strContextDbPathAndFile, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName,
                                  this.m_oProcessorScenarioItem.DbPath + "\\" + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableDbFile, 
                                  Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName);
            p_dao.CreateTableLink(this.m_strContextDbPathAndFile, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName,
                      this.m_oProcessorScenarioItem.DbPath + "\\" + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableDbFile,
                      Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName);
           
            if (p_dao != null)
            {
                p_dao.m_DaoWorkspace.Close();
                p_dao = null;

            }
        }

        /// <summary>
        /// create links to selected reference tables in the results .accdb
        /// </summary>
        private void CreateReferenceResultsTableLinks()
        {
            // Create links to relevant tables elsewhere in BioSum project
            dao_data_access oDao = new dao_data_access();
            // plot
            oDao.CreateTableLink(m_strSystemResultsDbPathAndFile, this.m_strPlotTable, this.m_strPlotPathAndFile, this.m_strPlotTable);
            // cond
            oDao.CreateTableLink(m_strSystemResultsDbPathAndFile, this.m_strCondTable, this.m_strCondPathAndFile, this.m_strCondTable);
            // psites
            oDao.CreateTableLink(m_strSystemResultsDbPathAndFile, this.m_strPSiteTable, this.m_strPSitePathAndFile, this.m_strPSiteTable);

            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
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
            ado_data_access oAdo = new ado_data_access();



            //
            //FVS PRE VALID COMBO TABLE
            //
            oAdo.OpenConnection(oAdo.getMDBConnString(this.m_strFVSPreValidComboDbPathAndFile, "", ""));
            strTableNames = oAdo.getTableNames(oAdo.m_OleDbConnection);
            for (int x = 0; x <= strTableNames.Length - 1; x++)
            {
                if (strTableNames[x] != null &&
                    strTableNames[x].Trim().Length > 0)
                {
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + strTableNames[x]);
                }
            }



            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateValidComboFVSPreTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsValidCombosFVSPreTableName);

            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            //
            //FVS POST VALID COMBO TABLE
            //
            oAdo.OpenConnection(oAdo.getMDBConnString(this.m_strFVSPostValidComboDbPathAndFile, "", ""));
            strTableNames = oAdo.getTableNames(oAdo.m_OleDbConnection);
            for (int x = 0; x <= strTableNames.Length - 1; x++)
            {
                if (strTableNames[x] != null &&
                    strTableNames[x].Trim().Length > 0)
                {
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + strTableNames[x]);
                }
            }
            frmMain.g_oTables.m_oOptimizerScenarioResults.CreateValidComboFVSPostTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsValidCombosFVSPostTableName);
            oAdo.CloseConnection(oAdo.m_OleDbConnection);

        }

        /// <summary>
        /// create links to the tables located in the optimizer_results.accdb file
        /// </summary>
        private void CreateValidComboTableLinks()
        {

            string[] strTableNames;
            strTableNames = new string[1];
            dao_data_access p_dao = new dao_data_access();
            p_dao.CreateTableLinks(this.m_strTempMDBFile, this.m_strFVSPreValidComboDbPathAndFile);
            if (p_dao.m_intError == 0)
            {

                int intCount = p_dao.getTableNames(m_strFVSPreValidComboDbPathAndFile, ref strTableNames);
                if (p_dao.m_intError == 0)
                {

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        if (intCount > 0)
                        {
                            for (int x = 0; x <= intCount - 1; x++)
                            {
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile,
                                        "validcombo\t" + strTableNames[x] + "\r\n");
                            }
                        }
                    }
                }
                else
                {
                    this.m_intError = p_dao.m_intError;
                }
            }
            else
            {
                if (frmMain.g_bDebug) frmMain.g_oUtils.WriteText(m_strDebugFile, p_dao.m_strError + "\r\n");
                this.m_intError = p_dao.m_intError;
            }
            if (p_dao.m_intError == 0)
            {
                p_dao.CreateTableLinks(this.m_strTempMDBFile, this.m_strFVSPostValidComboDbPathAndFile);
                if (p_dao.m_intError == 0)
                {

                    int intCount = p_dao.getTableNames(m_strFVSPostValidComboDbPathAndFile, ref strTableNames);
                    if (p_dao.m_intError == 0)
                    {

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        {
                            if (intCount > 0)
                            {
                                for (int x = 0; x <= intCount - 1; x++)
                                {
                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                        frmMain.g_oUtils.WriteText(m_strDebugFile,
                                            "validcombo\t" + strTableNames[x] + "\r\n");
                                }
                            }
                        }
                    }
                    else
                    {
                        this.m_intError = p_dao.m_intError;
                    }
                }
            }


        }
        private void CreateProcessorScenarioResultTableLinks()
        {
            dao_data_access oDao = new dao_data_access();
            oDao.CreateTableLink(this.m_strTempMDBFile,
                                 Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName,
                                 this.m_oProcessorScenarioItem.DbPath + "\\" +
                                 Tables.ProcessorScenarioRun.DefaultHarvestCostsTableDbFile,
                                 Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName);
            if (oDao.m_intError == 0)
            {
                oDao.CreateTableLink(this.m_strTempMDBFile,
                                 Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName,
                                 this.m_oProcessorScenarioItem.DbPath + "\\" +
                                 Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsDbFile,
                                 Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName);
            }

            if (oDao.m_intError == 0)
            {
                m_strTreeVolValBySpcDiamGroupsTable = Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName;
                m_strHvstCostsTable = Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName;
            }
            oDao.m_DaoWorkspace.Close();
            oDao = null;



        }
		

		/// <summary>
		/// create links to the audit tables
		/// </summary>
		private void CreateAuditTableLinks()
		{

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateAuditTableLinks\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			string[] strTableNames;
            int y;
			strTableNames = new string[1];
			dao_data_access oDao = new dao_data_access();

            string strAccdbPathAndFile = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\" + Tables.Audit.DefaultCondAuditTableDbFile;

            if (System.IO.File.Exists(strAccdbPathAndFile))
            {
                int intCount = oDao.getTableNames(strAccdbPathAndFile, ref strTableNames);
                if (oDao.m_intError == 0)
                {
                    if (intCount > 0)
                    {
                        for (y = 0; y <= intCount - 1; y++)
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, strAccdbPathAndFile + "\t" + strTableNames[y] + "\r\n");
                            oDao.CreateTableLink(this.m_strTempMDBFile, strTableNames[y].Trim(), strAccdbPathAndFile, strTableNames[y]);
                            if (oDao.m_intError != 0)
                            {
                                if (frmMain.g_bDebug)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!Error Creating Table Link!!!");
                                    this.m_intError = oDao.m_intError;
                                    break;
                            }
                        }

                    }
                }
                else
                {
                    this.m_intError = oDao.m_intError;
                }
            }

			oDao = null;
			
		}
			
		/// <summary>
		/// create links to the scenario tables that contain the 
		/// rule definitions defined by the user
		/// </summary>
		private void CreateScenarioTableLinks()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateScenarioTableLinks\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

			string[] strTableNames;
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();
			string strConn="";
			
			string strMDBPathAndFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
                Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;

			//if (strMDBPathAndFile.Substring(strMDBPathAndFile.Trim().Length - 6,6).ToUpper()==".ACCDB")
			strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + 
				strMDBPathAndFile + ";User Id=admin;Password=;";
			//else
			//	strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";

			
			int intCount = p_dao.getTableNames(strMDBPathAndFile,ref strTableNames,false);
			if (p_dao.m_intError==0)
			{
				if (intCount > 0)
				{
					for (int x=0; x <= intCount-1;x++)
					{
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, strMDBPathAndFile + "\t" + strTableNames[x] + "\r\n");
						
						p_dao.CreateTableLink(this.m_strTempMDBFile,strTableNames[x],strMDBPathAndFile,strTableNames[x]);
						if (p_dao.m_intError != 0)
						{
                            if (frmMain.g_bDebug)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "!!Error Creating Table Link!!!");
							this.m_intError = p_dao.m_intError;
							break;
						}
					}

				}
			}
			else
			{
				this.m_intError=p_dao.m_intError;
			}
			p_dao = null;
			
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
             **get the psites table name and path
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

			this.m_strSQL="DELETE FROM scenario_psites_work_table";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_strSQL = "INSERT INTO scenario_psites_work_table (psite_id,trancd,biocd) " + 
				"SELECT psite_id,trancd,biocd " + 
				"FROM scenario_psites " + 
				"WHERE TRIM(scenario_id)='" + ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim() + "' AND " + 
				"selected_yn='Y';";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


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

            /**************************************************
             * Insert plot_id/cond_id records into PSITE_ACCESSIBLE_WORKTABLE
             * ************************************************/

            this.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName +
               "(biosum_cond_id, biosum_plot_id) ";
            //this.m_strSQL += "SELECT DISTINCT (c.biosum_cond_id), c.biosum_plot_id" +
            //                " FROM " + this.m_strCondTable + " c, travel_time " +
            //                " WHERE (((c.biosum_plot_id)=[travel_time].[biosum_plot_id])) " +
            this.m_strSQL += "SELECT DISTINCT biosum_cond_id, biosum_plot_id " +
                            "FROM " + this.m_strCondTable + " c " +
                            "GROUP BY c.biosum_cond_id, c.biosum_plot_id";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			
            
            /********************************************************************
			 **update cond_too_far_steep field to Y if slope is 
			 **<= 40% and the yarding distance >= to maximum yarding distance 
			 **for a slope <= 40%
			 ********************************************************************/
            this.m_strSQL = "UPDATE (" + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " p " +
                "INNER JOIN " + this.m_strCondTable + " c ON c.biosum_cond_id = p.biosum_cond_id)  " +
                "INNER JOIN " + this.m_strPlotTable + " AS pt ON c.biosum_plot_id = pt.biosum_plot_id " +
                "AND p.biosum_plot_id = pt.biosum_plot_id " +
                "SET p.cond_too_far_steep_yn = 'Y' " +
                "WHERE c.slope IS NOT NULL AND " +
                "c.slope <  " + m_oProcessorScenarioItem.m_oHarvestMethod.SteepSlopePercent + " AND " +
                "pt.gis_yard_dist_ft >= " + ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_cond_filter1.strNonSteepYardingDistance.Trim() + ";";
            
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			/********************************************************************
			 **update cond_too_far_steep field to Y if slope is 
			 **> 40% and the yarding distance >= to maximum yarding distance 
			 **for a slope > 40%
			 ********************************************************************/

            this.m_strSQL = "UPDATE (" + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " p " +
                "INNER JOIN " + this.m_strCondTable + " c ON c.biosum_cond_id = p.biosum_cond_id)  " +
                "INNER JOIN " + this.m_strPlotTable + " AS pt ON c.biosum_plot_id = pt.biosum_plot_id " +
                "AND p.biosum_plot_id = pt.biosum_plot_id " +
                "SET p.cond_too_far_steep_yn = 'Y' " +
                "WHERE c.slope IS NOT NULL AND " +
                "c.slope >= " + m_oProcessorScenarioItem.m_oHarvestMethod.SteepSlopePercent + " AND " +
                "pt.gis_yard_dist_ft >= " + ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_cond_filter1.strSteepYardingDistance.Trim() + ";";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();



			/*************************************************************
			 **set the remainder of the cond_too_far_steep_yn fields to N
			 *************************************************************/
            this.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " p " + 
				"SET p.cond_too_far_steep_yn = 'N' " + 
				"WHERE p.cond_too_far_steep_yn IS NULL OR (p.cond_too_far_steep_yn <> 'Y' AND " + 
				"p.cond_too_far_steep_yn <> 'N');";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			/*************************************************************
			 **update the condition accessible flag
			 *************************************************************/
            this.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + 
                " SET cond_accessible_yn = " + 
				"IIF(cond_too_far_steep_yn='Y','N','Y');";



            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			/********************************************************************
			 **get the current condition record counts for each plot
			 ********************************************************************/
			//insert the condition counts into the work table
			this.m_strSQL = "INSERT INTO plot_cond_accessible_work_table (biosum_plot_id, num_cond) " + 
				" SELECT biosum_plot_id , COUNT(biosum_plot_id) " +
                " FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + 
				" GROUP BY biosum_plot_id;";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			this.m_strSQL = "INSERT INTO plot_cond_accessible_work_table2 (biosum_plot_id, num_cond, num_cond_not_accessible) " + 
				"SELECT a.biosum_plot_id, a.num_cond,b.cond_not_accessible_count as num_cond_not_accessible FROM plot_cond_accessible_work_table a," +
                "(SELECT biosum_plot_id, Count(*) AS cond_not_accessible_count FROM " + 
                Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " " + 
				"WHERE cond_accessible_yn='N' GROUP BY biosum_plot_id) b " + 
				"WHERE a.biosum_plot_id = b.biosum_plot_id;";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

			if (this.m_ado.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");

			}
			else
			{
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!Error Executing SQL!!\r\n");
				this.m_intError = this.m_ado.m_intError;

			}
           


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

			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Processing Haul Costs");
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Visible",true);
			frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Refresh");

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nDelete all records from haul_costs table\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "-------------------------------------------------------------\r\n");
            }
            /********************************************
             **delete all records in the table
             ********************************************/
            this.m_strSQL = "delete from " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHaulCostsTableName;
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nUpdate Plot And Haul Cost Tables With Merch And Chip Haul Costs\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile,"-------------------------------------------------------------\r\n");
            }

			/********************************************
			 **get the haul cost per green ton per hour
			 ********************************************/
            strTruckHaulCost = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_costs1.RoadHaulCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTruckHaulCost = strTruckHaulCost.Replace(",","");
			
			if (strTruckHaulCost.Trim().Length == 1) strTruckHaulCost = "0.00";
			
            strRailHaulCost = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_costs1.RailHaulCostDollarsPerGreenTonPerMile.Replace("$", "").ToString();
			strRailHaulCost = strRailHaulCost.Replace(",","");
			if (strRailHaulCost.Trim().Length == 1) strRailHaulCost = "0.00";
			
			/***********************************************
			 **get the transfer cost per green to per hour
			 ***********************************************/
            strTransferMerchCost = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_costs1.RailMerchTransferCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTransferMerchCost = strTransferMerchCost.Replace(",","");
			if (strTransferMerchCost.Trim().Length == 1) strTransferMerchCost = "0.00";
			
            strTransferChipCost = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_costs1.RailChipTransferCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTransferChipCost = strTransferChipCost.Replace(",","");
			if (strTransferChipCost.Trim().Length == 1) strTransferChipCost = "0.00";
			 



			/*******************************************************************************
			 **zap the haul_costs table
			 *******************************************************************************/
			this.m_strSQL = "DELETE FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHaulCostsTableName;
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                 frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ndelete records in haul_costs table\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     
			
			
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Null The PSite Work Table's Haul Cost Fields");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			this.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + 
				" SET merch_haul_cost_id = null," + 
				"merch_haul_psite = null, " +
                "merch_haul_psite_name = null, " +
				"merch_haul_cost_dpgt = null ," + 
				"chip_haul_cost_id = null," + 
				"chip_haul_psite = null," +
                "chip_haul_psite_name = null," +
                "chip_haul_cost_dpgt = null;";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nnull the psite work table's haul cost fields\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;


            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();


			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     
			
			

			/*****************************************************************
			 **delete any records that may exist in the work tables
			 *****************************************************************/

			
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Delete Records In Work Tables");

            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			//all plots and road-accessible psite  work table
			this.m_strSQL = "delete from all_road_merch_haul_costs_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "delete from all_road_chip_haul_costs_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//railheads to rail-accessible collector psite work tables
			this.m_strSQL = "delete from merch_rh_to_collector_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "delete from chip_rh_to_collector_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//plots to railheads to rail-accessible collector psite work tables
			this.m_strSQL = "delete from merch_plot_to_rh_to_collector_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "delete from chip_plot_to_rh_to_collector_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//cheapest merch road route work table
			this.m_strSQL = "delete from cheapest_road_merch_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			//cheapest chip road route work table
			this.m_strSQL = "delete from cheapest_road_chip_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			//cheapest merch rail route work table
			this.m_strSQL = "delete from cheapest_rail_merch_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			//cheapest chip rail route work table
			this.m_strSQL = "delete from cheapest_rail_chip_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


			//combine cheapest road and rail into single table
			this.m_strSQL = "DELETE FROM combine_merch_rail_road_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "DELETE FROM combine_chip_rail_road_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//overall cheapest routes for merch and chip
			this.m_strSQL = "delete from cheapest_merch_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "delete from cheapest_chip_haul_costs_work_table;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     

			//MERCH AND CHIP ROAD PROCESSING SITE HAUL COSTS
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Road Haul Costs For Merchantable Wood Processing Sites");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\n\r\n--Merchantable wood processing site haul costs--\r\n");
			/***************************************************************
			 **process the merch travel times first
			 ***************************************************************/
			/*****************************************************************
			 **Insert into a table all travel time records where the psite 
			 **has road only or road/rail access and processes 
			 **merch only or merch/chip
			 *****************************************************************/
			this.m_strSQL = "INSERT into all_road_merch_haul_costs_work_table " + 
				"SELECT t.biosum_plot_id, 0 AS railhead_id," +
                "0 AS transfer_cost_dpgt, s.psite_id," +
                "(" + strTruckHaulCost.Trim() + " * t.one_way_hours) AS road_cost_dpgt," +
                "0 AS rail_cost_dpgt, (transfer_cost_dpgt+road_cost_dpgt+rail_cost_dpgt) AS complete_haul_cost_dpgt," + 
				"'M' as materialcd " +
				"FROM " + this.m_strTravelTimeTable + " t," + 
				this.m_strPSiteWorkTable + " s " + 
				"WHERE t.psite_id=s.psite_id AND " + 
				"(s.trancd=1 OR s.trancd =3) AND " +
                "(s.biocd=3 OR s.biocd=1)  AND t.one_way_hours > 0;";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into work table all travel time records where psite has road access and processes merch\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
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
			this.m_strSQL = "INSERT INTO cheapest_road_merch_haul_costs_work_table " + 
				"SELECT b.biosum_plot_id,c.psite_id,null AS railhead_id," +
                "0 AS transfer_cost_dpgt, a.road_cost_dpgt, 0 AS rail_cost_dpgt," +
                "b.min_cost AS complete_haul_cost_dpgt," + 
				"'M' AS materialcd " + 
				"FROM  all_road_merch_haul_costs_work_table  a," +
                "(SELECT biosum_plot_id,MIN(complete_haul_cost_dpgt) AS min_cost " + 
				"FROM all_road_merch_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id)  b," + 
				"(SELECT biosum_plot_id,  psite_id ," +
                "MIN(complete_haul_cost_dpgt) AS min_cost2 " + 
				"FROM all_road_merch_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id, psite_id) c " + 
				"WHERE  c.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.psite_id = c.psite_id AND " + 
				"b.min_cost = c.min_cost2;";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into work table. Find the cheapest plot to merch processing site road route.\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Road Haul Costs For Chip Wood Processing Sites");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			/***********************************************************************
			 **Append to a table all travel time records where the psite 
			 **has road only or road/rail access and processes 
			 **chip only or merch/chip
			 ***********************************************************************/
			this.m_strSQL = "INSERT INTO all_road_chip_haul_costs_work_table " + 
				"SELECT t.biosum_plot_id, 0 AS railhead_id," +
                "0 AS transfer_cost_dpgt,s.psite_id," +
                "(" + strTruckHaulCost.Trim() + " * t.one_way_hours) AS road_cost_dpgt," +
                "0 AS rail_cost_dpgt, " +
                "(transfer_cost_dpgt+road_cost_dpgt+rail_cost_dpgt) AS complete_haul_cost_dpgt," + 
				"'C' AS materialcd " +
                "FROM " + this.m_strTravelTimeTable + " t," + 
				this.m_strPSiteWorkTable + " s " + 
				"WHERE t.psite_id=s.psite_id AND " + 
				"(s.trancd=1 OR s.trancd=3) AND " +
                "(s.biocd=3 OR s.biocd=2)  AND t.one_way_hours > 0;";

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into work table all travel time records where psite has road access and processes chips.\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     


			/******************************************************************
			 **For each plot get the cheapest road route to a psite. 
			 ******************************************************************/
			this.m_strSQL = "INSERT INTO cheapest_road_chip_haul_costs_work_table " + 
				"SELECT b.biosum_plot_id, c.psite_id, null AS railhead_id," +
                "0 AS transfer_cost_dpgt,a.road_cost_dpgt," +
                "0 AS rail_cost_dpgt, b.min_cost AS complete_haul_cost_dpgt," + 
				"'C' AS materialcd " + 
				"FROM all_road_chip_haul_costs_work_table  a," +
                "(SELECT biosum_plot_id,MIN(complete_haul_cost_dpgt) AS min_cost " + 
				"FROM all_road_chip_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id)  b," + 
				"(SELECT biosum_plot_id,  psite_id ," +
                "MIN(complete_haul_cost_dpgt) AS min_cost2 " + 
				"FROM all_road_chip_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id, psite_id)  c " + 
				"WHERE c.biosum_plot_id=b.biosum_plot_id AND " + 
				"a.biosum_plot_id=b.biosum_plot_id AND " + 
				"a.psite_id=c.psite_id AND b.min_cost=c.min_cost2;";



			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into work table. Find the cheapest plot to chip processing site road route.\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     

			//MERCH AND CHIP RAIL PROCESSING SITE HAUL COSTS
			/*********************************************************
			 **Append to a table all travel time collector_id (psite)
			 **records where the psite has rail access
			 *********************************************************/
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Rail Haul Costs For Merchantable Wood Processing Sites");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			this.m_strSQL = "INSERT INTO merch_rh_to_collector_haul_costs_work_table " + 
				"SELECT t.psite_id AS railhead_id," + 
				"t.collector_id AS psite_id," +
                strTransferMerchCost.Trim() + " AS transfer_cost_dpgt," +
                "0 AS road_cost_dpgt," +
                "((t.one_way_hours * 45) * " + strRailHaulCost.Trim() + ") AS rail_cost_dpgt," +
                "0 AS complete_haul_cost_dpgt,  'M' AS materialcd " +
                "FROM " + this.m_strTravelTimeTable + " t  " + 
				"INNER JOIN  " + this.m_strPSiteWorkTable + " s " + 
				"ON t.collector_id = s.psite_id " +
                "WHERE  s.trancd=3 And (s.biocd=3 Or s.biocd=1)  AND t.one_way_hours > 0 AND " + 
				"EXISTS (SELECT ss.psite_id " + 
				"FROM " + this.m_strPSiteWorkTable + " ss " + 
				"WHERE t.psite_id=ss.psite_id AND ss.trancd=2 AND (ss.biocd=3 Or ss.biocd=1));";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into work table all travel time collector_id (psite) records where the psite has rail access and processes merch.\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
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
			this.m_strSQL = "INSERT INTO merch_plot_to_rh_to_collector_haul_costs_work_table " + 
				"SELECT  t.biosum_plot_id, r.railhead_id, r.psite_id," +
                "r.transfer_cost_dpgt," +
                "(" + strTruckHaulCost.Trim() + " * t.one_way_hours) AS road_cost_dpgt," +
                "r.rail_cost_dpgt, (r.transfer_cost_dpgt + road_cost_dpgt + r.rail_cost_dpgt) AS complete_haul_cost_dpgt," +
				"'M' AS materialcd " + 
				"FROM  " + this.m_strTravelTimeTable + " t," + 
				"merch_rh_to_collector_haul_costs_work_table r " +
                "WHERE r.railhead_id = t.psite_id  AND t.one_way_hours > 0;";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into work table travel time plot records and previous work rail/merch table results\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     


			this.m_strSQL = "UPDATE merch_plot_to_rh_to_collector_haul_costs_work_table " +
                "SET complete_haul_cost_dpgt = transfer_cost_dpgt + road_cost_dpgt + rail_cost_dpgt;";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nupdate merch by road and rail total haul cost\r\n ");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
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
			this.m_strSQL = "INSERT INTO cheapest_rail_merch_haul_costs_work_table " + 
				"SELECT a.biosum_plot_id,c.psite_id, a.railhead_id," +
                "a.transfer_cost_dpgt, a.road_cost_dpgt,a.rail_cost_dpgt," +
                "c.min_cost AS complete_haul_cost_dpgt,'M' as materialcd " +
				"FROM merch_plot_to_rh_to_collector_haul_costs_work_table a," +
                "(SELECT biosum_plot_id, MIN(complete_haul_cost_dpgt) AS min_cost2 " + 
				"FROM merch_plot_to_rh_to_collector_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id) b," + 
				"(SELECT biosum_plot_id, psite_id," +
                "MIN(complete_haul_cost_dpgt) AS min_cost " + 
				"FROM merch_plot_to_rh_to_collector_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id,psite_id) c " + 
				"WHERE  c.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.biosum_plot_id = c.biosum_plot_id AND " + 
				"a.psite_id = c.psite_id AND " +
                "a.complete_haul_cost_dpgt = c.min_cost AND " + 
				"min_cost2 = min_cost;";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Find the cheapest plot to merch processing site rail routes\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     

			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Rail Haul Costs For Chip Wood Processing Sites");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			/***********************************************************************
			 **Append to a table all travel time collector_id (psite) records
			 **where the psite has rail access and processes chips only or 
			 **both merch/chips
			 ***********************************************************************/
			this.m_strSQL = "INSERT INTO chip_rh_to_collector_haul_costs_work_table " + 
				"SELECT  t.psite_id AS railhead_id,"  + 
				"t.collector_id AS psite_id," +
                strTransferChipCost.Trim() + " AS transfer_cost_dpgt," +
                "0 AS road_cost_dpgt," +
                "((t.one_way_hours * 45) * " + strRailHaulCost.Trim() + ") AS rail_cost_dpgt," +
                "0 AS complete_haul_cost_dpgt,  'C' AS materialcd " +
                "FROM " + this.m_strTravelTimeTable + " t  " + 
				"INNER JOIN  " + this.m_strPSiteWorkTable + " s " + 
				"ON t.collector_id = s.psite_id " + 
				"WHERE s.trancd=3 AND  " +
                "(s.biocd=3 OR s.biocd=2)  AND t.one_way_hours > 0 AND " + 
				"EXISTS (SELECT ss.psite_id " + 
				"FROM " + this.m_strPSiteWorkTable + " ss " + 
				"WHERE t.psite_id=ss.psite_id AND ss.trancd=2 AND (ss.biocd=3 Or ss.biocd=2));";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table all travel time collector_id (psite) records where the psite has rail access and processes chips.\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
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
			this.m_strSQL = "INSERT INTO chip_plot_to_rh_to_collector_haul_costs_work_table " + 
				"SELECT  t.biosum_plot_id, r.railhead_id, r.psite_id," +
                "r.transfer_cost_dpgt," +
                "(" + strTruckHaulCost.Trim() + " * t.one_way_hours) AS road_cost_dpgt," +
                "r.rail_cost_dpgt, " +
                "(r.transfer_cost_dpgt + road_cost_dpgt + r.rail_cost_dpgt) AS complete_haul_cost_dpgt," + 
				"'C' AS materialcd " + 
				"FROM  " + this.m_strTravelTimeTable + " t," + 
				"chip_rh_to_collector_haul_costs_work_table r " +
                "WHERE r.railhead_id = t.psite_id  AND t.one_way_hours > 0;";




			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into work table travel time plot records and previous rail/chips work table results\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}    
 
			this.m_strSQL = "UPDATE chip_plot_to_rh_to_collector_haul_costs_work_table " +
                "SET complete_haul_cost_dpgt = transfer_cost_dpgt + road_cost_dpgt + rail_cost_dpgt;";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nupdate chips by road and rail total haul cost\r\n ");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
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
			this.m_strSQL = "INSERT INTO cheapest_rail_chip_haul_costs_work_table " + 
				"SELECT a.biosum_plot_id,b.psite_id, a.railhead_id," +
                "a.transfer_cost_dpgt, a.road_cost_dpgt, a.rail_cost_dpgt," +
                "b.min_cost AS complete_haul_cost_dpgt,'C' AS materialcd " + 
				"FROM chip_plot_to_rh_to_collector_haul_costs_work_table a," + 
				"(SELECT biosum_plot_id, " +
                "MIN(complete_haul_cost_dpgt) AS min_cost2 " + 
				"FROM chip_plot_to_rh_to_collector_haul_costs_work_table " +
				"GROUP BY biosum_plot_id) c, " +
				"(SELECT biosum_plot_id, psite_id," +
                "MIN(complete_haul_cost_dpgt) AS min_cost " + 
				"FROM chip_plot_to_rh_to_collector_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id,psite_id) b " + 
				"WHERE  b.biosum_plot_id = c.biosum_plot_id AND " + 
				"a.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.psite_id = b.psite_id AND  " +
                "a.complete_haul_cost_dpgt = b.min_cost AND " + 
				"min_cost2 = min_cost;";

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Find the cheapest plot to chip processing site rail routes\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     

		    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Combine Road And Rail Haul Costs For Merchantable Wood Processing Sites");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");


			/**************************************************************
			 **combine the cheapest road and rail total cost for each plot
			 **to a merch psite
			 **After the insert there should be two records for each
			 **plot - one with cheapest haul cost by road and another
			 **with cheapest haul cost by rail
			 ***************************************************************/
			this.m_strSQL = "INSERT INTO combine_merch_rail_road_haul_costs_work_table " + 
				"SELECT * FROM cheapest_road_merch_haul_costs_work_table;";

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Cheapest road route to merch psite\r\n ");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     

            // Need to specify all fields except the haul_cost_id because there may be duplicate haul_cost_id's  
            // between the rail and road tables. This allows MS Access to auto-assign the haul_cost_id for the
            // inserted records
            this.m_strSQL = "INSERT INTO combine_merch_rail_road_haul_costs_work_table " +
                "SELECT biosum_plot_id, railhead_id, psite_id, transfer_cost_dpgt, road_cost_dpgt, " +
                "rail_cost_dpgt, complete_haul_cost_dpgt, materialcd FROM cheapest_rail_merch_haul_costs_work_table;";

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Cheapest rail route to merch psite\r\n ");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}   
  

			/***************************************************
			 **Get the overall cheapest merch route
			 ***************************************************/
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Get Overall Least Expensive Merch Route");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
			this.m_strSQL = "INSERT INTO cheapest_merch_haul_costs_work_table " + 
				"SELECT a.biosum_plot_id,b.psite_id, a.railhead_id," +
                "a.transfer_cost_dpgt, a.road_cost_dpgt,  a.rail_cost_dpgt," +
                "b.min_cost AS complete_haul_cost_dpgt,'M' AS materialcd " + 
				"FROM combine_merch_rail_road_haul_costs_work_table a," +
                "(SELECT biosum_plot_id, MIN(complete_haul_cost_dpgt) AS min_cost2 " + 
				"FROM combine_merch_rail_road_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id) c, " + 
				"(SELECT biosum_plot_id, psite_id," +
                "MIN(complete_haul_cost_dpgt) AS min_cost " + 
				"FROM combine_merch_rail_road_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id,psite_id) b " + 
				"WHERE  b.biosum_plot_id = c.biosum_plot_id AND " + 
				"a.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.psite_id = b.psite_id AND " +
                "a.complete_haul_cost_dpgt = b.min_cost AND " + 
				"min_cost2 = min_cost;";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Get the overall cheapest merch route\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}   

			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Combine Road And Rail Haul Costs For Chip Wood Processing Sites");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");

			/**************************************************************
			 **combine the cheapest road and rail total cost for each plot
			 **to a chips psite
			 **After the insert there should be two records for each
			 **plot - one with cheapest haul cost by road and another
			 **with cheapest haul cost by rail
			 ***************************************************************/
			this.m_strSQL = "INSERT INTO combine_chip_rail_road_haul_costs_work_table " + 
				"SELECT * FROM cheapest_road_chip_haul_costs_work_table;";

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Cheapest road route to chip psite\r\n ");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}     

            // Need to specify all fields except the haul_cost_id because there may be duplicate haul_cost_id's  
            // between the rail and road tables. This allows MS Access to auto-assign the haul_cost_id for the
            // inserted records
            this.m_strSQL = "INSERT INTO combine_chip_rail_road_haul_costs_work_table " +
                "SELECT biosum_plot_id, railhead_id, psite_id, transfer_cost_dpgt, road_cost_dpgt, " +
                "rail_cost_dpgt, complete_haul_cost_dpgt, materialcd FROM cheapest_rail_chip_haul_costs_work_table;";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Cheapest rail route to chip psite\r\n ");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}   
  
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Get Overall Least Expensive Chip Route");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");


			/******************************************
			 **Get the overall cheapest chip route
			 ******************************************/
			this.m_strSQL = "INSERT INTO cheapest_chip_haul_costs_work_table " + 
				"SELECT a.biosum_plot_id,b.psite_id, a.railhead_id," +
                "a.transfer_cost_dpgt, a.road_cost_dpgt,  a.rail_cost_dpgt," +
                "b.min_cost AS complete_haul_cost_dpgt,'C' AS materialcd " + 
				"FROM combine_chip_rail_road_haul_costs_work_table a, " +
                "(SELECT biosum_plot_id,MIN(complete_haul_cost_dpgt) AS min_cost2 " + 
				"FROM combine_chip_rail_road_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id) c, " + 
				"(SELECT biosum_plot_id, psite_id," +
                "MIN(complete_haul_cost_dpgt) AS min_cost " + 
				"FROM combine_chip_rail_road_haul_costs_work_table " + 
				"GROUP BY biosum_plot_id,psite_id) b " + 
				"WHERE  b.biosum_plot_id = c.biosum_plot_id AND " + 
				"a.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.psite_id = b.psite_id AND " +
                "a.complete_haul_cost_dpgt = b.min_cost AND " + 
				"min_cost2 = min_cost;";



			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into work table. Get the overall cheapest chip route\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}   




			//INSERT INTO HAUL_COSTS TABLE
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Inserting Results Into Haul Costs Table");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
			this.m_strSQL = "INSERT INTO haul_costs " + 
				"SELECT * FROM cheapest_merch_haul_costs_work_table;";
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into haul_costs table cheapest merch route for each plot\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}   


			this.m_strSQL = "INSERT INTO haul_costs " + 
				"SELECT * FROM cheapest_chip_haul_costs_work_table;";
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nInsert into haul_costs table cheapest chip route for each plot\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
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

            this.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " p " + 
				"INNER JOIN haul_costs h " + 
				"ON p.biosum_plot_id = h.biosum_plot_id " +
                "SET p.merch_haul_cost_id = h.haul_cost_id," + 
				"p.merch_haul_psite = h.psite_id," +
                "p.merch_haul_cost_dpgt=h.complete_haul_cost_dpgt " +
				"WHERE h.materialcd='M';";

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nUpdate plot merch haul cost fields\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

            this.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " w " +
                "INNER JOIN processing_site p " +
                "ON w.merch_haul_psite = p.psite_id " +
                "SET w.merch_haul_psite_name = p.name";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate merch psite name\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			} 
  

			/*****************************************
			 **Update  cheapest chip routes
			 *****************************************/
            this.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " p " + 
				"INNER JOIN haul_costs h " + 
				"ON p.biosum_plot_id = h.biosum_plot_id " + 
				"SET p.chip_haul_cost_id = h.haul_cost_id," + 
				"p.chip_haul_psite = h.psite_id," +
                "p.chip_haul_cost_dpgt=h.complete_haul_cost_dpgt " + 
				"WHERE h.materialcd='C';";


			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nUpdate plot chip haul cost fields\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); 
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

            this.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " w " +
                "INNER JOIN processing_site p " +
                "ON w.chip_haul_psite = p.psite_id " +
                "SET w.chip_haul_psite_name = p.name";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate chip psite name\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); ;
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
            
            
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			} 
  

			/******************************************
			 **clean up work tables
			 ******************************************/

			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Cleaning Up Haul Cost Work Tables...Stand By");
            frmMain.g_oDelegate.ExecuteControlMethod((Control)ReferenceUserControlScenarioRun.lblMsg, "Refresh");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nCleaning up haul cost work tables\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); 
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
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
			
            // if the work table doesn't exist, create it
            if (! this.m_ado.TableExist(this.m_TempMDBFileConn, m_strTreeVolValSumTable))
            {
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateTreeVolValSumTable(this.m_ado, this.m_TempMDBFileConn, m_strTreeVolValSumTable);
            }
            // otherwise delete any existing rows in table
            else
            {
                this.m_strSQL = "delete from " + m_strTreeVolValSumTable;
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            }
			
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				this.m_intError = this.m_ado.m_intError;
				return;
			}
			this.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsTreeVolValSumTableName + 
				" (biosum_cond_id,rxpackage,rx,rxcycle,chip_vol_cf," + 
				"chip_wt_gt,chip_val_dpa,merch_vol_cf," + 
				"merch_wt_gt,merch_val_dpa,place_holder) ";
			this.m_strSQL += "SELECT biosum_cond_id, " + 
				"rxpackage,rx,rxcycle," +
				"Sum(chip_vol_cf) AS chip_vol_cf," + 
				"Sum(chip_wt_gt) AS chip_wt_gt," + 
				"Sum(chip_val_dpa) AS chip_val_dpa," + 
				"Sum(merch_vol_cf) AS merch_vol_cf," +
				"Sum(merch_wt_gt) AS merch_wt_gt," + 
				"Sum(merch_val_dpa) AS merch_val_dpa, " +
                "place_holder";

			this.m_strSQL += " FROM " + this.m_strTreeVolValBySpcDiamGroupsTable.Trim();
			this.m_strSQL += " GROUP BY biosum_cond_id,rxpackage,rx,rxcycle, place_holder";
			this.m_strSQL += " ORDER BY biosum_cond_id,rxpackage,rx,rxcycle, place_holder ;";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\ninsert into tree_vol_val_sum_by_rx_cycle_work table tree volume and value sums\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); 

			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_ado.m_intError != 0)
			{
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");

				this.m_intError = this.m_ado.m_intError;

				return;
			}

            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


			if (this.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");

			}
		}
        /// <summary>
        /// sum the tree_vol_val_by_species_diam_groups table values to tree_vol_val_sum_by_rxpackage
        /// </summary>
        private void sumTreeVolValByRxPackage()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//sumTreeVolValByRxPackage\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 6;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;

            /**************************************************************
             **sum the tree_vol_val_by_species_diam_groups table to
             **        tree_vol_val_sum_by_rxpackage
             **************************************************************/
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n\r\nSum Tree Volumes and Values By Treatment Package\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "------------------------------------------------\r\n");
            }
           


            this.m_strSQL = "delete from tree_vol_val_sum_by_rxpackage";
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }
            this.m_strSQL = "INSERT INTO tree_vol_val_sum_by_rxpackage " +
                "(biosum_cond_id,rxpackage,chip_vol_cf," +
                "chip_wt_gt,chip_val_dpa,merch_vol_cf," +
                "merch_wt_gt,merch_val_dpa) ";
            this.m_strSQL += "SELECT biosum_cond_id, " +
                "rxpackage," +
                "Sum(chip_vol_cf) AS chip_vol_cf," +
                "Sum(chip_wt_gt) AS chip_wt_gt," +
                "Sum(chip_val_dpa) AS chip_val_dpa," +
                "Sum(merch_vol_cf) AS merch_vol_cf," +
                "Sum(merch_wt_gt) AS merch_wt_gt," +
                "Sum(merch_val_dpa) AS merch_val_dpa ";

            this.m_strSQL += " FROM " + this.m_strTreeVolValBySpcDiamGroupsTable.Trim();
            this.m_strSQL += " GROUP BY biosum_cond_id,rxpackage";
            this.m_strSQL += " ORDER BY biosum_cond_id,rxpackage;";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into tree_vol_val_sum_by_rxpackage table tree volume and value sums\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n"); 

            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }

            // Create worktable for HVST_TYPE_BY_CYCLE field
            string strTreeSumWorktableName = "TREE_SUM_WORKTABLE";
            if (this.m_ado.TableExist(this.m_TempMDBFileConn, strTreeSumWorktableName))
            {
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, "DELETE FROM " + strTreeSumWorktableName);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nDelete all records from tree_sum_worktable\r\n");

            }
            else
            {
                this.m_strSQL = "CREATE TABLE " + strTreeSumWorktableName + " (" +
                                "biosum_cond_id CHAR(25), " +
                                "rxpackage CHAR(3), " +
                                "CYCLE_1 CHAR(1), " +
                                "CYCLE_2 CHAR(1), " +
                                "CYCLE_3 CHAR(1), " +
                                "CYCLE_4 CHAR(1))";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nCreate tree_sum_worktable\r\n");

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            }

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }

            //Add rows to worktable
            this.m_strSQL = "INSERT INTO " + strTreeSumWorktableName +
                            " SELECT BIOSUM_COND_ID, RXPACKAGE, '0' as CYCLE_1, '0' AS CYCLE_2, '0' AS CYCLE_3, '0' AS CYCLE_4" +
                            " FROM TREE_VOL_VAL_SUM_BY_RXPACKAGE";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nAdd condition/rxpackages to tree_sum_worktable\r\n");

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }


            // Update worktable from tree_vol_val_sum_by_rxpackage
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate cycle values on tree_sum_worktable\r\n");
            string[] arrFieldToUpdate = {"CYCLE_1", "CYCLE_2", "CYCLE_3", "CYCLE_4"};
            string[] arrRxCycle= {"1","2","3","4"};
            for (int arrIdx = 0; arrIdx < 4; arrIdx++)
            {
                this.m_strSQL = "UPDATE " + strTreeSumWorktableName +
                                " INNER JOIN tree_vol_val_sum_by_rx_cycle_work ON (" + strTreeSumWorktableName + ".BIOSUM_COND_ID=tree_vol_val_sum_by_rx_cycle_work.BIOSUM_COND_ID)" + 
                                " AND (" + strTreeSumWorktableName + ".RXPACKAGE= tree_vol_val_sum_by_rx_cycle_work.RXPACKAGE)" +
                                " SET " + strTreeSumWorktableName + "." + arrFieldToUpdate[arrIdx] + " = IIF (chip_vol_cf > 0 AND merch_vol_cf > 0, '3'," + 
                                " IIF (chip_vol_cf = 0 AND merch_vol_cf = 0, '0'," +
                                " IIF (chip_vol_cf > 0, '2','1'))) " +
                                " WHERE RXCYCLE = '" + arrRxCycle[arrIdx] + "'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            }
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }

            // Populate hvst_type_by_cycle from worktable
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate hvst_type_by_cycle from tree_sum_worktable\r\n");
            this.m_strSQL = "UPDATE tree_vol_val_sum_by_rxpackage" +
                            " INNER JOIN " + strTreeSumWorktableName + 
                            " ON (" + strTreeSumWorktableName + ".BIOSUM_COND_ID=TREE_VOL_VAL_SUM_BY_RXPACKAGE.BIOSUM_COND_ID)" +
                            " AND (" + strTreeSumWorktableName +".RXPACKAGE= TREE_VOL_VAL_SUM_BY_RXPACKAGE.RXPACKAGE)" +
                            " SET hvst_type_by_cycle = CYCLE_1 + CYCLE_2 + CYCLE_3 + CYCLE_4 ";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }

            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");

            }
        }
		private void getHaulCost(string p_strBiosumPlotId, int p_intPSiteId)
		{
			/********************************************
			 **get the haul cost per green ton per hour
			 ********************************************/
            string strTruckHaulCost = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_costs1.RoadHaulCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTruckHaulCost = strTruckHaulCost.Replace(",","");
			
			if (strTruckHaulCost.Trim().Length == 1) strTruckHaulCost = "0.00";
			
            string strRailHaulCost = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_costs1.RailHaulCostDollarsPerGreenTonPerMile.Replace("$", "").ToString();
			strRailHaulCost = strRailHaulCost.Replace(",","");
			if (strRailHaulCost.Trim().Length == 1) strRailHaulCost = "0.00";
			
			/***********************************************
			 **get the transfer cost per green to per hour
			 ***********************************************/
            string strTransferMerchCost = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_costs1.RailMerchTransferCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTransferMerchCost = strTransferMerchCost.Replace(",","");
			if (strTransferMerchCost.Trim().Length == 1) strTransferMerchCost = "0.00";
			
            string strTransferBioCost = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_costs1.RailChipTransferCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTransferBioCost = strTransferBioCost.Replace(",","");
			if (strTransferBioCost.Trim().Length == 1) strTransferBioCost = "0.00";
			



		}

		/// <summary>
		/// load the validcombos table with biosum_cond_id,rxpackage,rx and rxcycle values
		/// that exist in the user defined plot filters, condition, ffe, travel times, and
		/// harvest cost, and tree volume/value tables
		/// </summary>
		private void validcombos()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//validcombos\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			string strRxList="";
			string strGrpCd="";
			int x=0;
			string strTable="";
            int intListViewIndex = -1;
            System.Windows.Forms.CheckBox oCheckBox = null;
            
			//int y=0;
			//int z=0;

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



			/*****************************************************************
			 **delete audit tables
			 *****************************************************************/

            this.m_strSQL = "delete from " + Tables.Audit.DefaultCondAuditTableName;
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }


            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            this.m_strSQL = "delete from " + Tables.Audit.DefaultCondRxAuditTableName;
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }
			
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			/**************************
			 **get the treatment list
			 **************************/
			this.m_strSQL = "SELECT rx FROM " + this.m_strRxTable + ";";
			this.m_ado.SqlQueryReader(this.m_TempMDBFileConn,this.m_strSQL);
			if (!this.m_ado.m_OleDbDataReader.HasRows)
			{
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				this.m_intError = -1;
				MessageBox.Show("No Treatments Found In The Treatment Table");
				return;
			}
			while (this.m_ado.m_OleDbDataReader.Read())
			{
				strRxList+=this.m_ado.m_OleDbDataReader["rx"].ToString().Trim();
			}
			this.m_ado.m_OleDbDataReader.Close();

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();


            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute User Defined Plot SQL And Insert Resulting Records Into Table userdefinedplotfilter_work\r\n");
			
            this.m_strSQL = "INSERT INTO userdefinedplotfilter_work " + this.m_strUserDefinedPlotSQL;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"--Execute User Defined Cond SQL And Insert Resulting Records Into Table userdefinedcondfilter_work--\r\n");
			this.m_strSQL = "INSERT INTO userdefinedcondfilter_work " + this.m_strUserDefinedCondSQL;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
			
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"--Execute rule definition filters for the condition table. The filters include ownership and condition accessible--\r\n");
			this.m_strSQL = "INSERT INTO ruledefinitionscondfilter SELECT c.* FROM " + this.m_strCondTable + " AS c ";
			this.m_strSQL += " WHERE c.owngrpcd IN (";

			//usfs ownnership
			if (ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_owner_groups1.chkOwnGrp10.Checked==true)
			{
				strGrpCd = "10,1";
			}
			//other federal ownership
			if (ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_owner_groups1.chkOwnGrp20.Checked==true)
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
			if (ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_owner_groups1.chkOwnGrp30.Checked==true)
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
			if (ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_owner_groups1.chkOwnGrp40.Checked==true)
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
			this.m_strSQL +=  strGrpCd + ")";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();


			 if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"--Execute SQL that deletes from the condition rule definitions table (ruledefinitionscondfilter) those biosum_cond_id that are not found in the user defined condition SQL filter table (userdefinedcondfilter_work)--\r\n");

			this.m_strSQL = "DELETE FROM ruledefinitionscondfilter a " + 
				            "WHERE NOT EXISTS " + 
				                "(SELECT b.biosum_cond_id " + 
								 "FROM userdefinedcondfilter_work b " + 
				                 "WHERE a.biosum_cond_id =  b.biosum_cond_id)";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();


			 if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"--Execute SQL That Includes  Rule Definitions Defined By The User Into Table ruledefinitionsplotfilter--\r\n");
			this.m_strSQL = "INSERT INTO ruledefinitionsplotfilter SELECT DISTINCT userdefinedplotfilter_work.* from userdefinedplotfilter_work INNER JOIN ruledefinitionscondfilter ON userdefinedplotfilter_work.biosum_plot_id = ruledefinitionscondfilter.biosum_plot_id" ;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			this.m_strSQL = "DELETE FROM validcombos;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            _uc_scenario_run.uc_filesize_monitor3.BeginMonitoringFile(
                          m_strFVSPreValidComboDbPathAndFile, 2000000000, "2GB");
            _uc_scenario_run.uc_filesize_monitor3.Information = "Valid combinations for FVS Pre-Treatment records";

            _uc_scenario_run.uc_filesize_monitor4.BeginMonitoringFile(
                           m_strFVSPostValidComboDbPathAndFile, 2000000000, "2GB");
            _uc_scenario_run.uc_filesize_monitor4.Information = "Valid combinations for FVS Post-Treatment records";
           
            CompactMDB(m_strFVSPostValidComboDbPathAndFile, m_TempMDBFileConn);

			/**********************************************************************
			 **create valid combiniations of biosum_cond_id and treatment and 
			 **also user defined filters by owngrpcd
			 **********************************************************************/
			//insert all the possible valid POST plot+rxpackage+rxcycle records
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "--insert all possible valid post plot+rxpackage+rxcycle records--\r\n");
            //cycle1
			m_strSQL = "INSERT INTO validcombos_fvspost " + 
                       "SELECT a.biosum_cond_id,b.rxpackage,b.simyear1_rx AS rx,'1' AS rxcycle " + 
                       "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " + 
                       "WHERE b.simyear1_rx IS NOT NULL AND LEN(TRIM(b.simyear1_rx)) > 0 AND b.simyear1_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            //cycle2
            m_strSQL = "INSERT INTO validcombos_fvspost " +
                       "SELECT a.biosum_cond_id,b.rxpackage,b.simyear2_rx AS rx,'2' AS rxcycle " +
                       "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
                       "WHERE b.simyear2_rx IS NOT NULL AND LEN(TRIM(b.simyear2_rx)) > 0 AND b.simyear2_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            //cycle3
            m_strSQL = "INSERT INTO validcombos_fvspost " +
                       "SELECT a.biosum_cond_id,b.rxpackage,b.simyear3_rx AS rx,'3' AS rxcycle " +
                       "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
                       "WHERE b.simyear3_rx IS NOT NULL AND LEN(TRIM(b.simyear3_rx)) > 0 AND b.simyear3_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            //cycle4
            m_strSQL = "INSERT INTO validcombos_fvspost " +
           "SELECT a.biosum_cond_id,b.rxpackage,b.simyear4_rx AS rx,'4' AS rxcycle " +
           "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
           "WHERE b.simyear4_rx IS NOT NULL AND LEN(TRIM(b.simyear4_rx)) > 0 AND b.simyear4_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            CompactMDB(m_strFVSPostValidComboDbPathAndFile, m_TempMDBFileConn);

            //insert all the possible valid pre plot+rxpackage+rxcycle records
            CompactMDB(m_strFVSPreValidComboDbPathAndFile, m_TempMDBFileConn);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"--insert all possible valid PRE plot+rxpackage+rxcycle records--\r\n");
            //cycle1
            m_strSQL = "INSERT INTO validcombos_fvspre " +
           "SELECT a.biosum_cond_id,b.rxpackage,b.simyear1_rx AS rx,'1' AS rxcycle " +
           "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
           "WHERE b.simyear1_rx IS NOT NULL AND LEN(TRIM(b.simyear1_rx)) > 0 AND b.simyear1_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            //cycle2
            m_strSQL = "INSERT INTO validcombos_fvspre " +
                       "SELECT a.biosum_cond_id,b.rxpackage,b.simyear2_rx AS rx,'2' AS rxcycle " +
                       "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
                       "WHERE b.simyear2_rx IS NOT NULL AND LEN(TRIM(b.simyear2_rx)) > 0 AND b.simyear2_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            //cycle3
            m_strSQL = "INSERT INTO validcombos_fvspre " +
                       "SELECT a.biosum_cond_id,b.rxpackage,b.simyear3_rx AS rx,'3' AS rxcycle " +
                       "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
                       "WHERE b.simyear3_rx IS NOT NULL AND LEN(TRIM(b.simyear3_rx)) > 0 AND b.simyear3_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            //cycle4
            m_strSQL = "INSERT INTO validcombos_fvspre " +
           "SELECT a.biosum_cond_id,b.rxpackage,b.simyear4_rx AS rx,'4' AS rxcycle " +
           "FROM " + this.m_strCondTable + " a," + this.m_strRxPackageTable + " b " +
           "WHERE b.simyear4_rx IS NOT NULL AND LEN(TRIM(b.simyear4_rx)) > 0 AND b.simyear4_rx<>'000'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            CompactMDB(m_strFVSPreValidComboDbPathAndFile, m_TempMDBFileConn);
			

			string strWhere="";
			for (x=0;x<=FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				strTable = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.TableName(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.m_strPostVarArray[x]);
				if (strTable.Trim().Length > 0)
				{
					m_strSQL = "UPDATE validcombos_fvspost a INNER JOIN " + strTable + " b ON a.biosum_cond_id = b.biosum_cond_id AND a.rxpackage = b.rxpackage AND a.rx = b.rx AND a.rxcycle = b.rxcycle SET variable" + Convert.ToString(x + 1).Trim() + "_yn='Y'";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
					m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
					m_strSQL = "UPDATE validcombos_fvspost SET variable" + Convert.ToString(x + 1).Trim() + "_yn='N' WHERE variable" + Convert.ToString(x + 1).Trim() + "_yn IS NULL OR LEN(TRIM(variable" + Convert.ToString(x + 1).Trim() + "_yn))=0";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
					m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
					strWhere=strWhere + "b.variable" + Convert.ToString(x + 1).Trim() + "_yn <> 'N' AND ";

				}

				strTable = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.TableName(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario_fvs_prepost_variables_effective1.m_oSavVar.m_strPreVarArray[x]);
				if (strTable.Trim().Length > 0)
				{
                    m_strSQL = "UPDATE validcombos_fvspre a INNER JOIN " + strTable + " b ON a.biosum_cond_id = b.biosum_cond_id  AND a.rxpackage = b.rxpackage AND a.rx = b.rx AND a.rxcycle = b.rxcycle SET variable" + Convert.ToString(x + 1).Trim() + "_yn='Y'";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
					m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
					m_strSQL = "UPDATE validcombos_fvspre SET variable" + Convert.ToString(x + 1).Trim() + "_yn='N' WHERE variable" + Convert.ToString(x + 1).Trim() + "_yn IS NULL OR LEN(TRIM(variable" + Convert.ToString(x + 1).Trim() + "_yn))=0";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
					m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);
					strWhere=strWhere + "a.variable" + Convert.ToString(x + 1).Trim() + "_yn <> 'N' AND ";

				}
			}
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
			if (strWhere.Trim().Length > 0)
			{
				strWhere = "WHERE a.biosum_cond_id = b.biosum_cond_id AND a.rxpackage=b.rxpackage AND a.rx=b.rx AND a.rxcycle=b.rxcycle AND " +  strWhere.Substring(0,strWhere.Length - 5);
			}

			

			m_strSQL = "INSERT INTO validcombos_fvsprepost " + 
				"SELECT a.biosum_cond_id,a.rxpackage,a.rx,a.rxcycle " + 
				"FROM validcombos_fvspost a, validcombos_fvspre b " + 
				strWhere;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_strSQL);

            // Set the fvs_variant from the plot table
            m_strSQL = "UPDATE (validcombos_fvsprepost INNER JOIN " + this.m_strCondTable +
                       " ON validcombos_fvsprepost.biosum_cond_id = " + this.m_strCondTable + ".biosum_cond_id)" +
                       " INNER JOIN " + this.m_strPlotTable + " ON " + this.m_strCondTable + ".biosum_plot_id" +
                       " = " + this.m_strPlotTable + ".biosum_plot_id" +
                       " SET validcombos_fvsprepost.FVS_VARIANT = " + this.m_strPlotTable + ".fvs_variant";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

            m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);

            // Delete records from the validcombos_fvsprepost if they are in variant/packages on the filter's exclusion list
            //string strUnselectedPackages = _frmScenario.uc_optimizer_scenario_select_packages1.RxPackagesNotSelected;
            string strUnselectedPackages =
                (string)frmMain.g_oDelegate.GetControlPropertyValue((uc_optimizer_scenario_select_packages)_frmScenario.uc_optimizer_scenario_select_packages1, "RxPackagesNotSelected", false);
            if (! string.IsNullOrEmpty(strUnselectedPackages))
            {
                string[] strUnselectedPackagesArray = frmMain.g_oUtils.ConvertListToArray(strUnselectedPackages, ",");
                string strTempWhere = " WHERE FVS_VARIANT & RXPACKAGE IN ( ";
                foreach (string strVariantPackage in strUnselectedPackagesArray)
                {
                    strTempWhere = strTempWhere + "'" + strVariantPackage + "',";
                }
                string strTrimWhere = strTempWhere.TrimEnd(',') + ")";
                m_strSQL = "DELETE FROM validcombos_fvsprepost" + strTrimWhere;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

                m_ado.SqlNonQuery(this.m_TempMDBFileConn, m_strSQL);
            }

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            this.m_strSQL = "INSERT INTO validcombos (biosum_cond_id,rxpackage,rx,rxcycle) SELECT DISTINCT ruledefinitionscondfilter.biosum_cond_id,validcombos_fvsprepost.rxpackage,validcombos_fvsprepost.rx,validcombos_fvsprepost.rxcycle " +
				"FROM (((ruledefinitionsplotfilter INNER JOIN ruledefinitionscondfilter ON ruledefinitionsplotfilter.biosum_plot_id = ruledefinitionscondfilter.biosum_plot_id) " + 
				" INNER JOIN validcombos_fvsprepost ON ruledefinitionscondfilter.biosum_cond_id = validcombos_fvsprepost.biosum_cond_id) " +
                " INNER JOIN " + this.m_strHvstCostsTable + " ON " + 
                                  "(validcombos_fvsprepost.rxpackage=" + this.m_strHvstCostsTable + ".rxpackage) AND " + 
                                  "(validcombos_fvsprepost.rx = " + this.m_strHvstCostsTable + ".rx) AND " + 
                                  "(validcombos_fvsprepost.rxcycle = " + this.m_strHvstCostsTable + ".rxcycle) AND " + 
                                  "(validcombos_fvsprepost.biosum_cond_id = " + this.m_strHvstCostsTable + ".biosum_cond_id))";


			 if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\nCreate Valid Combinations\r\n");

             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                 frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            
			
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)==true) return;

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            this.m_strSQL = "DELETE v.* " + 
                            "FROM validcombos v WHERE " +
                            "EXISTS (SELECT * from " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " p " +
                            "WHERE v.biosum_cond_id = p.biosum_cond_id and (p.merch_haul_psite is null))";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nDelete combinations that don't have rows in the travel_times table\r\n");

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);


            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Delete inaccessible plots from Table validcombos\r\n");
            this.m_strSQL = "DELETE FROM validcombos v " +
                            "WHERE EXISTS " +
                             "(SELECT p.* " +
                             "FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " p " +
                             "WHERE v.biosum_cond_id =  p.biosum_cond_id and p.COND_ACCESSIBLE_YN = 'N')";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

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

                //
                //create an audit DB file for every variant
                //
                    
                string strAccdbPathAndFile = ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() + 
                    "\\db\\" + Tables.Audit.DefaultCondAuditTableDbFile;
                _uc_scenario_run.uc_filesize_monitor4.BeginMonitoringFile(strAccdbPathAndFile, 2000000000, "2GB");

                //BIOSUM_COND_ID RECORD AUDIT
                /******************************************************************************************
                 **insert all the plots that are being processed into the plot audit table
                 ******************************************************************************************/
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n--cond_audit--\r\n\r\n");
                this.m_strSQL = "INSERT INTO " + Tables.Audit.DefaultCondAuditTableName + " (biosum_cond_id) SELECT ruledefinitionscondfilter.biosum_cond_id FROM ruledefinitionscondfilter INNER JOIN userdefinedplotfilter_work ON ruledefinitionscondfilter.biosum_plot_id = userdefinedplotfilter_work.biosum_plot_id";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Insert All Plots From ruledefinitionscoldfilter table into cond_audit\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                if (this.m_ado.m_intError != 0)
                {
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                    this.m_intError = this.m_ado.m_intError;
                    return;
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();


                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                /************************************************************************
                 **check to see if the plot record exists in the frcs harvest cost table
                 ************************************************************************/
                m_ado.m_strSQL="SELECT COUNT(*) FROM " + 
                                    "(SELECT TOP 1 a.biosum_cond_id " + 
                                    "FROM " + Tables.Audit.DefaultCondAuditTableName + " AS a," + 
                                            m_strHvstCostsTable + " b " + 
                                    "WHERE a.biosum_cond_id = b.biosum_cond_id)";
                if ((int)m_ado.getRecordCount(m_TempMDBFileConn, m_ado.m_strSQL,"temp") > 0)
                {
                    this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + " AS a " +
                                    "SET a.harvest_costs_yn = 'Y' " +
                                    "WHERE a.biosum_cond_id " +
                                    "IN (SELECT biosum_cond_id FROM " + this.m_strHvstCostsTable + ");";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSee if cond record exists in the harvest_costs table\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + " a " + 
                                "SET a.harvest_costs_yn = 'N' " +
                                "WHERE a.harvest_costs_yn IS NULL OR LEN(TRIM(a.harvest_costs_yn))=0;";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet harvest_costs_yn=N if column value is null\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;

                /****************************************************************************
                 **check to see if the plot record exists in the validcombos_fvsprepost table
                 ****************************************************************************/
                 m_ado.m_strSQL="SELECT COUNT(*) FROM " + 
                                    "(SELECT TOP 1 a.biosum_cond_id " + 
                                    "FROM " + Tables.Audit.DefaultCondAuditTableName + " AS a," + 
                                          "validcombos_fvsprepost b " + 
                                    "WHERE a.biosum_cond_id = b.biosum_cond_id)";
                 if ((int)m_ado.getRecordCount(m_TempMDBFileConn, m_ado.m_strSQL, "temp") > 0)
                 {
                     this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + " a " +
                                     "SET a.fvs_prepost_variables_yn = 'Y' " +
                                     "WHERE a.biosum_cond_id " +
                                     "IN (SELECT biosum_cond_id FROM validcombos_fvsprepost);";
                     if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                         frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet fvs_prepost_variables_yn=Y if plot record exists in validcombos_fvsprepost table\r\n");
                     if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                         frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                     this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                 }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


                this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + " a " + 
                                "SET a.fvs_prepost_variables_yn = 'N' " +
                                "WHERE a.fvs_prepost_variables_yn IS NULL OR " + 
                                      "a.fvs_prepost_variables_yn <> 'Y' ;";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet fvs_prepost_variables_yn=N if column value is null\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


                /********************************************************************************************************
                 **check to see if the plot record exists in the processor tree volume and value tableharvest cost table
                 ********************************************************************************************************/
                  m_ado.m_strSQL="SELECT COUNT(*) FROM " + 
                                    "(SELECT TOP 1 a.biosum_cond_id " + 
                                    "FROM " + Tables.Audit.DefaultCondAuditTableName + " AS a," + 
                                          this.m_strTreeVolValSumTable + " b " + 
                                    "WHERE a.biosum_cond_id = b.biosum_cond_id)";
                  if ((int)m_ado.getRecordCount(m_TempMDBFileConn, m_ado.m_strSQL, "temp") > 0)
                  {
                      this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + " a " +
                                      "SET a.processor_tree_vol_val_yn = 'Y' " +
                                      "WHERE a.biosum_cond_id " +
                                      "IN (SELECT biosum_cond_id FROM " + this.m_strTreeVolValSumTable + ");";
                      if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                          frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet processor_tree_vol_val_yn=Y if plot record exists in " + this.m_strTreeVolValSumTable + " table\r\n");
                      if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                          frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                      this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                  }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


                this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + " a " + 
                                "SET a.processor_tree_vol_val_yn = 'N' " +
                                "WHERE a.processor_tree_vol_val_yn IS NULL OR  " + 
                                      "a.processor_tree_vol_val_yn<>'Y' ;";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet processor_tree_vol_val_yn=N if column value is null\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;



                /**********************************************************************
                 **check to see if the plot record exists in the gis travel times table
                 **********************************************************************/
                 m_ado.m_strSQL="SELECT COUNT(*) FROM " + 
                                    "(SELECT TOP 1 a.biosum_cond_id " + 
                                    "FROM " + Tables.Audit.DefaultCondAuditTableName + " AS a," + 
                                         "ruledefinitionscondfilter b," + 
                                          m_strTravelTimeTable + " c " + 
                                    "WHERE a.biosum_cond_id = b.biosum_cond_id AND c.biosum_plot_id=b.biosum_plot_id)";
                 if ((int)m_ado.getRecordCount(m_TempMDBFileConn, m_ado.m_strSQL, "temp") > 0)
                 {
                     this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + " a " +
                                     "SET a.gis_travel_times_yn = 'Y' " +
                                     "WHERE a.biosum_cond_id " +
                                     "IN (SELECT biosum_cond_id FROM ruledefinitionscondfilter " +
                                         "WHERE ruledefinitionscondfilter.biosum_plot_id " +
                                                "IN (SELECT biosum_plot_id FROM " +
                                                     this.m_strTravelTimeTable + "));";
                     if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                         frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet gis_travel_times_yn=Y if plot record exists in " + this.m_strTravelTimeTable + " table\r\n");
                     if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                         frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                     this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                 }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


                this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + " a " + 
                                "SET a.gis_travel_times_yn = 'N' " +
                                "WHERE a.gis_travel_times_yn IS NULL OR  " + 
                                      "a.gis_travel_times_yn<>'Y' ;";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet gis_travel_times_yn=N if column value is null\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


                this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondAuditTableName + " a " +
                "INNER JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " p " +
                "ON a.biosum_cond_id = p.biosum_cond_id " +
                "SET a.cond_too_far_steep_yn = p.cond_too_far_steep_yn, " +
                "a.psite_merch_yn = IIF(p.merch_haul_psite IS NULL,'N','Y'), " +
                "a.psite_chip_yn = IIF(p.chip_haul_psite IS NULL,'N','Y') ";
                
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet accessibility and psite flags\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


                    //BIOSUM_COND_ID + RX RECORD AUDIT
                    /**********************************************************************************
                    **Insert all the biosum_cond_id + rx combinations into the cond_rx_audit table
                    ***********************************************************************************/
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n--cond_rx_audit--\r\n");
                    //cycle1
                    this.m_strSQL = "INSERT INTO " + Tables.Audit.DefaultCondRxAuditTableName + 
                         "(biosum_cond_id,rxpackage,rx,rxcycle)  " +
                         "SELECT a.biosum_cond_id, b.rxpackage,b.rx,b.rxcycle " +
                         "FROM " + Tables.Audit.DefaultCondAuditTableName + " a, " +
                        "(SELECT DISTINCT rxpackage,simyear1_rx AS rx,'1' AS rxcycle " +
                         "FROM " + this.m_strRxPackageTable + " " +
                         "WHERE simyear1_rx IS NOT NULL AND LEN(TRIM(simyear1_rx)) > 0 AND simyear1_rx<>'000') b ;";  //+ 
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                    //cycle2
                    this.m_strSQL = "INSERT INTO " + Tables.Audit.DefaultCondRxAuditTableName + 
                         "(biosum_cond_id,rxpackage,rx,rxcycle)  " +
                         "SELECT a.biosum_cond_id, b.rxpackage,b.rx,b.rxcycle " +
                         "FROM " + Tables.Audit.DefaultCondAuditTableName + " a, " +
                        "(SELECT DISTINCT rxpackage,simyear2_rx AS rx,'2' AS rxcycle " +
                         "FROM " + this.m_strRxPackageTable + " " +
                         "WHERE simyear2_rx IS NOT NULL AND LEN(TRIM(simyear2_rx)) > 0 AND simyear2_rx<>'000') b ;";  //+ 
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                    //cycle3
                    this.m_strSQL = "INSERT INTO " + Tables.Audit.DefaultCondRxAuditTableName + 
                        "(biosum_cond_id,rxpackage,rx,rxcycle)  " +
                         "SELECT a.biosum_cond_id, b.rxpackage,b.rx,b.rxcycle " +
                         "FROM " + Tables.Audit.DefaultCondAuditTableName + " a, " +
                        "(SELECT DISTINCT rxpackage,simyear3_rx AS rx,'3' AS rxcycle " +
                         "FROM " + this.m_strRxPackageTable + " " +
                         "WHERE simyear3_rx IS NOT NULL AND LEN(TRIM(simyear3_rx)) > 0 AND simyear3_rx<>'000') b ;";  //+ 
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                    //cycle4
                    this.m_strSQL = "INSERT INTO " + Tables.Audit.DefaultCondRxAuditTableName + 
                         "(biosum_cond_id,rxpackage,rx,rxcycle)  " +
                         "SELECT a.biosum_cond_id, b.rxpackage,b.rx,b.rxcycle " +
                         "FROM " + Tables.Audit.DefaultCondAuditTableName + " a, " +
                        "(SELECT DISTINCT rxpackage,simyear4_rx AS rx,'4' AS rxcycle " +
                         "FROM " + this.m_strRxPackageTable + " " +
                         "WHERE simyear4_rx IS NOT NULL AND LEN(TRIM(simyear4_rx)) > 0 AND simyear4_rx<>'000') b ;";  //+ 
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);



                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


                    /*********************************************************************************
                     **check to see if the cond + rx record exists in the fvs prepost variables table
                     *********************************************************************************/
                    this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondRxAuditTableName + 
                        " SET fvs_prepost_variables_yn = 'Y' " +
                        "WHERE EXISTS (SELECT biosum_cond_id,rxpackage,rx,rxcycle " +
                        "FROM validcombos_fvsprepost " +
                        "WHERE " + Tables.Audit.DefaultCondRxAuditTableName + ".biosum_cond_id = " +
                        "validcombos_fvsprepost.biosum_cond_id AND " +
                        Tables.Audit.DefaultCondRxAuditTableName + ".rxpackage = validcombos_fvsprepost.rxpackage AND " +
                        Tables.Audit.DefaultCondRxAuditTableName + ".rx = validcombos_fvsprepost.rx AND " +
                        Tables.Audit.DefaultCondRxAuditTableName + ".rxcycle=validcombos_fvsprepost.rxcycle);";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet fvs_prepost_variables_yn=Y if plot + rx + rxpackage + rxcycle record exists in validcombos_fvsprepost table\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


                    this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondRxAuditTableName + " a " + 
                        "SET a.fvs_prepost_variables_yn = 'N' " +
                        "WHERE a.fvs_prepost_variables_yn IS NULL OR LEN(TRIM(a.fvs_prepost_variables_yn))=0;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet fvs_prepost_variables_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


                    /****************************************************************************
                     **check to see if the plot + rx record exists in the harvest costs table
                     ****************************************************************************/
                    this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondRxAuditTableName + 
                        " SET harvest_costs_yn = 'Y' " +
                        "WHERE EXISTS (SELECT biosum_cond_id,rxpackage,rx,rxcycle " +
                        "FROM " + this.m_strHvstCostsTable + " " +
                        "WHERE " + Tables.Audit.DefaultCondRxAuditTableName + ".biosum_cond_id = " +
                        this.m_strHvstCostsTable.Trim() + ".biosum_cond_id AND " +
                        Tables.Audit.DefaultCondRxAuditTableName + ".rxpackage = " + this.m_strHvstCostsTable.Trim() + ".rxpackage AND " +
                        Tables.Audit.DefaultCondRxAuditTableName + ".rx = " + this.m_strHvstCostsTable.Trim() + ".rx AND " +
                        Tables.Audit.DefaultCondRxAuditTableName + ".rxcycle = " + this.m_strHvstCostsTable.Trim() + ".rxcycle);";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet harvest_costs_yn=Y if plot + rx + rxpackage + rxcycle record exists in " + m_strHvstCostsTable + " table\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


                    this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondRxAuditTableName + " a " + 
                        "SET a.harvest_costs_yn = 'N' " +
                        "WHERE a.harvest_costs_yn IS NULL OR LEN(TRIM(a.harvest_costs_yn))=0 ;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet harvest_costs_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


                    /*********************************************************************************
                     **check to see if the plot + rx record exists in the processor tree vol val table
                     *********************************************************************************/
                    m_ado.m_strSQL="SELECT COUNT(*) FROM " + 
                                        "(SELECT TOP 1 a.biosum_cond_id " +
                                        "FROM " + Tables.Audit.DefaultCondRxAuditTableName + " a," + 
                                             m_strTreeVolValSumTable + " b " + 
                                        "WHERE a.biosum_cond_id = b.biosum_cond_id AND " + 
                                              "a.rxpackage = b.rxpackage AND " + 
                                              "a.rx=b.rx AND " + 
                                              "a.rxcycle=b.rxcycle)";
                    if ((int)m_ado.getRecordCount(m_TempMDBFileConn, m_ado.m_strSQL, "temp") > 0)
                    {


                        this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondRxAuditTableName + " a " +
                                       "INNER JOIN " + m_strTreeVolValSumTable + " b " +
                                       "ON  a.biosum_cond_id=b.biosum_cond_id AND " +
                                           "a.rxpackage=b.rxpackage AND " +
                                           "a.rx=b.rx AND " +
                                           "a.rxcycle=b.rxcycle " +
                                       "SET a.processor_tree_vol_val_yn = 'Y'";
                       
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet processor_tree_vol_val_yn=Y if plot + rx + rxpackage + rxcycle record exists in " + m_strTreeVolValSumTable + " table\r\n");
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                        this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                    }
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


                    this.m_strSQL = "UPDATE " + Tables.Audit.DefaultCondRxAuditTableName + " a " + 
                        "SET a.processor_tree_vol_val_yn = 'N' " +
                        "WHERE a.processor_tree_vol_val_yn IS NULL OR LEN(TRIM(a.processor_tree_vol_val_yn))=0;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nSet processor_tree_vol_val_yn=N if column value is null\r\n");
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


                    

                    //compact
                    CompactMDB(ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioPath.Text.Trim() 
                        + "\\db\\" + Tables.Audit.DefaultCondAuditTableDbFile, m_TempMDBFileConn);
				

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
			int x,y;
			string strPreTable="";
			string strPreColumn="";
			string strPostTable="";
			string strPostColumn="";
			string[] strEffectiveColumnArray;
			string[] strBetterIsNotNull= new string[uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES];
			string[] strWorseIsNotNull= new string[uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES];
			string[] strEffectiveIsNotNull= new string[uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES];
			string strOverallEffectiveIsNotNull = "";
			string strBetterSql="";
			string strWorseSql="";
			string strEffectiveSql="";
            int intListViewIndex = -1;

			string strVariableNumber="";
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


			//get all the column names in the effective table
            string strEffectiveTableName = ReferenceOptimizerScenarioForm.OutputTablePrefix + Tables.OptimizerScenarioResults.DefaultScenarioResultsEffectiveTableSuffix;
            strEffectiveColumnArray = m_ado.getFieldNamesArray(this.m_TempMDBFileConn,"select * from " + strEffectiveTableName);

			/********************************************
			 **delete all records in the effective table
			 ********************************************/
			this.m_strSQL = "delete from " + strEffectiveTableName;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			//insert the valid combos into the effective table
			m_strSQL = "INSERT INTO " + strEffectiveTableName + " (biosum_cond_id,rxpackage,rx,rxcycle) SELECT biosum_cond_id,rxpackage,rx,rxcycle FROM validcombos WHERE rxcycle='1'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			//insert net revenue per acre into the effective table
			m_strSQL = "UPDATE " + strEffectiveTableName + " e " + 
				       "INNER JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName +" p " + 
				       "ON e.biosum_cond_id=p.biosum_cond_id AND " + 
                          "e.rxpackage=p.rxpackage AND " + 
				          "e.rx=p.rx AND " + 
                          "e.rxcycle=p.rxcycle " + 
			           "SET e.nr_dpa = IIF(p.max_nr_dpa IS NOT NULL,p.max_nr_dpa,0)";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

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
                        m_strSQL = "UPDATE " + strEffectiveTableName + " e " +
                                    "INNER JOIN " + strDatabase[0] + " p " +
                                    "ON e.biosum_cond_id=p.biosum_cond_id AND " +
                                    "e.rxpackage=p.rxpackage " +
                                    "SET e." + this.m_oOptimizationVariable.strRevenueAttribute + " = IIF(p." + strDatabase[1] + " IS NOT NULL,p." + strDatabase[1] + ",0)";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                        this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                    }
                }
            }


            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

        
			//populate the variable table.column name and its value to the effective table
			for (x=0;x<=uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				
				
				m_strSQL="";
				strPreTable="";
				strPreColumn="";
				strPostTable="";
				strPostColumn="";
				strVariableNumber = Convert.ToString(x+1).Trim();
				if (oFvsVar.TableName(oFvsVar.m_strPreVarArray[x].Trim().ToUpper()) != "NOT DEFINED")
				{
					
					strPreTable=oFvsVar.TableName(oFvsVar.m_strPreVarArray[x].Trim());
					strPreColumn=oFvsVar.ColumnName(oFvsVar.m_strPreVarArray[x].Trim());
				}
				if (oFvsVar.m_strPostVarArray[x].Trim().ToUpper() != "NOT DEFINED")
				{
					
					strPostTable=oFvsVar.TableName(oFvsVar.m_strPostVarArray[x].Trim());
					strPostColumn=oFvsVar.ColumnName(oFvsVar.m_strPostVarArray[x].Trim());
				}
				if (strPreTable.Trim().Length > 0)
				{
					m_strSQL = "e.pre_variable" + strVariableNumber + "_name='" + strPreTable + "." + strPreColumn + "',";
					m_strSQL = m_strSQL + "e.pre_variable" + strVariableNumber + "_value=pre." + strPreColumn;
				}
				else
				{
					m_strSQL = "e.pre_variable" + strVariableNumber + "_name=null,";
					m_strSQL = m_strSQL + "e.pre_variable" + strVariableNumber + "_value=null";
				}

				if (strPostTable.Trim().Length > 0)
				{
					m_strSQL = m_strSQL + ",e.post_variable" + strVariableNumber + "_name='" + strPostTable + "." + strPostColumn + "',";
					m_strSQL = m_strSQL + "e.post_variable" + strVariableNumber + "_value=post." + strPostColumn;
				}
				else
				{
					m_strSQL = m_strSQL + ",e.post_variable" + strVariableNumber + "_name=null,";
					m_strSQL = m_strSQL + "e.post_variable" + strVariableNumber + "_value=null";
				}
				if (strPreTable.Trim().Length > 0 && strPostTable.Trim().Length > 0)
				{
					m_strSQL = "UPDATE " + strEffectiveTableName + " e " + 
                               "INNER JOIN (" + strPostTable + " post " + 
                               "INNER JOIN " + 	strPreTable  + " pre " + 
                               "ON post.biosum_cond_id = pre.biosum_cond_id AND " + 
                                  "post.rxpackage=pre.rxpackage AND " + 
                                  "post.rx=pre.rx AND " + 
                                  "post.rxcycle=pre.rxcycle) "  + 
						       "ON e.biosum_cond_id=post.biosum_cond_id AND " + 
                                  "e.rxpackage=post.rxpackage AND " + 
                                  "e.rx=post.rx AND " + 
                                  "e.rxcycle=post.rxcycle " + 
						       "SET " + m_strSQL;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

				}
				

			}

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
           
			//populate the change column by subtracting pre value from post value
			m_strSQL="";
			for (x=0;x<=uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				strVariableNumber = Convert.ToString(x+1).Trim();
				m_strSQL = m_strSQL + "variable" + strVariableNumber + "_change=IIF(pre_variable" + strVariableNumber + "_value IS NOT NULL AND post_variable" + strVariableNumber + "_value IS NOT NULL,post_variable" + strVariableNumber + "_value - pre_variable" + strVariableNumber + "_value,null),";
			}
			m_strSQL = m_strSQL.Substring(0,m_strSQL.Length - 1);

			m_strSQL = "UPDATE " + strEffectiveTableName + " SET " + m_strSQL;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            

			//see what variables are referenced in the sql expression and make sure they are not null
			strOverallEffectiveIsNotNull="";
			for (x=0;x<=uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				strBetterIsNotNull[x]="";
				strWorseIsNotNull[x]="";
				strEffectiveIsNotNull[x]="";

				for (y=0;y<=strEffectiveColumnArray.Length - 1;y++)
				{
					if (oFvsVar.m_strBetterExpr[x].Trim().Length > 0)
					{
						if (oFvsVar.m_strBetterExpr[x].Trim().ToUpper().IndexOf(strEffectiveColumnArray[y].ToUpper(),0) >= 0)
						{
							strBetterIsNotNull[x] = strBetterIsNotNull[x] + strEffectiveColumnArray[y] + " IS NOT NULL AND ";
						}
					}
					if (oFvsVar.m_strWorseExpr[x].Trim().Length > 0)
					{
						if (oFvsVar.m_strWorseExpr[x].Trim().ToUpper().IndexOf(strEffectiveColumnArray[y].ToUpper(),0) >= 0)
						{
							strWorseIsNotNull[x] = strWorseIsNotNull[x] + strEffectiveColumnArray[y] + " IS NOT NULL AND ";
						}
					}
					if (oFvsVar.m_strEffectiveExpr[x].Trim().Length > 0)
					{
						if (oFvsVar.m_strEffectiveExpr[x].Trim().ToUpper().IndexOf(strEffectiveColumnArray[y].ToUpper(),0) >= 0)
						{
							strEffectiveIsNotNull[x] = strEffectiveIsNotNull[x] + strEffectiveColumnArray[y] + " IS NOT NULL AND ";
						}
					}
				}

			}
			if (oFvsVar.m_strOverallEffectiveExpr.Trim().Length > 0)
			{
				for (y=0;y<=strEffectiveColumnArray.Length - 1;y++)
				{
					
					if (oFvsVar.m_strOverallEffectiveExpr.Trim().ToUpper().IndexOf(strEffectiveColumnArray[y].ToUpper(),0) >= 0)
					{
						strOverallEffectiveIsNotNull = strOverallEffectiveIsNotNull + strEffectiveColumnArray[y] + " IS NOT NULL AND ";
					}
					
				}
			}
			//remove the last AND
			for (x=0;x<=uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				if (strBetterIsNotNull[x].Trim().Length > 0)
				{
					strBetterIsNotNull[x] = strBetterIsNotNull[x].Substring(0,strBetterIsNotNull[x].Length - 5);
				}
				if (strWorseIsNotNull[x].Trim().Length > 0)
				{
					strWorseIsNotNull[x] = strWorseIsNotNull[x].Substring(0,strWorseIsNotNull[x].Length - 5);
				}
				if (strEffectiveIsNotNull[x].Trim().Length > 0)
				{
					strEffectiveIsNotNull[x] = strEffectiveIsNotNull[x].Substring(0,strEffectiveIsNotNull[x].Length - 5);
				}
			}
			if (strOverallEffectiveIsNotNull.Trim().Length > 0)
			{
				strOverallEffectiveIsNotNull = strOverallEffectiveIsNotNull.Substring(0,strOverallEffectiveIsNotNull.Length - 5);
			}
			//populate the better,worse,effective, and overall effective columns
			m_strSQL="";
			strBetterSql="";
			strWorseSql="";
			strEffectiveSql="";
			for (x=0;x<=uc_optimizer_scenario_fvs_prepost_variables_effective.NUMBER_OF_VARIABLES-1;x++)
			{
				strVariableNumber = Convert.ToString(x+1).Trim();
				if (oFvsVar.m_strBetterExpr[x].Trim().Length > 0)
				{
					strBetterSql = strBetterSql + "variable" + strVariableNumber + "_better_yn=IIF(" + strBetterIsNotNull[x].Trim() + ",IIF(" + oFvsVar.m_strBetterExpr[x].Trim() + ",'Y','N'),null),";
				}
				if (oFvsVar.m_strWorseExpr[x].Trim().Length > 0)
				{
					strWorseSql = strWorseSql + "variable" + strVariableNumber + "_worse_yn=IIF(" + strWorseIsNotNull[x].Trim() + ",IIF(" + oFvsVar.m_strWorseExpr[x].Trim() + ",'Y','N'),null),";
				}
				if (oFvsVar.m_strEffectiveExpr[x].Trim().Length > 0)
				{
					strEffectiveSql = strEffectiveSql + "variable" + strVariableNumber + "_effective_yn=IIF(" + strEffectiveIsNotNull[x].Trim() + ",IIF(" + oFvsVar.m_strEffectiveExpr[x].Trim() + ",'Y','N'),null),";
				}
			}
			// Mark effective treatments with a 'Y'
            m_strSQL = m_strSQL + "overall_effective_yn=IIF(" + strOverallEffectiveIsNotNull.Trim() + ",IIF(" + oFvsVar.m_strOverallEffectiveExpr.Trim() + ",'Y','N'),null),";

			//better
            if (strBetterSql.Trim().Length > 0)
            {
                strBetterSql = strBetterSql.Substring(0, strBetterSql.Length - 1);
                strBetterSql = "UPDATE " + strEffectiveTableName + " SET " + strBetterSql;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Improvement--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strBetterSql + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strBetterSql);

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
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, strWorseSql);
                
            }

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			//effective
			strEffectiveSql = strEffectiveSql.Substring(0,strEffectiveSql.Length - 1);
			strEffectiveSql = "UPDATE " + strEffectiveTableName + " SET " + strEffectiveSql;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "--Variable Effective--\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strEffectiveSql + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strEffectiveSql);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            
			//overall effective
			m_strSQL = m_strSQL.Substring(0,m_strSQL.Length - 1);
			m_strSQL = "UPDATE " + strEffectiveTableName + " SET " + m_strSQL;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "--Overall Effective--\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

           
            

			
			if (this.m_ado.m_intError != 0)
			{
                 if (frmMain.g_bDebug)  frmMain.g_oUtils.WriteText(m_strDebugFile,"!!!Error Executing SQL!!!\r\n");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (Convert.ToInt32(m_ado.getRecordCount(m_TempMDBFileConn,"SELECT COUNT(*) FROM " + strEffectiveTableName + " WHERE overall_effective_yn='Y'","temp"))==0)
            {
                
                    MessageBox.Show("No overall effective treatments found. Processing is cancelled");
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Cancelled!!");
                    ReferenceUserControlScenarioRun.m_bUserCancel = true;
                    return;
                
            }

            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			
            
            
			if (this.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
			}
            


		}

		private void Optimization()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Optimization\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			
			string strPreTable="";
			string strPreColumn="";
			string strPostTable="";
			string strPostColumn="";


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


			/********************************************
			 **delete all records in the optimization table
			 ********************************************/
            string strOptimizationTableName = ReferenceOptimizerScenarioForm.OutputTablePrefix +
                Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix;
			this.m_strSQL = "delete from " + strOptimizationTableName;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			//insert the valid combos into the optimization table
            this.m_strSQL = "INSERT INTO " + strOptimizationTableName + " (biosum_cond_id,rxpackage,rx,rxcycle,affordable_YN";
            if (this.m_oOptimizationVariable.bUseFilter == true)
                this.m_strSQL = this.m_strSQL + "," + this.m_oOptimizationVariable.strRevenueAttribute;
            this.m_strSQL = this.m_strSQL + ") SELECT biosum_cond_id,rxpackage,rx,rxcycle,'Y' ";
            if (this.m_oOptimizationVariable.bUseFilter == true)
                this.m_strSQL = this.m_strSQL + "," + this.m_oOptimizationVariable.strRevenueAttribute;
            this.m_strSQL = this.m_strSQL + " FROM " + ReferenceOptimizerScenarioForm.OutputTablePrefix +
                Tables.OptimizerScenarioResults.DefaultScenarioResultsEffectiveTableSuffix + 
                " WHERE overall_effective_yn='Y'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nOptimization Type: " + this.m_oOptimizationVariable.strOptimizedVariable.ToUpper() + "\r\n");
            
			//populate the variable table.column name and its value to the optimization table
			if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "REVENUE")
			{
                m_strSQL = "UPDATE " + strOptimizationTableName + " e " +
					"LEFT OUTER JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName + " p " + 
					"ON e.biosum_cond_id=p.biosum_cond_id AND " + 
                    "e.rxpackage=p.rxpackage AND " + 
                    "e.rx=p.rx AND " + 
                    "e.rxcycle = p.rxcycle " + 
					"SET e.pre_variable_name = 'p.max_nr_dpa'," + 
					    "e.post_variable_name = 'p.max_nr_dpa'," + 
					    "e.pre_variable_value = IIF(p.max_nr_dpa IS NOT NULL,p.max_nr_dpa,0)," + 
					    "e.post_variable_value = IIF(p.max_nr_dpa IS NOT NULL,p.max_nr_dpa,0)," + 
					    "e.change_value = 0";																									 
			}
			else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "MERCHANTABLE VOLUME")
			{
                m_strSQL = "UPDATE " + strOptimizationTableName + " e " + 
					"LEFT OUTER JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName + " p " + 
					"ON e.biosum_cond_id=p.biosum_cond_id AND " +
                    "e.rxpackage=p.rxpackage AND " +
                    "e.rx=p.rx AND " +
                    "e.rxcycle = p.rxcycle " + 
					"SET e.pre_variable_name = 'p.merch_vol_cf'," + 
					"e.post_variable_name = 'p.merch_vol_cf'," + 
					"e.pre_variable_value = IIF(p.merch_vol_cf IS NOT NULL,p.merch_vol_cf,0)," + 
					"e.post_variable_value = IIF(p.merch_vol_cf IS NOT NULL,p.merch_vol_cf,0)," + 
					"e.change_value = 0";

			}
            else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "ECONOMIC ATTRIBUTE")
            {
                m_strSQL = getEconomicOptimizationSql();

            }
			else
			{
						
				string[] strCol = frmMain.g_oUtils.ConvertListToArray(this.m_oOptimizationVariable.strFVSVariableName,".");
				strPreTable="PRE_" + strCol[0].Trim();
				strPreColumn=strCol[1].Trim();
				strPostTable="POST_" + strCol[0].Trim();
				strPostColumn=strCol[1].Trim();

				m_strSQL = "UPDATE " + strOptimizationTableName + " e " + 
						"INNER JOIN (" + strPostTable + " a " + 
						"INNER JOIN " + strPreTable + " b " + 
						"ON a.biosum_cond_id=b.biosum_cond_id AND " +
                           "a.rxpackage=b.rxpackage AND " + 
                           "a.rx = b.rx AND " + 
                           "a.rxcycle = b.rxcycle) " + 
						"ON e.biosum_cond_id=a.biosum_cond_id AND " +
                        "e.rxpackage=a.rxpackage AND " +
                        "e.rx=a.rx AND " +
                        "e.rxcycle = a.rxcycle " + 
						"SET e.change_value=" + 
								"IIF(a." + strPostColumn + " IS NOT NULL AND b." + strPreColumn + " IS NOT NULL," + 
									"a." + strPostColumn + " - b." + strPreColumn + "," + 
								"IIF(a." + strPostColumn + " IS NOT NULL," + 
									"a." + strPostColumn + "," + 
								"IIF(b." + strPreColumn  + " IS NOT NULL,0-b." + strPreColumn + ",0)))," + 
                           "e.pre_variable_name='Pre_" + this.m_oOptimizationVariable.strFVSVariableName.Trim() + "'," + 
						   "e.post_variable_name='Post_" + this.m_oOptimizationVariable.strFVSVariableName.Trim() + "'," + 
						   "e.pre_variable_value=IIF(b." + strPreColumn + " IS NOT NULL,b." + strPreColumn + ",0)," + 
						   "e.post_variable_value=IIF(a." + strPostColumn + " IS NOT NULL,a." + strPostColumn + ",0)";
			}
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            // Update affordable flag for revenue filter
            if (this.m_oOptimizationVariable.bUseFilter == true)
            {
                m_strSQL = "UPDATE " + strOptimizationTableName + " e " +
                           "SET e.affordable_YN = IIF(e." + this.m_oOptimizationVariable.strRevenueAttribute +
                           " " + this.m_oOptimizationVariable.strFilterOperator + " " + this.m_oOptimizationVariable.dblFilterValue + ",'Y','N')";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            }

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");

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

		private void tiebreaker()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//tiebreaker\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			string strPreTable="";
			string strPreColumn="";
			string strPostTable="";
			string strPostColumn="";
			string[] strArray=null;
			string strVariableNumber="";

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


           


			/********************************************
			 **delete all records in the tie breaker table
			 ********************************************/
			this.m_strSQL = "delete from tiebreaker";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			//insert the valid combos into the tiebreaker table
            string strOptimizationTableName = ReferenceOptimizerScenarioForm.OutputTablePrefix +
                Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix;
			m_strSQL = "INSERT INTO tiebreaker (biosum_cond_id,rxpackage,rx,rxcycle) " + 
				          "SELECT biosum_cond_id,rxpackage,rx,rxcycle FROM " + strOptimizationTableName + " WHERE affordable_YN='Y'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			//populate the variable table.column name and its value to the tiebreaker table
			oItem = oTieBreakerCollection.Item(0);  // STAND ATTRIBUTE
			if (oItem.bSelected)
			{
				strArray=frmMain.g_oUtils.ConvertListToArray(oItem.strFVSVariableName,".");

				strPreTable="PRE_" + strArray[0].Trim();
				strPreColumn=strArray[1].Trim();
				strPostTable="POST_" + strArray[0].Trim();
				strPostColumn=strArray[1].Trim();
				strVariableNumber="1";
				if (strPreTable.Trim().Length > 0)
				{
					m_strSQL = "e.pre_variable" + strVariableNumber + "_name='" + strPreTable + "." + strPreColumn + "',";
					m_strSQL = m_strSQL + "e.pre_variable" + strVariableNumber + "_value=pre." + strPreColumn;
				}
				else
				{
					m_strSQL = "e.pre_variable" + strVariableNumber + "_name=null,";
					m_strSQL = m_strSQL + "e.pre_variable" + strVariableNumber + "_value=null";
				}

				if (strPostTable.Trim().Length > 0)
				{
					m_strSQL = m_strSQL + ",e.post_variable" + strVariableNumber + "_name='" + strPostTable + "." + strPostColumn + "',";
					m_strSQL = m_strSQL + "e.post_variable" + strVariableNumber + "_value=post." + strPostColumn;
				}
				else
				{
					m_strSQL = m_strSQL + ",e.post_variable" + strVariableNumber + "_name=null,";
					m_strSQL = m_strSQL + "e.post_variable" + strVariableNumber + "_value=null";
				}
				if (strPreTable.Trim().Length > 0 && strPostTable.Trim().Length > 0)
				{
					m_strSQL = "UPDATE tiebreaker e INNER JOIN (" + strPostTable + " post INNER JOIN " + 
						strPreTable  + " pre ON post.biosum_cond_id = pre.biosum_cond_id AND " + 
                                               "post.rxpackage = pre.rxpackage AND " + 
                                               "post.rx = pre.rx AND " + 
                                               "post.rxcycle = pre.rxcycle) "  + 
						"ON e.biosum_cond_id=post.biosum_cond_id AND " + 
                           "e.rxpackage=post.rxpackage AND  " + 
                           "e.rx=post.rx AND " + 
                           "e.rxcycle=post.rxcycle " + 
						"SET " + m_strSQL;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
					this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

				}

				//populate the change column by subtracting pre value from post value
				m_strSQL="";
				m_strSQL = m_strSQL + "variable" + strVariableNumber + "_change=IIF(pre_variable" + strVariableNumber + "_value IS NOT NULL AND post_variable" + strVariableNumber + "_value IS NOT NULL,post_variable" + strVariableNumber + "_value - pre_variable" + strVariableNumber + "_value,null),";
				m_strSQL = m_strSQL.Substring(0,m_strSQL.Length - 1);

				m_strSQL = "UPDATE tiebreaker SET " + m_strSQL;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			
				
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
                            m_strSQL = "UPDATE tiebreaker e " +
                                     "INNER JOIN " + strDatabase[0] + " p " +
                                     "ON e.biosum_cond_id=p.biosum_cond_id AND " +
                                     "e.rxpackage=p.rxpackage " +
                                     "SET e.pre_variable1_name = '" + oItem.strFVSVariableName + "'," +
                                     "e.post_variable1_name = '" + oItem.strFVSVariableName + "'," +
                                     "e.pre_variable1_value = IIF(p." + strDatabase[1] + " IS NOT NULL,p." + strDatabase[1] + ",0)," +
                                     "e.post_variable1_value = IIF(p." + strDatabase[1] + " IS NOT NULL,p." + strDatabase[1] + ",0)," +
                                     "e.variable1_change = 0";
                        }
                        // We specify a calculation for the total volume
                        else if (oItem.strFVSVariableName.Equals("total_volume_1"))
                        {
                            m_strSQL = "UPDATE tiebreaker e " +
                                 "INNER JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName + " p " +
                                 "ON e.biosum_cond_id=p.biosum_cond_id AND " +
                                 "e.rxpackage=p.rxpackage " +
                                 "SET e.pre_variable1_name = '" + oItem.strFVSVariableName + "'," +
                                 "e.post_variable1_name = '" + oItem.strFVSVariableName + "'," +
                                 "e.pre_variable1_value = IIF(p.chip_vol_cf_utilized + p.merch_vol_cf IS NOT NULL,p.chip_vol_cf_utilized + p.merch_vol_cf,0)," +
                                 "e.post_variable1_value = IIF(p.chip_vol_cf_utilized + p.merch_vol_cf IS NOT NULL,p.chip_vol_cf_utilized + p.merch_vol_cf,0)," +
                                 "e.variable1_change = 0";
                        }
                        else if (oItem.strFVSVariableName.Equals("treatment_haul_costs_1"))
                        {
                            m_strSQL = "UPDATE tiebreaker e " +
                            "INNER JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName + " p " +
                            "ON e.biosum_cond_id=p.biosum_cond_id AND " +
                            "e.rxpackage=p.rxpackage " +
                            "SET e.pre_variable1_name = '" + oItem.strFVSVariableName + "'," +
                            "e.post_variable1_name = '" + oItem.strFVSVariableName + "'," +
                            "e.pre_variable1_value = HARVEST_ONSITE_COST_DPA + MERCH_HAUL_COST_DPA + IIF (MERCH_CHIP_NR_DPA < MAX_NR_DPA,0, CHIP_HAUL_COST_DPA_utilized)," +
                            "e.post_variable1_value = HARVEST_ONSITE_COST_DPA + MERCH_HAUL_COST_DPA + IIF (MERCH_CHIP_NR_DPA < MAX_NR_DPA,0, CHIP_HAUL_COST_DPA_utilized)," +
                            "e.variable1_change = 0";
                        }
                    }
                    // This is a custom-weighted economic variable
                    // This is the same SQL used for built-in economic variables where the table/field are stored in database (see above)
                    else
                    {
                        string[] strDatabase = frmMain.g_oUtils.ConvertListToArray(oWeightItem.strVariableSource, ".");
                        m_strSQL = "UPDATE tiebreaker e " +
                                 "INNER JOIN " + strDatabase[0] + " p " +
                                 "ON e.biosum_cond_id=p.biosum_cond_id AND " +
                                 "e.rxpackage=p.rxpackage " +
                                 "SET e.pre_variable1_name = '" + oItem.strFVSVariableName + "'," +
                                 "e.post_variable1_name = '" + oItem.strFVSVariableName + "'," +
                                 "e.pre_variable1_value = IIF(p." + strDatabase[1] + " IS NOT NULL,p." + strDatabase[1] + ",0)," +
                                 "e.post_variable1_value = IIF(p." + strDatabase[1] + " IS NOT NULL,p." + strDatabase[1] + ",0)," +
                                 "e.variable1_change = 0";
                    }

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
                }
            }


            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			oItem = oTieBreakerCollection.Item(2);  // LAST TIEBREAK RANK
			if (oItem.bSelected)
			{
                m_strSQL = "UPDATE tiebreaker a INNER JOIN scenario_last_tiebreak_rank b ON a.rxpackage=b.rxpackage SET a.last_tiebreak_rank=b.last_tiebreak_rank " + 
					       "WHERE trim(ucase(b.scenario_id))= '" + ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			}

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
			


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

			string[]  strUpdateSQL = new string[25];

           
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
			ado_data_access p_ado = new ado_data_access();
			string strScenarioMDB="";
			string strScenarioConn="";
			p_ado.getScenarioConnStringAndMDBFile(ref strScenarioMDB,ref strScenarioConn, frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text);
			p_ado.OpenConnection(strScenarioConn);
			if (p_ado.m_intError != 0)
			{
				this.m_intError = p_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				p_ado = null;
				return;
			}

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			/********************************************
			 **delete all records in the table
			 ********************************************/
            this.m_strSQL = "delete from " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            

			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!!!Error Executing SQL!!!\r\n");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert Records\r\n");



            if (m_ado.TableExist(this.m_TempMDBFileConn, this.m_strEconByRxWorkTableName))
                m_ado.SqlNonQuery(this.m_TempMDBFileConn, "DROP TABLE " + this.m_strEconByRxWorkTableName);

            m_strSQL =
                "SELECT validcombos.biosum_cond_id,validcombos.rxpackage,validcombos.rx," +
                       "validcombos.rxcycle," +
                       this.m_strTreeVolValSumTable.Trim() + ".merch_vol_cf," +
                       this.m_strTreeVolValSumTable.Trim() + ".chip_vol_cf," +
                       this.m_strTreeVolValSumTable.Trim() + ".chip_wt_gt," +
                       this.m_strTreeVolValSumTable.Trim() + ".chip_val_dpa," +
                       this.m_strTreeVolValSumTable.Trim() + ".merch_wt_gt," +
                       this.m_strTreeVolValSumTable.Trim() + ".merch_val_dpa," +
                       this.m_strHvstCostsTable.Trim() + ".complete_cpa AS harvest_onsite_cost_dpa," +
                      "IIF(" + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + ".merch_haul_cost_dpgt IS NOT NULL," +
                      "IIF(" + this.m_strTreeVolValSumTable.Trim() + ".rxcycle='2'," + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + ".merch_haul_cost_dpgt * " + this.m_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle2 + "," +
                      "IIF(" + this.m_strTreeVolValSumTable.Trim() + ".rxcycle='3'," + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + ".merch_haul_cost_dpgt * " + this.m_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle3 + "," +
                      "IIF(" + this.m_strTreeVolValSumTable.Trim() + ".rxcycle='4'," + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + ".merch_haul_cost_dpgt * " + this.m_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle4 + "," +
                       Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + ".merch_haul_cost_dpgt))),0) AS escalator_merch_haul_cpa_pt," +
                      "escalator_merch_haul_cpa_pt * " + this.m_strTreeVolValSumTable.Trim() + ".merch_wt_gt AS merch_haul_cost_dpa," +
                      "IIF(" + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + ".chip_haul_cost_dpgt IS NOT NULL," +
                      "IIF(" + this.m_strTreeVolValSumTable.Trim() + ".rxcycle='2'," + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + ".chip_haul_cost_dpgt * " + this.m_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle2 + "," +
                      "IIF(" + this.m_strTreeVolValSumTable.Trim() + ".rxcycle='3'," + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + ".chip_haul_cost_dpgt * " + this.m_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle3 + "," +
                      "IIF(" + this.m_strTreeVolValSumTable.Trim() + ".rxcycle='4'," + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + ".chip_haul_cost_dpgt * " + this.m_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle4 + "," +
                       Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + ".chip_haul_cost_dpgt))),0) AS escalator_chip_haul_cpa_pt," +
                      "escalator_chip_haul_cpa_pt * " + this.m_strTreeVolValSumTable.Trim() + ".chip_wt_gt AS chip_haul_cost_dpa," +
                      "MERCH_VAL_DPA + CHIP_VAL_DPA - HARVEST_ONSITE_COST_DPA - (CHIP_HAUL_COST_DPA + MERCH_HAUL_COST_DPA) AS merch_chip_nr_dpa," +
                      "merch_val_dpa - harvest_onsite_cost_dpa - merch_haul_cost_dpa AS merch_nr_dpa," +
                      "IIF(" + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + ".chip_haul_psite IS NULL, 'N','Y') AS usebiomass_yn," +
                      "CDbl(0) AS max_nr_dpa, " +
                      this.m_strCondTable.Trim() + ".acres, " +
                      this.m_strCondTable.Trim() + ".owngrpcd, " +
                      "IIF(usebiomass_yn = 'Y',merch_haul_cost_dpa + chip_haul_cost_dpa, merch_haul_cost_dpa) AS haul_costs_dpa " +
                      "INTO " + this.m_strEconByRxWorkTableName +
                      " FROM ((((validcombos " +
                      "INNER JOIN " + this.m_strCondTable + " " +
                      "ON validcombos.biosum_cond_id = " + this.m_strCondTable.Trim() + ".biosum_cond_id) " +
                      "INNER JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " " +
                      "ON " +
                        this.m_strCondTable.Trim() + ".biosum_cond_id=" + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName
                        + ".biosum_cond_id) " +
                      "INNER JOIN " + this.m_strTreeVolValSumTable.Trim() + " " +
                      "ON " +
                      "(validcombos.biosum_cond_id=" + this.m_strTreeVolValSumTable.Trim() + ".biosum_cond_id) AND " +
                      "(validcombos.rxpackage=" + this.m_strTreeVolValSumTable.Trim() + ".rxpackage) AND " +
                      "(validcombos.rx=" + this.m_strTreeVolValSumTable.Trim() + ".rx) AND " +
                      "(validcombos.rxcycle=" + this.m_strTreeVolValSumTable.Trim() + ".rxcycle)) " +
                      "INNER JOIN " +
                        this.m_strHvstCostsTable.Trim() + " " +
                      "ON " +
                      "(validcombos.biosum_cond_id=" + this.m_strHvstCostsTable.Trim() + ".biosum_cond_id) AND " +
                      "(validcombos.rxpackage=" + this.m_strHvstCostsTable.Trim() + ".rxpackage) AND " +
                      "(validcombos.rx=" + this.m_strHvstCostsTable.Trim() + ".rx) AND " +
                      "(validcombos.rxcycle=" + this.m_strHvstCostsTable.Trim() + ".rxcycle)) " +
                      "WHERE " + this.m_strTreeVolValSumTable.Trim() + ".place_holder = 'N'";


            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");

                
                return;
            }


            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdates to " + this.m_strEconByRxWorkTableName + "\r\n");

            // Get the highest value for chips out of the Processor item. They should all be the same, but just in case ...            
            string strChipValue = "-1";
            for (int i = 0; i < this.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Count - 1; i++)
            {
                ProcessorScenarioItem.TreeSpeciesAndDbhDollarValuesItem oItem =
                  this.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(i);
                if (Convert.ToDouble(strChipValue) < Convert.ToDouble(oItem.ChipsDollarPerCubicFootValue.Trim()))
                {
                    strChipValue = oItem.ChipsDollarPerCubicFootValue.Trim();
                }
            }
            // Only use Biomass if chip revenue is higher than the cost of hauling them
            this.m_strSQL = "UPDATE " + this.m_strEconByRxWorkTableName +
                            " INNER JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " " +
                            "ON " +
                            this.m_strEconByRxWorkTableName + ".biosum_cond_id=" + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName
                            + ".biosum_cond_id " +
                            " SET usebiomass_yn = IIF(" + strChipValue + " < CHIP_HAUL_COST_DPGT, 'N', 'Y')" +
                            " WHERE usebiomass_yn = 'Y' AND RXCYCLE = '1'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");


                return;
            }

            this.m_strSQL = "UPDATE " + this.m_strEconByRxWorkTableName +
                            " INNER JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " " +
                            "ON " +
                            this.m_strEconByRxWorkTableName + ".biosum_cond_id=" + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName
                            + ".biosum_cond_id " +
                            " SET usebiomass_yn = IIF(" + strChipValue + " < CHIP_HAUL_COST_DPGT * " + this.m_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle2 + ", 'N', 'Y')" +
                            " WHERE usebiomass_yn = 'Y' AND RXCYCLE = '2'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");


                return;
            }

            this.m_strSQL = "UPDATE " + this.m_strEconByRxWorkTableName +
                " INNER JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " " +
                "ON " +
                this.m_strEconByRxWorkTableName + ".biosum_cond_id=" + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName
                + ".biosum_cond_id " +
                " SET usebiomass_yn = IIF(" + strChipValue + " < CHIP_HAUL_COST_DPGT * " + this.m_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle3 + ", 'N', 'Y')" +
                " WHERE usebiomass_yn = 'Y' AND RXCYCLE = '3'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");


                return;
            }

            this.m_strSQL = "UPDATE " + this.m_strEconByRxWorkTableName +
                " INNER JOIN " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " " +
                "ON " +
                this.m_strEconByRxWorkTableName + ".biosum_cond_id=" + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName
                + ".biosum_cond_id " +
                " SET usebiomass_yn = IIF(" + strChipValue + " < CHIP_HAUL_COST_DPGT * " + this.m_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle4 + ", 'N', 'Y')" +
                " WHERE usebiomass_yn = 'Y' AND RXCYCLE = '4'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");


                return;
            }

            // Don't use Biomass if no chip weight
            this.m_strSQL = "UPDATE " + this.m_strEconByRxWorkTableName +
                " SET usebiomass_yn = IIF(CHIP_WT_GT = 0 , 'N', 'Y')" +
                " WHERE usebiomass_yn = 'Y'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");


                return;
            }

            // Update fields based on USEBIOMASS_YN
            this.m_strSQL = "UPDATE " + this.m_strEconByRxWorkTableName +
                            " SET HAUL_COSTS_DPA = IIF(USEBIOMASS_YN = 'N', MERCH_HAUL_COST_DPA, MERCH_HAUL_COST_DPA + CHIP_HAUL_COST_DPA ), " +
                            " MAX_NR_DPA = IIF(USEBIOMASS_YN = 'Y', MERCH_CHIP_NR_DPA, MERCH_NR_DPA)";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");


                return;
            }

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            // Create temp table with inactive stands from harvest_costs table
            // This should only contain the inactive stands missing from the econ_by_rx_cycle worktable, but we will
            // do an outer join when appending them to be sure
            string strInactiveStandsWorkTable = "InactiveStandsWorkTable";
            this.m_strSQL = "SELECT validcombos.biosum_cond_id,validcombos.rxpackage,validcombos.rx,validcombos.rxcycle," +
                            " 0 AS merch_vol_cf, 0 as chip_vol_cf, 0 as chip_wt_gt," +
                            " 0 as chip_val_dpa, 0 as merch_wt_gt, 0 AS merch_val_dpa," +
                            " harvest_costs.complete_cpa AS harvest_onsite_cost_dpa, 0 AS escalator_merch_haul_cpa_pt," +
                            " 0 AS merch_haul_cost_dpa, 0 AS escalator_chip_haul_cpa_pt, 0 AS chip_haul_cost_dpa," +
                            " 0 AS merch_chip_nr_dpa, 0 AS merch_nr_dpa, 0 AS max_nr_dpa," +
                            " cond.acres, cond.owngrpcd, 0 AS haul_costs_dpa" +
                            " INTO " + strInactiveStandsWorkTable + " FROM" +
                            " ((validcombos INNER JOIN cond ON validcombos.biosum_cond_id = cond.biosum_cond_id)" +
                            " INNER JOIN harvest_costs ON (validcombos.biosum_cond_id=harvest_costs.biosum_cond_id) AND" +
                            " (validcombos.rxpackage=harvest_costs.rxpackage) AND" +
                            " (validcombos.rx=harvest_costs.rx) AND (validcombos.rxcycle=harvest_costs.rxcycle))" +
                            " WHERE harvest_costs.harvest_cpa = 0 and harvest_costs.complete_cpa <> 0";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
            }

            this.m_strSQL = "INSERT INTO " + this.m_strEconByRxWorkTableName +
                            " SELECT " + strInactiveStandsWorkTable + ".biosum_cond_id," + strInactiveStandsWorkTable + ".rxPackage, " +
                            strInactiveStandsWorkTable + ".rx, " + strInactiveStandsWorkTable + ".rxcycle," +
                            strInactiveStandsWorkTable + ".merch_vol_cf, " + strInactiveStandsWorkTable + ".chip_vol_cf," +
                            strInactiveStandsWorkTable + ".chip_wt_gt, " + strInactiveStandsWorkTable + ".chip_val_dpa," +
                            strInactiveStandsWorkTable + ".merch_wt_gt, " + strInactiveStandsWorkTable + ".merch_val_dpa," +
                            strInactiveStandsWorkTable + ".harvest_onsite_cost_dpa, " + strInactiveStandsWorkTable + ".escalator_merch_haul_cpa_pt," +
                            strInactiveStandsWorkTable + ".merch_haul_cost_dpa, " + strInactiveStandsWorkTable + ".escalator_chip_haul_cpa_pt," +
                            strInactiveStandsWorkTable + ".chip_haul_cost_dpa, " + strInactiveStandsWorkTable + ".merch_chip_nr_dpa," +
                            strInactiveStandsWorkTable + ".merch_nr_dpa, 'N' as usebiomass_yn, " + 
                            strInactiveStandsWorkTable + ".max_nr_dpa, " + strInactiveStandsWorkTable + ".acres," +
                            strInactiveStandsWorkTable + ".owngrpcd, " + strInactiveStandsWorkTable + ".haul_costs_dpa" +
                            " FROM " + strInactiveStandsWorkTable +
                            " LEFT outer join " + this.m_strEconByRxWorkTableName + " on " + 
                            this.m_strEconByRxWorkTableName + ".biosum_cond_id = " + strInactiveStandsWorkTable + ".biosum_cond_id" +
                            " and " + this.m_strEconByRxWorkTableName + ".rxpackage = " + strInactiveStandsWorkTable + ".rxpackage" +
                            " and " + this.m_strEconByRxWorkTableName + ".rx = " + strInactiveStandsWorkTable + ".rx" +
                            " and " + this.m_strEconByRxWorkTableName + ".rxcycle = " + strInactiveStandsWorkTable + ".rxcycle" + 
                            " where " + this.m_strEconByRxWorkTableName + ".biosum_cond_id is null";
            
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
            }
            
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            
            this.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName +
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
                            "FROM " + this.m_strEconByRxWorkTableName;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

			if (this.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
			}
     

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
                oOptimizerScenarioTools.LoadWeightedVariables(this.m_ado, oWeightedVariableCollection);
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

                string strEconConn = m_ado.getMDBConnString(m_strSystemResultsDbPathAndFile, "", "");


                try
                { 
                    using (var econConn = new OleDbConnection(strEconConn))
                    {
                    econConn.Open();

                    //Add columns to post_economic_weighted table to receive the data
                    foreach (string strFieldName in lstFieldNames)
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Adding columns for: " + strFieldName + "\r\n");
                        m_ado.AddColumn(econConn, Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName,
                            "c1_" + strFieldName, "DOUBLE", "");
                        m_ado.AddColumn(econConn, Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName,
                            "c2_" + strFieldName, "DOUBLE", "");
                        m_ado.AddColumn(econConn, Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName,
                            "c3_" + strFieldName, "DOUBLE", "");
                        m_ado.AddColumn(econConn, Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName,
                            "c4_" + strFieldName, "DOUBLE", "");
                        m_ado.AddColumn(econConn, Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName,
                            strFieldName, "DOUBLE", "");
                    }
                    
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                    System.Collections.Generic.IDictionary<string, ProductYields> dictProductYields =
                    new System.Collections.Generic.Dictionary<string, ProductYields>();
                    string strSql = "select * from " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strSql + "\r\n");
                    this.m_ado.SqlQueryReader(econConn, strSql);
                    ProductYields oProductYields = null;
                    while (this.m_ado.m_OleDbDataReader.Read())
                    {
                        string strCondId = this.m_ado.m_OleDbDataReader["biosum_cond_id"].ToString().Trim();
                        string strRxPackage = this.m_ado.m_OleDbDataReader["rxpackage"].ToString().Trim();
                        string strKey = strCondId + "_" + strRxPackage;
                        if (dictProductYields.ContainsKey(strKey))
                        {
                            oProductYields = dictProductYields[strKey];
                        }
                        else
                        {
                            oProductYields = new ProductYields(strCondId, strRxPackage);
                        }
                        string strRxCycle = this.m_ado.m_OleDbDataReader["rxcycle"].ToString().Trim();
                        double dblChipYieldCf = Convert.ToDouble(this.m_ado.m_OleDbDataReader["chip_vol_cf"]);
                        double dblMerchYieldCf = Convert.ToDouble(this.m_ado.m_OleDbDataReader["merch_vol_cf"]);
                        double dblHarvestOnsiteCpa = Convert.ToDouble(this.m_ado.m_OleDbDataReader["harvest_onsite_cost_dpa"]);
                        double dblMaxNrDpa = Convert.ToDouble(this.m_ado.m_OleDbDataReader["max_nr_dpa"]);
                        double dblHaulMerchCpa = Convert.ToDouble(this.m_ado.m_OleDbDataReader["merch_haul_cost_dpa"]);
                        double dblMerchChipNrDpa = Convert.ToDouble(this.m_ado.m_OleDbDataReader["merch_chip_nr_dpa"]);
                        double dblHaulChipCpa = Convert.ToDouble(this.m_ado.m_OleDbDataReader["chip_haul_cost_dpa"]);

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
                            strSql = strSqlPrefix + oSavedProductYields.CondId() + "', '" +
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

                            this.m_ado.SqlNonQuery(econConn, strSql);
                        }
                    }
                    }
                }
                catch (Exception err)
                {
                    this.m_ado.m_intError = -1;
                    this.m_ado.m_strError = "Error calculating weighted economic variables: " + err.Message;
                    MessageBox.Show("!! " + this.m_ado.m_strError + " !!", "FIA Biosum");
                }

                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                if (this.m_ado.m_intError != 0)
                {
                    if (frmMain.g_bDebug) frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
                    this.m_intError = this.m_ado.m_intError;
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


            /********************************************
             **delete all records in the table
             ********************************************/
            this.m_strSQL = "delete from " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName;
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
            }
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            // We manipulate econ_by_rx_worktable to zero out some chip fields if use_biomass_yn='N'
            // Then we sum from the worktable to get the correct numbers in the summary fields
            m_strSQL = "UPDATE " + this.m_strEconByRxWorkTableName +
                       " SET chip_vol_cf = IIF(usebiomass_yn = 'N', 0, chip_vol_cf)," +
                       " chip_wt_gt = IIF(usebiomass_yn = 'N', 0, chip_wt_gt)," +
                       " chip_val_dpa = IIF(usebiomass_yn = 'N', 0, chip_val_dpa)," +
                       " chip_haul_cost_dpa = IIF(usebiomass_yn = 'N', 0, chip_haul_cost_dpa)";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate " + this.m_strEconByRxWorkTableName + " based on use_biomass_yn values\r\n");

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");

                return;
            }

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert Records\r\n");

            m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName;
            m_strSQL =  m_strSQL + " (biosum_cond_id,rxpackage," +
                                    "chip_vol_cf_utilized,merch_vol_cf," +
                                    "chip_wt_gt_utilized,merch_wt_gt," +
                                    "chip_val_dpa_utilized,merch_val_dpa," +
                                    "harvest_onsite_cost_dpa,chip_haul_cost_dpa_utilized,merch_haul_cost_dpa," +
                                    "merch_chip_nr_dpa," +
                                    "merch_nr_dpa," +
                                    "max_nr_dpa,haul_costs_dpa," +
                                    "treated_acres,acres, owngrpcd)";
            m_strSQL = m_strSQL + " " +
                            "SELECT a.biosum_cond_id,a.rxpackage," +
                                   "a.sum_chip_yield_cf AS chip_vol_cf_utilized," +
                                   "a.sum_merch_yield_cf AS merch_vol_cf," +
                                   "a.sum_chip_yield_gt AS chip_wt_gt_utilized," +
                                   "a.sum_merch_yield_gt AS merch_wt_gt," +
                                   "a.sum_chip_val_dpa AS chip_val_dpa_utilized," +
                                   "a.sum_merch_val_dpa AS merch_val_dpa," +
                                   "a.sum_harvest_onsite_cpa AS harvest_onsite_cost_dpa," +
                                   "a.sum_haul_chip_cpa AS chip_haul_cost_dpa_utilized," +
                                   "a.sum_haul_merch_cpa AS merch_haul_cost_dpa," +
                                   "a.sum_merch_chip_nr_dpa AS merch_chip_nr_dpa," +
                                   "a.sum_merch_nr_dpa AS merch_nr_dpa, " +
                                   "a.sum_max_nr_dpa AS max_nr_dpa, " +
                                   "a.sum_haul_costs_dpa AS haul_costs_dpa, " +
                                   "a.sum_treated_acres AS treated_acres, " +
                                   "a.acres, a.owngrpcd " +
                           "FROM (" +
                           "SELECT biosum_cond_id,rxpackage," +
                                "SUM(IIF(chip_vol_cf IS NULL,0,chip_vol_cf)) AS sum_chip_yield_cf," +
                                "SUM(IIF(merch_vol_cf IS NULL,0,merch_vol_cf)) AS sum_merch_yield_cf," +
                                "SUM(IIF(chip_wt_gt IS NULL,0,chip_wt_gt)) AS sum_chip_yield_gt," +
                                "SUM(IIF(merch_wt_gt IS NULL,0,merch_wt_gt)) AS sum_merch_yield_gt," +
                                "SUM(IIF(chip_val_dpa IS NULL,0,chip_val_dpa)) AS sum_chip_val_dpa," +
                                "SUM(IIF(merch_val_dpa IS NULL,0,merch_val_dpa)) AS sum_merch_val_dpa," +
                                "SUM(IIF(harvest_onsite_cost_dpa IS NULL,0,harvest_onsite_cost_dpa)) AS sum_harvest_onsite_cpa," +
                                "SUM(IIF(chip_haul_cost_dpa IS NULL,0,chip_haul_cost_dpa)) AS sum_haul_chip_cpa," +
                                "SUM(IIF(merch_haul_cost_dpa IS NULL,0,merch_haul_cost_dpa)) AS sum_haul_merch_cpa," +
                                "SUM(IIF(merch_chip_nr_dpa IS NULL,0,merch_chip_nr_dpa)) AS sum_merch_chip_nr_dpa," +
                                "SUM(IIF(merch_nr_dpa IS NULL,0,merch_nr_dpa)) AS sum_merch_nr_dpa," +
                                "SUM(IIF(max_nr_dpa IS NULL,0,max_nr_dpa)) AS sum_max_nr_dpa, " +
                                "SUM(IIF(haul_costs_dpa IS NULL,0,haul_costs_dpa)) AS sum_haul_costs_dpa, " +
                                "SUM(IIF(acres IS NULL,0,acres)) AS sum_treated_acres, " +
                                "acres,owngrpcd " +
                           "FROM " + this.m_strEconByRxWorkTableName +
                           " GROUP BY biosum_cond_id,rxpackage,acres,owngrpcd) a";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            if (this.m_ado.m_intError != 0)
            {
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Executing SQL!!!\r\n");
                this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");

                return;
            }

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            // Create worktable for HVST_TYPE_BY_CYCLE field
            string strTreeSumWorktableName = "ECON_SUM_WORKTABLE";
            if (this.m_ado.TableExist(this.m_TempMDBFileConn, strTreeSumWorktableName))
            {
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, "DELETE FROM " + strTreeSumWorktableName);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nDelete all records from econ_sum_worktable\r\n");

            }
            else
            {
                this.m_strSQL = "CREATE TABLE " + strTreeSumWorktableName + " (" +
                                "biosum_cond_id CHAR(25), " +
                                "rxpackage CHAR(3), " +
                                "CYCLE_1 CHAR(1), " +
                                "CYCLE_2 CHAR(1), " +
                                "CYCLE_3 CHAR(1), " +
                                "CYCLE_4 CHAR(1))";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nCreate econ_sum_worktable\r\n");

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            }

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }

            //Add rows to worktable
            this.m_strSQL = "INSERT INTO " + strTreeSumWorktableName +
                            " SELECT BIOSUM_COND_ID, RXPACKAGE, '0' as CYCLE_1, '0' AS CYCLE_2, '0' AS CYCLE_3, '0' AS CYCLE_4" +
                            " FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nAdd condition/rxpackages to econ_sum_worktable\r\n");

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }


            // Update worktable from econ_by_rx_cycle
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate cycle values on econ_sum_worktable\r\n");
            string[] arrFieldToUpdate = { "CYCLE_1", "CYCLE_2", "CYCLE_3", "CYCLE_4" };
            string[] arrRxCycle = { "1", "2", "3", "4" };
            for (int arrIdx = 0; arrIdx < 4; arrIdx++)
            {
                this.m_strSQL = "UPDATE " + strTreeSumWorktableName +
                                " INNER JOIN econ_by_rx_cycle ON (" + strTreeSumWorktableName + ".BIOSUM_COND_ID=econ_by_rx_cycle.BIOSUM_COND_ID)" +
                                " AND (" + strTreeSumWorktableName + ".RXPACKAGE= econ_by_rx_cycle.RXPACKAGE)" +
                                " SET " + strTreeSumWorktableName + "." + arrFieldToUpdate[arrIdx] + " = IIF (chip_vol_cf > 0 AND merch_vol_cf > 0, '3'," +
                                " IIF (chip_vol_cf = 0 AND merch_vol_cf = 0, '0'," +
                                " IIF (chip_vol_cf > 0, '2','1'))) " +
                                " WHERE RXCYCLE = '" + arrRxCycle[arrIdx] + "'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            }
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }

            // Populate hvst_type_by_cycle from worktable
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate hvst_type_by_cycle from econ_sum_worktable\r\n");
            this.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName +
                            " INNER JOIN " + strTreeSumWorktableName +
                            " ON (" + strTreeSumWorktableName + ".BIOSUM_COND_ID = " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName + ".BIOSUM_COND_ID)" +
                            " AND (" + strTreeSumWorktableName + ".RXPACKAGE = " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName + ".RXPACKAGE)" +
                            " SET hvst_type_by_cycle = CYCLE_1 + CYCLE_2 + CYCLE_3 + CYCLE_4 ";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                this.m_intError = this.m_ado.m_intError;
                return;
            }

            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
            

            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");

            }


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

           
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn = p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
                 * set the application version in the database
                 * *******************************************/
                string strConn = p_ado.getMDBConnString(this.m_strSystemResultsDbPathAndFile, "", "");
                using (OleDbConnection versionConn = new OleDbConnection(strConn))
                {
                    versionConn.Open();
                    this.m_strSQL = "INSERT INTO VERSION (APPLICATION_VERSION)" +
                                    " VALUES ('" + frmMain.g_strAppVer + "') ";
                    p_ado.SqlNonQuery(versionConn, this.m_strSQL);
                }
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

                              
                /*********************************************
				 **get the harvest_costs structure
				 *********************************************/
				this.m_strSQL = "SELECT biosum_cond_id,rxpackage,rx,rxcycle, complete_cpa FROM " + this.m_strHvstCostsTable.Trim() + ";";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of harvest_costs_sum
				 *****************************************************************/
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create harvest_costs_sum Table Schema From harvest_costs table--\r\n");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"harvest_costs_sum",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
			}
			else
			{
				this.m_intError = p_ado.m_intError;
			}
			p_ado=null;

		}

		private void CreateTableStructureForScenarioProcessingSites()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForScenarioProcessingSites\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
				 **get the haul costs table structure
				 *********************************************/
				this.m_strSQL = "SELECT psite_id,trancd,biocd FROM scenario_psites;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

                 if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create scenario psites work table from  the scenario_psites table--\r\n");
				dao_data_access p_dao = new dao_data_access();
                this.m_strPSiteWorkTable = "scenario_psites_work_table";
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,this.m_strPSiteWorkTable,p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}	
				p_dt.Clear();
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
			}
			else
			{
				this.m_intError = p_ado.m_intError;
			}
			p_ado=null;

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
			ado_data_access oAdo = new ado_data_access();
			this.m_strConn= oAdo.getMDBConnString(this.m_strTempMDBFile,"admin","");
			oAdo.OpenConnection(this.m_strConn);
			if (oAdo.m_intError==0)
			{

                /*****************************************************************
                 **create the table structures in the temp mdb file
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

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--Create psite_accessible_work_table\r\n");

                if (oAdo.TableExist(oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName))
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName);

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "--Create haul costs work tables from the haul_costs table SQL\r\n");

                if (oAdo.m_intError == 0)
                    frmMain.g_oTables.m_oOptimizerScenarioResults.CreatePSitesWorktable(
                        oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName);
                
                if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"all_road_chip_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE all_road_chip_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"all_road_chip_haul_costs_work_table");


				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"all_road_merch_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE all_road_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"all_road_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"cheapest_road_merch_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE cheapest_road_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"cheapest_road_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"cheapest_road_chip_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE cheapest_road_chip_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"cheapest_road_chip_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"cheapest_rail_merch_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE cheapest_rail_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"cheapest_rail_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"cheapest_rail_chip_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE cheapest_rail_chip_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"cheapest_rail_chip_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"cheapest_merch_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE cheapest_merch_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostWorkTable(
						oAdo,oAdo.m_OleDbConnection,"cheapest_merch_haul_costs_work_table");

				

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"cheapest_chip_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE cheapest_chip_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostWorkTable(
						oAdo,oAdo.m_OleDbConnection,"cheapest_chip_haul_costs_work_table");

				
				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"chip_plot_to_rh_to_collector_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE chip_plot_to_rh_to_collector_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"chip_plot_to_rh_to_collector_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"merch_plot_to_rh_to_collector_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE merch_plot_to_rh_to_collector_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"merch_plot_to_rh_to_collector_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"combine_chip_rail_road_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE combine_chip_rail_road_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"combine_chip_rail_road_haul_costs_work_table");

				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"combine_merch_rail_road_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE combine_merch_rail_road_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostTable(
						oAdo,oAdo.m_OleDbConnection,"combine_merch_rail_road_haul_costs_work_table");


				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"merch_rh_to_collector_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE merch_rh_to_collector_haul_costs_work_table");

				/*****************************************************************
				 **create the railroad table structures in the temp mdb file
				 **and give them the name OF merch_rh_to_collector_haul_costs_work_table
				 **                          chip_rh_to_collector_haul_costs_work_table
				 *****************************************************************/
				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostRailroadTable(
						oAdo,oAdo.m_OleDbConnection,"merch_rh_to_collector_haul_costs_work_table");


				if (oAdo.m_intError==0)
					if (oAdo.TableExist(oAdo.m_OleDbConnection,"chip_rh_to_collector_haul_costs_work_table"))
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE chip_rh_to_collector_haul_costs_work_table");

				if (oAdo.m_intError==0)
					frmMain.g_oTables.m_oOptimizerScenarioResults.CreateHaulCostRailroadTable(
						oAdo,oAdo.m_OleDbConnection,"chip_rh_to_collector_haul_costs_work_table");

				
				if (oAdo.m_intError!=0)
				{

                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=oAdo.m_intError;
					oAdo.CloseConnection(oAdo.m_OleDbConnection);
					oAdo=null;
					return;
				}
				oAdo.CloseConnection(oAdo.m_OleDbConnection);
			}
			else
			{
				this.m_intError = oAdo.m_intError;
			}
			oAdo=null;

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
			ado_data_access oAdo = new ado_data_access();
			this.m_strConn= oAdo.getMDBConnString(this.m_strTempMDBFile,"admin","");
			oAdo.OpenConnection(this.m_strConn);


			if (oAdo.m_intError==0)
			{
				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of rx_intensity_work_table
				 *****************************************************************/
                 if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create rx_intensity_work_table Schema--\r\n");

				frmMain.g_oTables.m_oOptimizerScenarioResults.CreateIntensityWorkTable(oAdo,oAdo.m_OleDbConnection,"rx_intensity_duplicates_work_table");
				if (oAdo.m_intError==0) frmMain.g_oTables.m_oOptimizerScenarioResults.CreateIntensityWorkTable(oAdo,oAdo.m_OleDbConnection,"rx_intensity_duplicates_work_table2");
				if (oAdo.m_intError==0) frmMain.g_oTables.m_oOptimizerScenarioResults.CreateIntensityWorkTable(oAdo,oAdo.m_OleDbConnection,"rx_intensity_unique_work_table");

				if (oAdo.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=oAdo.m_intError;
					oAdo.CloseConnection(oAdo.m_OleDbConnection);
					return;
				}
				oAdo.CloseConnection(oAdo.m_OleDbConnection);
			}
			else
			{
				this.m_intError = oAdo.m_intError;
			}
			oAdo=null;

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
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
				 **get fields from the plot table
				 *********************************************/
				this.m_strSQL = "SELECT biosum_plot_id, num_cond, num_cond as num_cond_not_accessible FROM " + this.m_strPlotTable;

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of plot_cond_accessible_work_table
				 *****************************************************************/
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create plot_cond_accessible_work_table Schema--\r\n");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_cond_accessible_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_cond_accessible_work_table2",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
			}
			else
			{
				this.m_intError = p_ado.m_intError;
			}
			p_ado=null;

		}
		private void CreateTableStructureForUserDefinedConditionTable()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateTableStructureForUserDefinedConditionTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/********************************************************
				 **get the user defined PLOT filter sql
				 ********************************************************/
				this.m_strSQL = this.m_strUserDefinedCondSQL;
				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = this.m_ado.getTableSchema(this.m_TempMDBFileConn,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the optimizer_results.accdb file
				 **and give it the name of userdefinedcondfilter_work
				 *****************************************************************/
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create userdefinedcondfilter_work Table Schema From User Defined Condition Filter SQL--\r\n");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"userdefinedcondfilter_work",p_dt,true);
				p_dt.Dispose();
				this.m_ado.m_OleDbDataReader.Close();
				if (p_dao.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_dao=null;
					return;
				}

				/*********************************************
				 **get fields from the plot table
				 *********************************************/
				this.m_strSQL = "SELECT * FROM " + this.m_strCondTable;

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of plot_cond_accessible_work_table
				 *****************************************************************/
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--Create ruledefinitionscondfilter Schema--\r\n");
				p_dao.CreateMDBTableFromDataSetTable(this.m_strSystemResultsDbPathAndFile,"ruledefinitionscondfilter",p_dt,true);
				/***********************************************************************
   			     **create a table link to the ruledefinitionscondfilter table in the
			     **temporary MDB file located on the hard drive of the user
			     ***********************************************************************/
				p_dao.CreateTableLink(this.m_strTempMDBFile,"ruledefinitionscondfilter",m_strSystemResultsDbPathAndFile,"ruledefinitionscondfilter",true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n!!!Error Creating Table Schema!!!\r\n");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
			}
			else
			{
				this.m_intError = p_ado.m_intError;
			}
			p_ado=null;

		}

		/// <summary>
		/// create temporary work table for summing up wood volumes,values and costs
		/// by wood processing site
		/// </summary>
		private void CreateTableStructureForPSiteSumOld()
		{
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/**********************************************************
				 **create the max_nr_plots structure with sum_ field names
				 **********************************************************/
				this.m_strSQL = "SELECT s.psite_id, " +
					"s.biocd," + 
					"p.acres AS sum_acres," + 
					"p.merch_haul_cents AS sum_merch_haul_cents," + 
					"p.chip_haul_cents AS sum_chip_haul_cents," + 
					"p.merch_vol AS sum_merch_vol," + 
					"p.chip_yield AS sum_chip_yield," +
					"p.net_rev AS sum_net_rev," + 
					"p.merch_dollars_val AS sum_merch_dollars_val," + 
					"p.chip_dollars_val AS sum_chip_dollars_val," + 
					"p.harv_costs AS sum_harv_costs," + 
					"p.haul_costs AS sum_haul_costs," + 
					"p.ti_chg_acres AS sum_ti_chg_acres," + 
					"p.ci_chg_acres AS sum_ci_chg_acres " + 
					"FROM max_nr_plots p," + this.m_strPSiteWorkTable.Trim() + " s;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of rx_intensity_work_table
				 *****************************************************************/
				
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"psite_sum_work_table",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					//this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
			}
			else
			{
				this.m_intError = p_ado.m_intError;
			}
			p_ado=null;

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
			string strTable="";
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Text","Finding Best Treatments: Maximum Net Revenue");
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Visible",true);
			frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)ReferenceUserControlScenarioRun.lblMsg,"Refresh");
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"\r\n\r\nBest Rx Summary\r\n");
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--------------------\r\n");

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


		    string strScenarioId = this.ReferenceUserControlScenarioRun.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim();
			string strTieBreakerAggregate="MAX";
			
			

	
			//
			//CREATE WORK TABLES
			//
			//best_rx_summary_work_table
            strTable = "cycle1_best_rx_summary_work_table";
			if (m_ado.TableExist(this.m_TempMDBFileConn, strTable))
			{
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,
					"DROP TABLE " + strTable);
			}
			m_ado.SqlNonQuery(m_TempMDBFileConn,Tables.OptimizerScenarioResults.CreateBestRxSummaryCycle1TableSQL(strTable));
            //best_rx_summury_optimization_and_tiebreaker_work_table
			if (m_ado.TableExist(m_TempMDBFileConn,"cycle1_best_rx_summary_optimization_and_tiebreaker_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table");
			}
			m_ado.SqlNonQuery(m_TempMDBFileConn,Tables.OptimizerScenarioResults.CreateBestRxSummaryCycle1TieBreakerTableSQL("cycle1_best_rx_summary_optimization_and_tiebreaker_work_table"));
			//best_rx_summury_optimization_and_tiebreaker_work_table2
			if (m_ado.TableExist(m_TempMDBFileConn,"cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2");
			}
			m_ado.SqlNonQuery(m_TempMDBFileConn,Tables.OptimizerScenarioResults.CreateBestRxSummaryCycle1TieBreakerTableSQL("cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2"));
			//best_rx_summury_optimization_and_tiebreaker_work_table3
			if (m_ado.TableExist(m_TempMDBFileConn,"cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3");
			}
			m_ado.SqlNonQuery(m_TempMDBFileConn,Tables.OptimizerScenarioResults.CreateBestRxSummaryCycle1TieBreakerTableSQL("cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3"));

			/**********************************************
			 **insert unique biosum_cond_id's into the
			 **best_rx_summary table so we dont have
			 **to worry about whether the biosum_cond_id 
			 **record is in the table or not
			 **********************************************/
			this.m_strSQL = "INSERT INTO " + strTable  +  " " +
				"SELECT DISTINCT c.biosum_cond_id,c.acres,c.owngrpcd " + 
				"FROM " + this.m_strCondTable.Trim() + " c, " + 
				Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + " p, "  + 
               
				this.ReferenceOptimizerScenarioForm.OutputTablePrefix +
                Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix + " e " +
				"WHERE c.biosum_cond_id = p.biosum_cond_id  AND " + 
				"e.biosum_cond_id = c.biosum_cond_id AND " + 
				"e.affordable_YN = 'Y' AND e.rxcycle='1' AND " + 
				"(p.merch_haul_cost_id IS NOT NULL OR " + 
				"p.chip_haul_cost_id IS NOT NULL);";

             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--insert condition records that have MERCH or CHIP haul costs into best_rx_summary--\r\n");
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL:" + this.m_strSQL + "\r\n\r\n");
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

           

			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();


            string strWorkTable = "cycle1_effective_" + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;
            if (m_ado.TableExist(this.m_TempMDBFileConn, strWorkTable))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE " + strWorkTable);
			}

			m_ado.m_strSQL = "SELECT p.* " + 
				"INTO " + strWorkTable + 
				" FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName + " p," + 
                ReferenceOptimizerScenarioForm.OutputTablePrefix + Tables.OptimizerScenarioResults.DefaultScenarioResultsEffectiveTableSuffix + " e " + 
				"WHERE p.biosum_cond_id = e.biosum_cond_id AND " + 
				"p.rxpackage=e.rxpackage AND p.rx=e.rx AND p.rxcycle=e.rxcycle AND e.overall_effective_yn='Y'";

             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--write overall effective treatments to the " + strWorkTable + "--\r\n");
             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
			
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			

			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (m_ado.TableExist(this.m_TempMDBFileConn,"cycle1_effective_optimization_treatments"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE cycle1_effective_optimization_treatments");
			}

			m_ado.m_strSQL = "SELECT e.* " + 
				"INTO cycle1_effective_optimization_treatments " + 
				"FROM " + ReferenceOptimizerScenarioForm.OutputTablePrefix +
                Tables.OptimizerScenarioResults.DefaultScenarioResultsEffectiveTableSuffix + " e " + 
				"WHERE e.overall_effective_yn='Y'";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--write overall effective treatments to the cycle1_effective_optimization_treatments--\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "REVENUE")
			{
				Best_rx_summary(oTieBreakerCollection,strTieBreakerAggregate,false);

			}
			else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "MERCHANTABLE VOLUME")
			{
				Best_rx_summary(oTieBreakerCollection,strTieBreakerAggregate,false);
			}
            else if (this.m_oOptimizationVariable.strOptimizedVariable.Trim().ToUpper() == "ECONOMIC ATTRIBUTE")
            {
                Best_rx_summary(oTieBreakerCollection, strTieBreakerAggregate, false);
            }
			else
			{
				Best_rx_summary(oTieBreakerCollection,strTieBreakerAggregate,true);
			}

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            string strBestRxSummaryTableName = this.ReferenceOptimizerScenarioForm.OutputTablePrefix +
                Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryTableSuffix;
			
            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}


            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

			/****************************************************************************
			 **finished with minimum merchantable wood removal with positive net revenue
			 ****************************************************************************/


			if (this.m_intError == 0)
			{
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");
			}


		
		}

		private void Best_rx_summary(
			FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker.TieBreaker_Collection oTieBreakerCollection, 
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
			string strSql="";



            string strOptimizationTableName = ReferenceOptimizerScenarioForm.OutputTablePrefix +
                Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix;
            if (bFVSVariable==false)
			{
				//find the treatment for each plot that produces the MAX/MIN revenue value
				strSql = "SELECT a.biosum_cond_id,a.rxpackage,a.rx,a." + this.m_strOptimizationColumnNameSql + " AS optimization_value " + //LPOTTS,a.rx_intensity " + 
					"FROM " + strOptimizationTableName + " a,";


				strSql=strSql + "(SELECT " + this.m_strOptimizationAggregateSql + "(" + this.m_strOptimizationColumnNameSql + ") AS " + this.m_strOptimizationAggregateColumnName + ",biosum_cond_id " +
                    "FROM " + strOptimizationTableName + " where affordable_YN = 'Y'";

				strSql = strSql + " GROUP BY biosum_cond_id) b ";

				strSql = strSql + "WHERE a.biosum_cond_id=b.biosum_cond_id AND a." + this.m_strOptimizationColumnNameSql + " = b." + this.m_strOptimizationAggregateColumnName +
                                  " AND a.affordable_YN = 'Y'";
			}
			else
			{
				strSql = "SELECT a.biosum_cond_id,a.rxpackage,a.rx,a." + this.m_strOptimizationColumnNameSql + " AS optimization_value " + //LPOTTS,a.rx_intensity " + 
					"FROM " + strOptimizationTableName + " a,";


				strSql=strSql + "(SELECT " + this.m_strOptimizationAggregateSql + "(" + this.m_strOptimizationColumnNameSql + ") AS " + this.m_strOptimizationAggregateColumnName + ",biosum_cond_id " +
                    "FROM " + strOptimizationTableName + " where affordable_YN = 'Y'";

				strSql = strSql + " GROUP BY biosum_cond_id) b ";

                strSql = strSql + "WHERE a.biosum_cond_id=b.biosum_cond_id AND a." + this.m_strOptimizationColumnNameSql + " = b." + this.m_strOptimizationAggregateColumnName +
                    " AND a.affordable_YN = 'Y'";
			}

			strSql = "INSERT INTO cycle1_best_rx_summary_optimization_and_tiebreaker_work_table " + 
				strSql;
		
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--filter effective treatments to find " + this.m_strOptimizationAggregateSql + " " + this.m_oOptimizationVariable.strOptimizedVariable + "--\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL:" + strSql + "\r\n\r\n");
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,strSql);
			if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
			
			if (this.m_ado.m_intError != 0)
			{
                if (frmMain.g_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"!!!Error Executing SQL!!!\r\n");
				this.m_intError = this.m_ado.m_intError;
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
				return;
			}

				

			m_ado.m_strSQL = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a " + 
				"INNER JOIN cycle1_best_rx_summary_work_table b " + 
				"ON a.biosum_cond_id=b.biosum_cond_id " + 
				"SET a.acres = b.acres,a.owngrpcd=b.owngrpcd";
           if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"Execute SQL:" + m_ado.m_strSQL + "\r\n");
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

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
				strSql = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a " + 
					"INNER JOIN tiebreaker b " + 
					"ON a.biosum_cond_id=b.biosum_cond_id AND a.rxpackage=b.rxpackage AND a.rx=b.rx " +
                    "SET a.tiebreaker_value = b." + strTiebreakerValueField + ",a.last_tiebreak_rank=b.last_tiebreak_rank";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + strSql + "\r\n");
		        m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);						


				m_ado.m_strSQL ="INSERT INTO " + ReferenceOptimizerScenarioForm.OutputTablePrefix +
                    Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryBeforeTiebreaksTableSuffix +
                    " SELECT DISTINCT * FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);


				//find the treatment for each plot that produces the MAX/MIN tiebreaker value
				m_ado.m_strSQL ="SELECT a.biosum_cond_id,a.acres,a.owngrpcd,a.optimization_value,a.tiebreaker_value,a.rxpackage,a.rx,a.last_tiebreak_rank " + 
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a," +
					"(SELECT biosum_cond_id," + strTieBreakerAggregate + "(tiebreaker_value) AS tiebreaker " +
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table " + 
					"GROUP BY biosum_cond_id) c " + 
					"WHERE a.biosum_cond_id=c.biosum_cond_id AND a.tiebreaker_value=c.tiebreaker";

					
				m_ado.m_strSQL = "INSERT INTO cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 " + m_ado.m_strSQL;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--break any ties by finding the " + strTieBreakerAggregate + " tie breaker value--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
				
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
				
				if (this.m_ado.m_intError != 0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
					this.m_intError = this.m_ado.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
					return;
				}
					


				m_ado.m_strSQL = "SELECT a.biosum_cond_id,a.acres,a.owngrpcd,a.optimization_value," + 
					"a.tiebreaker_value,a.rxpackage,a.rx,a.last_tiebreak_rank " + 
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 a," +
					"(SELECT biosum_cond_id,MIN(last_tiebreak_rank) AS min_intensity " + 
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 " + 
					"GROUP BY biosum_cond_id) c " +
                    "WHERE a.biosum_cond_id=c.biosum_cond_id AND a.last_tiebreak_rank=c.min_intensity";

				m_ado.m_strSQL = "INSERT INTO cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3 " + m_ado.m_strSQL;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--break any additional ties by finding the least intense treatment--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
				
				if (this.m_ado.m_intError != 0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
					this.m_intError = this.m_ado.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
					return;
				}
					

				m_ado.m_strSQL = "UPDATE cycle1_best_rx_summary_work_table a " + 
					"INNER JOIN cycle1_best_rx_summary_optimization_and_tiebreaker_work_table3 b " + 
					"ON a.biosum_cond_id=b.biosum_cond_id " + 
					"SET a.optimization_value=b.optimization_value," + 
					"a.tiebreaker_value=b.tiebreaker_value," +
                    "a.rxpackage=b.rxpackage," + 
					"a.rx=b.rx," +
                    "a.last_tiebreak_rank=b.last_tiebreak_rank";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n");
				m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);
				m_ado.m_strSQL="INSERT INTO " + ReferenceOptimizerScenarioForm.OutputTablePrefix + 
                    Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryTableSuffix +
                    " SELECT * FROM cycle1_best_rx_summary_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--insert the work table records into the best_rx_summary table--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;
				
				if (this.m_ado.m_intError != 0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
					this.m_intError = this.m_ado.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
					return;
				}

			}
			// Last tie-break rank ONLY
            else if (oTieBreakerCollection.Item(2).bSelected)
			{
				//update the rx intensity fields for each plot
				strSql = "UPDATE cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a " + 
					"INNER JOIN tiebreaker b " +
                    "ON a.biosum_cond_id=b.biosum_cond_id AND a.rx=b.rx AND a.rxpackage=b.rxpackage " +
                    "SET a.last_tiebreak_rank=b.last_tiebreak_rank";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + strSql + "\r\n");
				m_ado.SqlNonQuery(m_TempMDBFileConn,strSql);


				m_ado.m_strSQL ="INSERT INTO " + 
                    ReferenceOptimizerScenarioForm.OutputTablePrefix +
                    Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryBeforeTiebreaksTableSuffix + 
                    " SELECT DISTINCT * FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);


					
				m_ado.m_strSQL = "SELECT a.biosum_cond_id,a.acres,a.owngrpcd,a.optimization_value," +
                    "a.tiebreaker_value,a.rxpackage,a.rx,a.last_tiebreak_rank " + 
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table a," +
                    "(SELECT biosum_cond_id,MIN(last_tiebreak_rank) AS min_intensity " + 
					"FROM cycle1_best_rx_summary_optimization_and_tiebreaker_work_table " + 
					"GROUP BY biosum_cond_id) c " +
                    "WHERE a.biosum_cond_id=c.biosum_cond_id AND a.last_tiebreak_rank=c.min_intensity";

				m_ado.m_strSQL = "INSERT INTO cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 " + m_ado.m_strSQL;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--break any additional ties by finding the least intense treatment--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

				if (this.m_ado.m_intError != 0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
					this.m_intError = this.m_ado.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
					return;
				}
					

				m_ado.m_strSQL = "UPDATE cycle1_best_rx_summary_work_table a " + 
					"INNER JOIN cycle1_best_rx_summary_optimization_and_tiebreaker_work_table2 b " + 
					"ON a.biosum_cond_id=b.biosum_cond_id " + 
					"SET a.optimization_value=b.optimization_value," + 
					"a.tiebreaker_value=b.tiebreaker_value," +
                    "a.rxpackage=b.rxpackage," +
					"a.rx=b.rx," +
                    "a.last_tiebreak_rank=b.last_tiebreak_rank";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n");
				m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);
                m_ado.m_strSQL="INSERT INTO " + ReferenceOptimizerScenarioForm.OutputTablePrefix +
                    Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryTableSuffix +
                    " SELECT * FROM cycle1_best_rx_summary_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile,"--insert the work table records into the best_rx_summary table--\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL:" + m_ado.m_strSQL + "\r\n\r\n");
				m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
                if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic) == true) return;

				if (this.m_ado.m_intError != 0)
				{
                    if (frmMain.g_bDebug)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "!!!Error Executing SQL!!!\r\n");
					this.m_intError = this.m_ado.m_intError;
                    FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                    FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
					return;
				}

			}


		}
	
		//private void DeleteScenarioResultRecords()
		//{

			
		//	ado_data_access oAdo = new ado_data_access();
		//	string strConn=oAdo.getMDBConnString(this.m_strSystemResultsDbPathAndFile,"admin","");
		//	oAdo.OpenConnection(strConn);

		//	/*************************************************
		//	 **delete all records in the by plot tables
		//	 *************************************************/
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_nr_plots"))
		//	{
		//		this.m_strSQL = "delete from max_nr_plots";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_pnr_plots"))
		//	{
		//		this.m_strSQL = "delete from max_pnr_plots";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_plots"))
		//	{
		//		this.m_strSQL = "delete from max_ti_imp_plots";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_pnr_plots"))
		//	{
		//		this.m_strSQL = "delete from max_ti_imp_pnr_plots";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_plots"))
		//	{
		//		this.m_strSQL = "delete from max_ci_imp_plots";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_pnr_plots"))
		//	{
		//		this.m_strSQL = "delete from max_ci_imp_pnr_plots";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_plots"))
		//	{
		//		this.m_strSQL = "delete from min_merch_plots";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_pnr_plots"))
		//	{
		//		this.m_strSQL = "delete from min_merch_pnr_plots";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}


		//	/*************************************************
		//	 **delete all records in the by ownership tables
		//	 *************************************************/
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_nr_sum_own"))
		//	{
		//		this.m_strSQL = "delete from max_nr_sum_own";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_pnr_sum_own"))
		//	{
		//		this.m_strSQL = "delete from max_pnr_sum_own";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_sum_own"))
		//	{
		//		this.m_strSQL = "delete from max_ti_imp_sum_own";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_pnr_sum_own"))
		//	{
		//		this.m_strSQL = "delete from max_ti_imp_pnr_sum_own";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_sum_own"))
		//	{
		//		this.m_strSQL = "delete from max_ci_imp_sum_own";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_pnr_sum_own"))
		//	{
		//		this.m_strSQL = "delete from max_ci_imp_pnr_sum_own";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_sum_own"))
		//	{
		//		this.m_strSQL = "delete from min_merch_sum_own";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_pnr_sum_own"))
		//	{
		//		this.m_strSQL = "delete from min_merch_pnr_sum_own";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}

		//	/******************************************************
		//	 **delete all records in the by processing site tables
		//	 ******************************************************/
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_nr_sum_psite"))
		//	{
		//		this.m_strSQL = "delete from max_nr_sum_psite";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_pnr_sum_psite"))
		//	{
		//		this.m_strSQL = "delete from max_pnr_sum_psite";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_sum_psite"))
		//	{
		//		this.m_strSQL = "delete from max_ti_imp_sum_psite";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_pnr_sum_psite"))
		//	{
		//		this.m_strSQL = "delete from max_ti_imp_pnr_sum_psite";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_sum_psite"))
		//	{
		//		this.m_strSQL = "delete from max_ci_imp_sum_psite";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_imp_pnr_sum_psite"))
		//	{
		//		this.m_strSQL = "delete from max_ci_imp_pnr_sum_psite";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_sum_psite"))
		//	{
		//		this.m_strSQL = "delete from min_merch_sum_psite";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}
		//	if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_pnr_sum_psite"))
		//	{
		//		this.m_strSQL = "delete from min_merch_pnr_sum_psite";
		//		oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
		//	}

		//	oAdo.CloseConnection(oAdo.m_OleDbConnection);
		//	oAdo=null;



		//}
		
		private void BestRxAcreageExpansionTableInsertForAirCurtainDestructionOld(string strTable,string strTypeField, string strRxField, string strWhereExpression)
		{
			string p = this.m_strPlotTable.Trim();
			string c = this.m_strCondTable.Trim();

			this.m_strSQL = "INSERT INTO " + strTable + " " + 
				"(" + strTypeField + ", owngrpcd,acres,merch_haul_cost_psite," + 
				"merch_haul_cents,merch_vol,merch_dollars_val," + 
				"chip_haul_cost_psite,chip_haul_cents," +
				"chip_yield,chip_dollars_val,net_rev,harv_costs," + 
				"haul_costs,ti_chg_acres,ci_chg_acres) " ;
			this.m_strSQL += " SELECT  " + 
				"best_rx_summary_air_dest.biosum_cond_id," + 
				"best_rx_summary_air_dest.owngrpcd," + 
				"best_rx_summary_air_dest.acres," +
				"null AS merch_haul_cost_psite," + 
				"null AS merch_haul_cents," + 
				"effective_product_yields_net_rev_costs_summary.merch_yield_cf * best_rx_summary_air_dest.acres AS merch_vol," + 
				"effective_product_yields_net_rev_costs_summary.merch_val_dpa * best_rx_summary_air_dest.acres AS merch_dollars_val," + 
				"null AS chip_haul_cost_psite," + 
				"null AS chip_haul_cents," + 
				"effective_product_yields_net_rev_costs_summary.chip_yield_gt * best_rx_summary_air_dest.acres AS chip_yield," + 
				"effective_product_yields_net_rev_costs_summary.chip_val_dpa * best_rx_summary_air_dest.acres AS chip_dollars_val," + 
				"null AS net_rev," + 
				"effective_product_yields_net_rev_costs_summary.harvest_onsite_cpa * best_rx_summary_air_dest.acres AS harv_costs," + 
				"null AS haul_costs," + 
				"effective.ti_change * best_rx_summary_air_dest.acres  AS ti_chg_acres," +
				"effective.ci_change * best_rx_summary_air_dest.acres  AS ci_chg_acres " + 
				"FROM (((best_rx_summary_air_dest  " + 
				"INNER JOIN effective_product_yields_net_rev_costs_summary  " + 
				"ON (best_rx_summary_air_dest.biosum_cond_id = effective_product_yields_net_rev_costs_summary.biosum_cond_id) AND " + 
				"(best_rx_summary_air_dest." + strRxField + " = effective_product_yields_net_rev_costs_summary.rx)) " + 
				"INNER JOIN effective " + 
				"ON (best_rx_summary_air_dest.biosum_cond_id = effective.biosum_cond_id) AND " +
				"(best_rx_summary_air_dest." + strRxField + "= effective.rx)) " + 
				"INNER JOIN (" + p + 
				" INNER JOIN " + c + 
				" ON " + p + ".biosum_plot_id = " + c + ".biosum_plot_id) " + 
				"ON " + c + ".biosum_cond_id = best_rx_summary_air_dest.biosum_cond_id)" ;
				
			if (strWhereExpression.Trim().Length > 0)
			{
				this.m_strSQL += " WHERE " + strWhereExpression + ";";
			}
			else
			{
				this.m_strSQL += ";";
			}

			//this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

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
        private void CompactMDB(string p_strMDBFileToCompact,System.Data.OleDb.OleDbConnection p_oConn)
        {
            string strConn = "";
            if (p_oConn != null)
            {
                strConn = p_oConn.ConnectionString;
                m_ado.CloseConnection(p_oConn);
            }
            m_oDao.CompactMDB(p_strMDBFileToCompact);
            System.Threading.Thread.Sleep(5000);
            if (strConn.Trim().Length > 0) m_ado.OpenConnection(strConn,ref p_oConn);
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

            this.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsCondPsiteTableName +
                " (BIOSUM_COND_ID,MERCH_PSITE_NUM,MERCH_PSITE_NAME," +
                "CHIP_PSITE_NUM,CHIP_PSITE_NAME) ";
            this.m_strSQL += "SELECT DISTINCT " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + ".BIOSUM_COND_ID, " +
                "MERCH_HAUL_PSITE,MERCH_HAUL_PSITE_NAME," +
                "CHIP_HAUL_PSITE,CHIP_HAUL_PSITE_NAME ";

            this.m_strSQL += " FROM " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName;
            this.m_strSQL += " INNER JOIN validcombos ON " +
                             " validcombos.biosum_cond_id=" + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName +
                             ".biosum_cond_id ";
            this.m_strSQL += " GROUP BY " + Tables.OptimizerScenarioResults.DefaultScenarioResultsPSiteAccessibleWorkTableName + 
                             ".BIOSUM_COND_ID,MERCH_HAUL_PSITE,MERCH_HAUL_PSITE_NAME,CHIP_HAUL_PSITE,CHIP_HAUL_PSITE_NAME";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert into cond_psite \r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");

                this.m_intError = this.m_ado.m_intError;

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
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 9;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;

            intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                ReferenceUserControlScenarioRun.listViewEx1, "Populate Context Database");

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);


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
            this.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName;
            this.m_strSQL += " SELECT RX, HarvestMethodLowSlope AS RX_HARVEST_METHOD_LOW, " + m_strHarvestMethodsTable +
                             ".HarvestMethodID AS RX_HARVEST_METHOD_LOW_ID, " + m_strHarvestMethodsTable + ".BIOSUM_CATEGORY AS RX_HARVEST_METHOD_LOW_CATEGORY, " +
                             m_strHarvestMethodsTable + ".top_limb_slope_status AS RX_HARVEST_METHOD_LOW_CATEGORY_DESCR, " +
                             "HarvestMethodSteepSlope AS RX_HARVEST_METHOD_STEEP, '" + strRxHarvestMethod + "' as USE_RX_HARVEST_METHOD_YN, " +
                             intSteepSlopePct + " AS STEEP_SLOPE_PCT";

            this.m_strSQL += " FROM " + this.m_strRxTable;
            this.m_strSQL += " INNER JOIN " + m_strHarvestMethodsTable + " ON " +
                             " TRIM(" + m_strHarvestMethodsTable + ".Method) = TRIM(RX.HarvestMethodLowSlope) ";
            this.m_strSQL += " WHERE STEEP_YN = 'N'";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\ninsert rx harvest methods into HARVEST_METHOD_REF \r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            this.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName;
            this.m_strSQL += " INNER JOIN " + m_strHarvestMethodsTable +
                            " ON TRIM(" + m_strHarvestMethodsTable + ".Method) = TRIM(" + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName + ".RX_HARVEST_METHOD_STEEP)";
            this.m_strSQL += " SET RX_HARVEST_METHOD_STEEP_ID = HarvestMethodID, RX_HARVEST_METHOD_STEEP_CATEGORY = BIOSUM_CATEGORY," +
                             "RX_HARVEST_METHOD_STEEP_CATEGORY_DESCR = top_limb_slope_status";
            this.m_strSQL += " WHERE STEEP_YN = 'Y'";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate HARVEST_METHOD_REF with properties of rx steep slope methods\r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            if (strRxHarvestMethod.Equals("N"))
            {
                this.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName;
                this.m_strSQL += " SET SCENARIO_HARVEST_METHOD_LOW = '" + m_oProcessorScenarioItem.m_oHarvestMethod.HarvestMethodLowSlope.Trim() +
                                 "', SCENARIO_HARVEST_METHOD_STEEP = '" + m_oProcessorScenarioItem.m_oHarvestMethod.HarvestMethodSteepSlope.Trim() + "'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate HARVEST_METHOD_REF for scenario harvest methods\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                this.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName;
                this.m_strSQL += " INNER JOIN " + m_strHarvestMethodsTable +
                                " ON TRIM(" + m_strHarvestMethodsTable + ".Method) = TRIM(" + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName + ".SCENARIO_HARVEST_METHOD_LOW)";
                this.m_strSQL += " SET SCENARIO_HARVEST_METHOD_LOW_ID = HarvestMethodID, SCENARIO_HARVEST_METHOD_LOW_CATEGORY = BIOSUM_CATEGORY," +
                                 "SCENARIO_HARVEST_METHOD_LOW_CATEGORY_DESCR = top_limb_slope_status";
                this.m_strSQL += " WHERE STEEP_YN = 'N'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate HARVEST_METHOD_REF with properties of scenario low slope methods\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

                this.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName;
                this.m_strSQL += " INNER JOIN " + m_strHarvestMethodsTable +
                                " ON TRIM(" + m_strHarvestMethodsTable + ".Method) = TRIM(" + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName + ".SCENARIO_HARVEST_METHOD_STEEP)";
                this.m_strSQL += " SET SCENARIO_HARVEST_METHOD_STEEP_ID = HarvestMethodID, SCENARIO_HARVEST_METHOD_STEEP_CATEGORY = BIOSUM_CATEGORY," +
                                 "SCENARIO_HARVEST_METHOD_STEEP_CATEGORY_DESCR = top_limb_slope_status";
                this.m_strSQL += " WHERE STEEP_YN = 'Y'";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate HARVEST_METHOD_REF with properties of scenario steep slope methods\r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            }

            this.m_strSQL = "UPDATE " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName +
                            " SET RX_HARVEST_METHOD_LOW = 'NA', RX_HARVEST_METHOD_LOW_ID = 999, RX_HARVEST_METHOD_LOW_CATEGORY = 999," +
                            " RX_HARVEST_METHOD_LOW_CATEGORY_DESCR = 'NA'," +
                            " RX_HARVEST_METHOD_STEEP = 'NA', RX_HARVEST_METHOD_STEEP_ID = 999, RX_HARVEST_METHOD_STEEP_CATEGORY = 999," +
                            " RX_HARVEST_METHOD_STEEP_CATEGORY_DESCR = 'NA'," +
                            " SCENARIO_HARVEST_METHOD_LOW = 'NA', SCENARIO_HARVEST_METHOD_LOW_ID = 999, SCENARIO_HARVEST_METHOD_LOW_CATEGORY = 999," +
                            " SCENARIO_HARVEST_METHOD_LOW_CATEGORY_DESCR = 'NA'," +
                            " SCENARIO_HARVEST_METHOD_STEEP = 'NA', SCENARIO_HARVEST_METHOD_STEEP_ID = 999, SCENARIO_HARVEST_METHOD_STEEP_CATEGORY = 999," +
                            " SCENARIO_HARVEST_METHOD_STEEP_CATEGORY_DESCR = 'NA'" +
                            " WHERE RX = '999'";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate HARVEST_METHOD_REF for rx 999 (grow-only) \r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            this.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsRxPackageRefTableName;
            this.m_strSQL += " SELECT RXPACKAGE, " + this.m_strRxPackageTable + ".description, simyear1_rx, simyear2_rx, simyear3_rx, simyear4_rx ";
            this.m_strSQL += " FROM " + this.m_strRxPackageTable;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nInsert records into RXPACKAGE_REF \r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            string[] arrRxCycle = new string[] { "SIMYEAR1_RX", "SIMYEAR2_RX", "SIMYEAR3_RX", "SIMYEAR4_RX" };
            foreach (string strRxCycle in arrRxCycle) 
            {
                this.m_strSQL = "UPDATE ((" + Tables.OptimizerScenarioResults.DefaultScenarioResultsRxPackageRefTableName;
                this.m_strSQL += " INNER JOIN " + this.m_strRxTable + " ON " + Tables.OptimizerScenarioResults.DefaultScenarioResultsRxPackageRefTableName + 
                                 "." + strRxCycle + " = " + this.m_strRxTable + ".rx) " +
                                 " INNER JOIN fvs_rx_category ON " + this.m_strRxTable + ".catid = fvs_rx_category.catid)" +
                                 " INNER JOIN fvs_rx_subcategory ON (" + this.m_strRxTable + ".catid = fvs_rx_subcategory.catid)" +
                                 " AND (" + this.m_strRxTable + ".subcatid = fvs_rx_subcategory.subcatid)";
                this.m_strSQL += " SET " + strRxCycle + "_DESCRIPTION = " + this.m_strRxTable + ".DESCRIPTION, " +
                                 strRxCycle + "_CATEGORY = FVS_RX_CATEGORY.DESC, " +
                                 strRxCycle + "_SUBCATEGORY = FVS_RX_SUBCATEGORY.DESC";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate RXPACKAGE_REF for " + strRxCycle + " \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            }

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

                this.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsDiameterSpeciesGroupRefTableName +
                                " (DBH_CLASS_NUM, DBH_RANGE_INCHES, SPP_GRP_CODE, SPP_GRP, TO_CHIPS, MERCH_VAL_DpCF, VALUE_IF_CHIPPED_DpGT)" +
                                " VALUES (" + oItem.DiameterGroupId + ", '" + oItem.DbhGroup + "'," + oItem.SpeciesGroupId + ",'" + oItem.SpeciesGroup + "','" +
                                strEnergyWood + "', " + oItem.MerchDollarPerCubicFootValue + ", " +
                                dblChippedValue + ")";

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate DIAMETER_SPP_GRP_REF_C table \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            }

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            this.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsFvsWeightedVariablesRefTableName;
            this.m_strSQL += " SELECT VARIABLE_NAME, VARIABLE_DESCRIPTION, BASELINE_RXPACKAGE, VARIABLE_SOURCE," +
                             " weight_1_pre, weight_1_post, weight_2_pre, weight_2_post, weight_3_pre, weight_3_post, weight_4_pre, weight_4_post";
            this.m_strSQL += " FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                             " INNER JOIN " + Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + " ON " +
                             Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + ".calculated_variables_id = " +
                             Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + ".ID";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate FVS_WEIGHTED_VARIABLES_REF table \r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            this.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconWeightedVariablesRefTableName;
            this.m_strSQL += " SELECT VARIABLE_NAME, VARIABLE_DESCRIPTION, VARIABLE_SOURCE," +
                             " weight as CYCLE_1_WEIGHT";
            this.m_strSQL += " FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                             " INNER JOIN " + Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName + " ON " +
                             Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName + ".calculated_variables_id = " +
                             Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + ".ID" +
                             " WHERE rxcycle = '1'";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate ECON_WEIGHTED_VARIABLES_REF table \r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            arrRxCycle = new string[] { "2", "3", "4"};
            
            foreach (string strRxCycle in arrRxCycle)
            {
                string strFieldName = "CYCLE_" + strRxCycle + "_WEIGHT";
                this.m_strSQL = "UPDATE (" + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconWeightedVariablesRefTableName;
                this.m_strSQL += " INNER JOIN " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + " ON " +
                                  Tables.OptimizerScenarioResults.DefaultScenarioResultsEconWeightedVariablesRefTableName + ".VARIABLE_NAME = " +
                                  Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + ".VARIABLE_NAME)" +
                                 " INNER JOIN " + Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName + " ON " +
                                 Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + ".ID = " +
                                 Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName + ".calculated_variables_id";
                this.m_strSQL += " SET " + Tables.OptimizerScenarioResults.DefaultScenarioResultsEconWeightedVariablesRefTableName + "." + strFieldName 
                                 + " = calculated_econ_variables_definition.weight";
                this.m_strSQL += " where " + Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName + ".rxcycle = '" + strRxCycle + "'";
                
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nUpdate " + strFieldName + " field \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);
            }

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            this.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName + "_C" +
                            " SELECT * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName  +
                            " WHERE SCENARIO_ID = '" + this.m_oProcessorScenarioItem.ScenarioId + "'";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate SCENARIO_ADDITIONAL_ HARVEST_COSTS table \r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            this.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioResults.DefaultScenarioResultsSpeciesGroupRefTableName +
                " SELECT species_group AS SPP_GRP_CD, COMMON_NAME, spcd AS FIA_SPCD" +
                " FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName +
                " WHERE SCENARIO_ID = '" + this.m_oProcessorScenarioItem.ScenarioId + "'";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate spp_grp_ref_C table \r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);

            this.m_strSQL = "INSERT INTO " + Tables.FVS.DefaultRxHarvestCostColumnsTableName + "_C" +
                            " SELECT * FROM " + Tables.FVS.DefaultRxHarvestCostColumnsTableName;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nPopulate rx_harvest_cost_columns_C table \r\n");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn, this.m_strSQL);


            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            // harvest_costs table
            string strContextConn = this.m_ado.getMDBConnString(this.m_strContextDbPathAndFile, "", "");
            using (var contextConn = new OleDbConnection(strContextConn))
            {
                contextConn.Open();
                this.m_strSQL = "SELECT " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + ".*" +
                                " INTO " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + "_C" +
                                " FROM " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nCreate and populate harvest_costs_C table \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

                this.m_ado.SqlNonQuery(strContextConn, this.m_strSQL);


                this.m_strSQL = "SELECT " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + ".*" +
                                " INTO " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + "_C" +
                                " FROM " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nCreate and populate tree_vol_val_by_species_diam_groups_C table \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

                this.m_ado.SqlNonQuery(strContextConn, this.m_strSQL);

                this.m_strSQL = "DROP TABLE " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nDrop harvest_costs table from context db \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

                this.m_ado.SqlNonQuery(strContextConn, this.m_strSQL);

                this.m_strSQL = "DROP TABLE " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nDrop tree_vol_val_by_species_diam_groups table from context db \r\n");
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + this.m_strSQL + "\r\n");

                this.m_ado.SqlNonQuery(strContextConn, this.m_strSQL);
            }
            
            FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

            if (this.m_ado.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");

                this.m_intError = this.m_ado.m_intError;

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

        public void FvsContextReferenceTables()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//FvsContextReferenceTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 10;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;

            intListViewIndex = FIA_Biosum_Manager.uc_optimizer_scenario_run.GetListViewItemIndex(
                ReferenceUserControlScenarioRun.listViewEx1, "Populate FVS PRE-POST Context Database");

            FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem = intListViewIndex;
            FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic = (ProgressBarBasic.ProgressBarBasic)ReferenceUserControlScenarioRun.listViewEx1.GetEmbeddedControl(1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.EnsureListViewExItemVisible(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "Selected", true);
            frmMain.g_oDelegate.SetListViewItemPropertyValue(ReferenceUserControlScenarioRun.listViewEx1, FIA_Biosum_Manager.RunOptimizer.g_intCurrentListViewItem, "focused", true);


            dao_data_access oDao = new dao_data_access();
            ado_data_access oAdo = new ado_data_access();

            // Prepare the database for the next run
            if (!System.IO.File.Exists(this.m_strFvsContextDbPathAndFile))
            {
                oDao.CreateMDB(this.m_strFvsContextDbPathAndFile);
            }
            else
            {
                string[] arrTableNames = new string[0];
                oDao.getTableNames(this.m_strFvsContextDbPathAndFile, ref arrTableNames);
                foreach (string strTableName in arrTableNames)
                {
                    if (!String.IsNullOrEmpty(strTableName))
                    {
                        oDao.DeleteTableFromMDB(this.m_strFvsContextDbPathAndFile, strTableName);
                    }
                }
            }

            _uc_scenario_run.uc_filesize_monitor3.BeginMonitoringFile(
              m_strFvsContextDbPathAndFile, 2000000000, "2GB");
            _uc_scenario_run.uc_filesize_monitor3.Information = "FVS Context Reference Tables";

            // Add FVS PRE-POST tables
            m_oRxTools.CreateTableLinksToFVSPrePostTables(this.m_strFvsContextDbPathAndFile);
            m_intError = m_oRxTools.m_intError;
            if (this.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");
                return;
            }

            string[] arrFvsTableNames = new string[0];
            oDao.getTableNames(this.m_strFvsContextDbPathAndFile, ref arrFvsTableNames);

            string strConn = oAdo.getMDBConnString(this.m_strFvsContextDbPathAndFile, "", "");
            using (OleDbConnection oConn = new OleDbConnection(strConn))
            {
                oConn.Open();
                foreach (string strFvsTableName in arrFvsTableNames)
                {
                    if (!String.IsNullOrEmpty(strFvsTableName))
                    {
                        oAdo.m_strSQL = "SELECT * INTO " + strFvsTableName + "_C" +
                        " from " + strFvsTableName;

                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + oAdo.m_strSQL + "\r\n");
                        oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                        oAdo.m_strSQL = "DROP TABLE " + strFvsTableName;
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + oAdo.m_strSQL + "\r\n");
                        oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                        FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                    }
                }
                CompactMDB(this.m_strFvsContextDbPathAndFile, oConn);               
            }

            if (oAdo.m_intError != 0)
            {
                FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic.TextColor = Color.Red;
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "!!Error!!");

                this.m_intError = oAdo.m_intError;

                return;
            }

            _uc_scenario_run.uc_filesize_monitor3.EndMonitoringFile();

            if (this.UserCancel(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic)) return;


            if (this.m_intError == 0)
            {
                FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermText(FIA_Biosum_Manager.RunOptimizer.g_oCurrentProgressBarBasic, "Done");

            }

            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
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

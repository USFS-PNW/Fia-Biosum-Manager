using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// allows user to run the treatment optimizer scenario and view results
	/// </summary>
	public class frmRunOptimizerScenario : System.Windows.Forms.Form
	{
		public System.Windows.Forms.Label lblMsg;
		public System.Windows.Forms.Button btnCancel;
		public System.Windows.Forms.ProgressBar progressBar1;
		private System.Windows.Forms.ImageList imageList1;
		public System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.ComponentModel.IContainer components;
		public System.Windows.Forms.Label lblProcBestRxOwner;
		public System.Windows.Forms.Label lblProcBestRxPSite;
		public System.Windows.Forms.Label lblProcBestRxPlot;
		public System.Windows.Forms.Label lblProcBestRx;
		public System.Windows.Forms.Label lblSumWoodProducts;
		public System.Windows.Forms.Label lblProcEffective;
		public System.Windows.Forms.Label lblProcValidCombos;
		public System.Windows.Forms.Label lblProcSumTree;
		public System.Windows.Forms.Label lblProcTravelTimes;
		public System.Windows.Forms.Label lblProcAccessible;
		public System.Windows.Forms.CheckBox chkProcSumTree;
		public System.Windows.Forms.CheckBox chkProcTravelTimes;
		public System.Windows.Forms.Button btnViewScenarioTables;
		public FIA_Biosum_Manager.frmOptimizerScenario m_frmScenario;
		private FIA_Biosum_Manager.frmGridView m_frmGridView;
		public System.Windows.Forms.Button btnViewAuditTables;
		private int m_intError=0;
		public System.Data.DataSet m_ds;
		public System.Data.OleDb.OleDbConnection m_conn;
		public System.Data.OleDb.OleDbDataAdapter m_da;
		public System.Windows.Forms.CheckBox chkAuditTables;
		public string m_strCustomPlotSQL="";
		public System.Windows.Forms.Button btnViewLog;
		private System.Windows.Forms.Button btnSelectAll;
		private System.Windows.Forms.Button btnClear;
		private string m_strSQL;
		private System.Windows.Forms.Button btnHelp;
		private System.Windows.Forms.Label label1;
		public bool m_bUserCancel=false;

	    

		public frmRunOptimizerScenario(FIA_Biosum_Manager.frmOptimizerScenario p_frmScenario)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			this.m_frmScenario = p_frmScenario;
			this.Enabled=true;

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmRunOptimizerScenario));
            this.lblMsg = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnViewLog = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.chkAuditTables = new System.Windows.Forms.CheckBox();
            this.lblProcBestRxOwner = new System.Windows.Forms.Label();
            this.lblProcBestRxPSite = new System.Windows.Forms.Label();
            this.lblProcBestRxPlot = new System.Windows.Forms.Label();
            this.lblProcBestRx = new System.Windows.Forms.Label();
            this.lblSumWoodProducts = new System.Windows.Forms.Label();
            this.lblProcEffective = new System.Windows.Forms.Label();
            this.lblProcValidCombos = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnViewScenarioTables = new System.Windows.Forms.Button();
            this.btnViewAuditTables = new System.Windows.Forms.Button();
            this.btnHelp = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnSelectAll = new System.Windows.Forms.Button();
            this.lblProcSumTree = new System.Windows.Forms.Label();
            this.lblProcTravelTimes = new System.Windows.Forms.Label();
            this.lblProcAccessible = new System.Windows.Forms.Label();
            this.chkProcSumTree = new System.Windows.Forms.CheckBox();
            this.chkProcTravelTimes = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblMsg
            // 
            this.lblMsg.Enabled = false;
            this.lblMsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMsg.Location = new System.Drawing.Point(9, 361);
            this.lblMsg.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(693, 15);
            this.lblMsg.TabIndex = 5;
            this.lblMsg.Text = "lblMsg";
            this.lblMsg.Visible = false;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(324, 413);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(77, 23);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Enabled = false;
            this.progressBar1.Location = new System.Drawing.Point(9, 383);
            this.progressBar1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(699, 23);
            this.progressBar1.TabIndex = 3;
            this.progressBar1.Visible = false;
            this.progressBar1.Click += new System.EventHandler(this.progressBar1_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "");
            this.imageList1.Images.SetKeyName(1, "");
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnViewLog);
            this.groupBox1.Controls.Add(this.groupBox3);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.btnViewScenarioTables);
            this.groupBox1.Controls.Add(this.btnViewAuditTables);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.progressBar1);
            this.groupBox1.Controls.Add(this.lblMsg);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Size = new System.Drawing.Size(448, 313);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
            // 
            // btnViewLog
            // 
            this.btnViewLog.Location = new System.Drawing.Point(579, 36);
            this.btnViewLog.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnViewLog.Name = "btnViewLog";
            this.btnViewLog.Size = new System.Drawing.Size(128, 19);
            this.btnViewLog.TabIndex = 30;
            this.btnViewLog.Text = "View Log File";
            this.btnViewLog.Click += new System.EventHandler(this.btnViewLog_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.chkAuditTables);
            this.groupBox3.Controls.Add(this.lblProcBestRxOwner);
            this.groupBox3.Controls.Add(this.lblProcBestRxPSite);
            this.groupBox3.Controls.Add(this.lblProcBestRxPlot);
            this.groupBox3.Controls.Add(this.lblProcBestRx);
            this.groupBox3.Controls.Add(this.lblSumWoodProducts);
            this.groupBox3.Controls.Add(this.lblProcEffective);
            this.groupBox3.Controls.Add(this.lblProcValidCombos);
            this.groupBox3.Controls.Add(this.label10);
            this.groupBox3.Controls.Add(this.label9);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Location = new System.Drawing.Point(9, 163);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox3.Size = new System.Drawing.Size(699, 188);
            this.groupBox3.TabIndex = 27;
            this.groupBox3.TabStop = false;
            // 
            // chkAuditTables
            // 
            this.chkAuditTables.Checked = true;
            this.chkAuditTables.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkAuditTables.Location = new System.Drawing.Point(444, 15);
            this.chkAuditTables.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.chkAuditTables.Name = "chkAuditTables";
            this.chkAuditTables.Size = new System.Drawing.Size(239, 16);
            this.chkAuditTables.TabIndex = 14;
            this.chkAuditTables.Text = "Populate Valid Combination Audit Data";
            // 
            // lblProcBestRxOwner
            // 
            this.lblProcBestRxOwner.Location = new System.Drawing.Point(9, 156);
            this.lblProcBestRxOwner.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblProcBestRxOwner.Name = "lblProcBestRxOwner";
            this.lblProcBestRxOwner.Size = new System.Drawing.Size(85, 15);
            this.lblProcBestRxOwner.TabIndex = 13;
            this.lblProcBestRxOwner.Text = "Not Done";
            // 
            // lblProcBestRxPSite
            // 
            this.lblProcBestRxPSite.Location = new System.Drawing.Point(9, 133);
            this.lblProcBestRxPSite.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblProcBestRxPSite.Name = "lblProcBestRxPSite";
            this.lblProcBestRxPSite.Size = new System.Drawing.Size(85, 15);
            this.lblProcBestRxPSite.TabIndex = 12;
            this.lblProcBestRxPSite.Text = "Not Done";
            // 
            // lblProcBestRxPlot
            // 
            this.lblProcBestRxPlot.Location = new System.Drawing.Point(9, 109);
            this.lblProcBestRxPlot.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblProcBestRxPlot.Name = "lblProcBestRxPlot";
            this.lblProcBestRxPlot.Size = new System.Drawing.Size(85, 15);
            this.lblProcBestRxPlot.TabIndex = 11;
            this.lblProcBestRxPlot.Text = "Not Done";
            // 
            // lblProcBestRx
            // 
            this.lblProcBestRx.Location = new System.Drawing.Point(9, 83);
            this.lblProcBestRx.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblProcBestRx.Name = "lblProcBestRx";
            this.lblProcBestRx.Size = new System.Drawing.Size(85, 16);
            this.lblProcBestRx.TabIndex = 10;
            this.lblProcBestRx.Text = "Not Done";
            // 
            // lblSumWoodProducts
            // 
            this.lblSumWoodProducts.Location = new System.Drawing.Point(9, 57);
            this.lblSumWoodProducts.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblSumWoodProducts.Name = "lblSumWoodProducts";
            this.lblSumWoodProducts.Size = new System.Drawing.Size(85, 16);
            this.lblSumWoodProducts.TabIndex = 9;
            this.lblSumWoodProducts.Text = "Not Done";
            // 
            // lblProcEffective
            // 
            this.lblProcEffective.Location = new System.Drawing.Point(9, 36);
            this.lblProcEffective.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblProcEffective.Name = "lblProcEffective";
            this.lblProcEffective.Size = new System.Drawing.Size(85, 15);
            this.lblProcEffective.TabIndex = 8;
            this.lblProcEffective.Text = "Not Done";
            // 
            // lblProcValidCombos
            // 
            this.lblProcValidCombos.Location = new System.Drawing.Point(9, 15);
            this.lblProcValidCombos.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblProcValidCombos.Name = "lblProcValidCombos";
            this.lblProcValidCombos.Size = new System.Drawing.Size(85, 16);
            this.lblProcValidCombos.TabIndex = 7;
            this.lblProcValidCombos.Text = "Not Done";
            // 
            // label10
            // 
            this.label10.Location = new System.Drawing.Point(101, 156);
            this.label10.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(581, 23);
            this.label10.TabIndex = 6;
            this.label10.Text = "Summarize Most Effective Treatment Yields, Revenue, Costs, And Acres By Land Owne" +
    "rship Groups";
            // 
            // label9
            // 
            this.label9.Location = new System.Drawing.Point(101, 133);
            this.label9.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(589, 15);
            this.label9.TabIndex = 5;
            this.label9.Text = "Summarize Most Effective Treatment Yields, Revenue, Costs, And Acres By Wood Proc" +
    "essing Facility";
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(101, 109);
            this.label8.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(581, 15);
            this.label8.TabIndex = 4;
            this.label8.Text = "Summarize Most Effective Treatment Yields, Revenue, Costs, And Acres By Stand";
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(101, 76);
            this.label7.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(590, 23);
            this.label7.TabIndex = 3;
            this.label7.Text = "Find Most Effective Treatment For Torch And Crown Index Improvement, Maximum  Rev" +
    "enue, And Minimum Merchantable Wood Removal";
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(101, 55);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(564, 16);
            this.label6.TabIndex = 2;
            this.label6.Text = "Summarize Wood Product Volume Yields, Costs, And Net Revenue For A Stand And Trea" +
    "tment";
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(101, 35);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(359, 15);
            this.label5.TabIndex = 1;
            this.label5.Text = "Identify Fuel And Fire Effective Treatments ";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(101, 15);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(325, 15);
            this.label4.TabIndex = 0;
            this.label4.Text = "Apply User Defined Filters And Get Valid Plot Combinations";
            // 
            // lblTitle
            // 
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(9, 15);
            this.lblTitle.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(230, 23);
            this.lblTitle.TabIndex = 26;
            this.lblTitle.Text = "Run Scenario";
            // 
            // btnViewScenarioTables
            // 
            this.btnViewScenarioTables.Location = new System.Drawing.Point(451, 36);
            this.btnViewScenarioTables.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnViewScenarioTables.Name = "btnViewScenarioTables";
            this.btnViewScenarioTables.Size = new System.Drawing.Size(128, 19);
            this.btnViewScenarioTables.TabIndex = 11;
            this.btnViewScenarioTables.Text = "View Results Tables";
            this.btnViewScenarioTables.Click += new System.EventHandler(this.btnViewScenarioTables_Click);
            // 
            // btnViewAuditTables
            // 
            this.btnViewAuditTables.Location = new System.Drawing.Point(579, 15);
            this.btnViewAuditTables.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnViewAuditTables.Name = "btnViewAuditTables";
            this.btnViewAuditTables.Size = new System.Drawing.Size(128, 20);
            this.btnViewAuditTables.TabIndex = 10;
            this.btnViewAuditTables.Text = "View Audit Data";
            this.btnViewAuditTables.Click += new System.EventHandler(this.btnViewAuditTables_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.Location = new System.Drawing.Point(9, 429);
            this.btnHelp.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(85, 23);
            this.btnHelp.TabIndex = 9;
            this.btnHelp.Text = "Help";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.btnClear);
            this.groupBox2.Controls.Add(this.btnSelectAll);
            this.groupBox2.Controls.Add(this.lblProcSumTree);
            this.groupBox2.Controls.Add(this.lblProcTravelTimes);
            this.groupBox2.Controls.Add(this.lblProcAccessible);
            this.groupBox2.Controls.Add(this.chkProcSumTree);
            this.groupBox2.Controls.Add(this.chkProcTravelTimes);
            this.groupBox2.Location = new System.Drawing.Point(9, 61);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Size = new System.Drawing.Size(699, 93);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Optional:  Checked Boxes Will Execute";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(100, 23);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(421, 15);
            this.label1.TabIndex = 8;
            this.label1.Text = "Determine If Plot And Conditions Are Accessible For Treatment And Harvest";
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(580, 55);
            this.btnClear.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(69, 23);
            this.btnClear.TabIndex = 7;
            this.btnClear.Text = "Clear";
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnSelectAll
            // 
            this.btnSelectAll.Location = new System.Drawing.Point(580, 23);
            this.btnSelectAll.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnSelectAll.Name = "btnSelectAll";
            this.btnSelectAll.Size = new System.Drawing.Size(69, 23);
            this.btnSelectAll.TabIndex = 6;
            this.btnSelectAll.Text = "Select All";
            this.btnSelectAll.Click += new System.EventHandler(this.btnSelectAll_Click);
            // 
            // lblProcSumTree
            // 
            this.lblProcSumTree.Location = new System.Drawing.Point(9, 61);
            this.lblProcSumTree.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblProcSumTree.Name = "lblProcSumTree";
            this.lblProcSumTree.Size = new System.Drawing.Size(85, 15);
            this.lblProcSumTree.TabIndex = 5;
            this.lblProcSumTree.Text = "Not Done";
            // 
            // lblProcTravelTimes
            // 
            this.lblProcTravelTimes.Location = new System.Drawing.Point(9, 43);
            this.lblProcTravelTimes.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblProcTravelTimes.Name = "lblProcTravelTimes";
            this.lblProcTravelTimes.Size = new System.Drawing.Size(85, 16);
            this.lblProcTravelTimes.TabIndex = 4;
            this.lblProcTravelTimes.Text = "Not Done";
            // 
            // lblProcAccessible
            // 
            this.lblProcAccessible.Location = new System.Drawing.Point(9, 23);
            this.lblProcAccessible.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblProcAccessible.Name = "lblProcAccessible";
            this.lblProcAccessible.Size = new System.Drawing.Size(77, 15);
            this.lblProcAccessible.TabIndex = 3;
            this.lblProcAccessible.Text = "Not Done";
            // 
            // chkProcSumTree
            // 
            this.chkProcSumTree.Checked = true;
            this.chkProcSumTree.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkProcSumTree.Location = new System.Drawing.Point(101, 63);
            this.chkProcSumTree.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.chkProcSumTree.Name = "chkProcSumTree";
            this.chkProcSumTree.Size = new System.Drawing.Size(231, 15);
            this.chkProcSumTree.TabIndex = 2;
            this.chkProcSumTree.Text = "Sum Tree Yields, Volume, And Value";
            // 
            // chkProcTravelTimes
            // 
            this.chkProcTravelTimes.Checked = true;
            this.chkProcTravelTimes.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkProcTravelTimes.Location = new System.Drawing.Point(101, 43);
            this.chkProcTravelTimes.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.chkProcTravelTimes.Name = "chkProcTravelTimes";
            this.chkProcTravelTimes.Size = new System.Drawing.Size(402, 16);
            this.chkProcTravelTimes.TabIndex = 1;
            this.chkProcTravelTimes.Text = "Get Least Expensive Route From Plot To Wood Processing Facility";
            // 
            // frmRunOptimizerScenario
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(448, 313);
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.MaximizeBox = false;
            this.Name = "frmRunOptimizerScenario";
            this.Text = "Treatment Optimizer Run Scenario";
            this.Resize += new System.EventHandler(this.frmRunCoreScenario_Resize);
            this.groupBox1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void progressBar1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			DialogResult result;

			if (this.btnCancel.Text.Trim().ToUpper() == "CANCEL")
			{
				result =  MessageBox.Show("Cancel Running The Scenario (Y/N)?","Cancel Process", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
				switch (result) 
				{
					case DialogResult.Yes:
						this.btnCancel.Text = "Start";
						this.m_bUserCancel=true;
						return;
					case DialogResult.No:
						return;
				}
			}
			else
			{
				this.progressBar1.Visible=false;
				this.lblMsg.Text = "";
				this.lblProcAccessible.ForeColor = System.Drawing.Color.Black;
				this.lblProcAccessible.Text = "Not Done";
				this.lblProcBestRx.ForeColor = System.Drawing.Color.Black;
				this.lblProcBestRx.Text = "Not Done";
				this.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Black;
				this.lblProcBestRxOwner.Text = "Not Done";
				this.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Black;
				this.lblProcBestRxPlot.Text = "Not Done";
				this.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Black;
				this.lblProcBestRxPSite.Text = "Not Done";
				this.lblProcEffective.ForeColor = System.Drawing.Color.Black;
				this.lblProcEffective.Text = "Not Done";
				this.lblProcSumTree.ForeColor = System.Drawing.Color.Black;
				this.lblProcSumTree.Text = "Not Done";
				this.lblProcTravelTimes.ForeColor = System.Drawing.Color.Black;
				this.lblProcTravelTimes.Text = "Not Done";
				this.lblProcValidCombos.ForeColor = System.Drawing.Color.Black;
				this.lblProcValidCombos.Text = "Not Done";
				this.lblSumWoodProducts.ForeColor = System.Drawing.Color.Black;
				this.lblSumWoodProducts.Text = "Not Done";
				this.Refresh();
				this.val_CoreRunData();
				if (this.m_intError==0)
				{
					this.btnCancel.Text = "Cancel";
					this.btnCancel.Refresh();
					this.btnViewAuditTables.Enabled=false;
					this.btnViewScenarioTables.Enabled=false;
					this.btnViewLog.Enabled=false;
					this.m_bUserCancel=false;
				}
				else
				{
					if (this.m_frmScenario.WindowState == System.Windows.Forms.FormWindowState.Minimized)
						this.m_frmScenario.WindowState = System.Windows.Forms.FormWindowState.Normal;
					this.m_frmScenario.Focus();

				}
			   
              
			  
			}


		}

		private void btnViewScenarioTables_Click(object sender, System.EventArgs e)
		{
			this.viewScenarioTables();
		
		}
		
		/// <summary>
		/// every scenario_results.mdb table is viewed in a uc_gridview control
		/// </summary>
		private void viewScenarioTables()
		{
			string strMDBPathAndFile="";
			string strConn="";
			string strSQL="";
			string[] strTableNames;
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();
			
			
			strMDBPathAndFile = m_frmScenario.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";

			int intCount = p_dao.getTableNames(strMDBPathAndFile,ref strTableNames);
			if (p_dao.m_intError==0)
			{
				if (intCount > 0)
				{
					this.progressBar1.Minimum=0;
					this.progressBar1.Maximum=intCount;
					this.progressBar1.Value=0;
					this.progressBar1.Visible=true;
					this.lblMsg.Text = "";
					this.lblMsg.Visible=true;
					if (strMDBPathAndFile.Substring(strMDBPathAndFile.Trim().Length - 6,6).ToUpper()==".ACCDB")
						strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + 
							strMDBPathAndFile + ";User Id=admin;Password=;";
					else
						strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";
					this.m_frmGridView = new frmGridView();
					this.m_frmGridView.Text = "Core Analysis: Run Scenario Results ("  + this.m_frmScenario.uc_scenario1.txtScenarioId.Text.Trim() + ")";
					for (int x=0; x <= intCount-1;x++)
					{
						this.lblMsg.Text = strTableNames[x];
						this.lblMsg.Refresh();
						strSQL = "select * from " + strTableNames[x].Trim();
						this.m_frmGridView.LoadDataSet(strConn,strSQL,strTableNames[x].Trim());
						this.progressBar1.Value = x + 1;

					}
					this.progressBar1.Visible=false;
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

        /// <summary>
        /// each audit table is viewed in a uc_gridview control
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
	
		private void btnViewAuditTables_Click(object sender, System.EventArgs e)
		{
	
			string strMDBPathAndFile="";
			string strConn="";
			string strTable="";
			FIA_Biosum_Manager.Datasource p_datasource = new Datasource(((frmMain)this.ParentForm).frmProject.uc_project1.txtRootDirectory.Text.Trim(),
				                                                        this.m_frmScenario.uc_scenario1.txtScenarioId.Text.Trim());
			this.m_frmGridView = new frmGridView();
			this.m_frmGridView.Text = "Core Analysis: Audit";

			//plot and condition record audit
			strMDBPathAndFile = p_datasource.getFullPathAndFile("PLOT AND CONDITION RECORD AUDIT");
			strTable = p_datasource.getValidDataSourceTableName("PLOT AND CONDITION RECORD AUDIT");

			if (strMDBPathAndFile.Substring(strMDBPathAndFile.Trim().Length - 6,6).ToUpper()==".ACCDB")
				strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + 
					strMDBPathAndFile + ";User Id=admin;Password=;";
			else
			    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";

			this.m_frmGridView.LoadDataSet(strConn,"select * from " + strTable,strTable);

			//plot,condition, and rx record audit
			strMDBPathAndFile = p_datasource.getFullPathAndFile("PLOT, CONDITION AND TREATMENT RECORD AUDIT");
			strTable = p_datasource.getValidDataSourceTableName("PLOT, CONDITION AND TREATMENT RECORD AUDIT");

			if (strMDBPathAndFile.Substring(strMDBPathAndFile.Trim().Length - 6,6).ToUpper()==".ACCDB")
				strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + 
					strMDBPathAndFile + ";User Id=admin;Password=;";
			else
				strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";
			this.m_frmGridView.LoadDataSet(strConn,"select * from " + strTable,strTable);
			this.m_frmGridView.TileGridViews();
			this.m_frmGridView.Show();
			this.m_frmGridView.Focus();
			p_datasource = null;
		}

		private void frmRunCoreScenario_Resize(object sender, System.EventArgs e)
		{
		
		}

		/// <summary>
		/// validate each component required for running core analysis
		/// </summary>
		private void val_CoreRunData()
		{
			this.m_intError=0;
			if (this.m_intError==0) this.m_intError = this.m_frmScenario.uc_scenario_owner_groups1.ValInput();
			if (this.m_intError==0)  this.m_intError = this.m_frmScenario.uc_scenario_costs1.val_costs();
			if (this.m_intError==0) this.m_intError = this.m_frmScenario.uc_scenario_psite1.val_psites();
            
          
			if (this.m_intError==0)
			
			{
				
				/***************************************************************************
					 **make sure all the scenario datasource tables and files are available
					 **and ready for use
					 ***************************************************************************/
					
				if (this.m_frmScenario.m_ldatasourcefirsttime==true)
				{
					this.m_frmScenario.uc_datasource1.populate_listview_grid();
					this.m_frmScenario.m_ldatasourcefirsttime=false;
				}
				this.m_intError = this.m_frmScenario.uc_datasource1.val_datasources();
				if (this.m_intError ==0)
				{
					this.m_frmScenario.SaveRuleDefinitions();
				}
			}
				

		}

		/// <summary>
		/// view the log file created in the most recent run
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnViewLog_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process proc = new System.Diagnostics.Process();
			proc.StartInfo.UseShellExecute = true;
			try
			{
				proc.StartInfo.FileName = this.m_frmScenario.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\runlog.txt";
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

		private void btnSelectAll_Click(object sender, System.EventArgs e)
		{
			this.chkProcTravelTimes.Checked=true;
			this.chkProcSumTree.Checked=true;
		}

		private void btnClear_Click(object sender, System.EventArgs e)
		{
			if (this.chkProcTravelTimes.Enabled) this.chkProcTravelTimes.Checked=false;
			if (this.chkProcSumTree.Enabled) this.chkProcSumTree.Checked=false;

		}

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.btnHelp.Top = this.groupBox1.Height - this.btnHelp.Height - 5;
				this.btnHelp.Left = 4;
				
				this.groupBox2.Width = this.groupBox1.Width - (int)(this.groupBox2.Left * 2);
				this.groupBox3.Width = this.groupBox2.Width;
				this.lblMsg.Width = this.groupBox1.Width - (int)(this.lblMsg.Left * 2);
				this.progressBar1.Width = this.groupBox1.Width - (int)(this.progressBar1.Left * 2);
				this.btnViewAuditTables.Left = this.groupBox2.Width - this.btnViewAuditTables.Width + this.groupBox2.Left ;
				this.btnViewLog.Left = this.btnViewAuditTables.Left;
				this.btnViewScenarioTables.Left = this.btnViewLog.Left - this.btnViewScenarioTables.Width;
				this.btnCancel.Left = (int)(this.progressBar1.Width * .50) - (int)(this.btnCancel.Width * .50);
			}
			catch
			{
			}
		}
	}
}

	

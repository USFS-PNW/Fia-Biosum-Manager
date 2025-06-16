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
		public RunCoreOld m_oRunCore;
		public System.Windows.Forms.CheckBox chkAuditTables;
		public System.Windows.Forms.Button btnAccess;
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
            this.btnAccess = new System.Windows.Forms.Button();
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
            this.groupBox1.Controls.Add(this.btnAccess);
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
            // btnAccess
            // 
            this.btnAccess.Enabled = false;
            this.btnAccess.Location = new System.Drawing.Point(451, 15);
            this.btnAccess.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnAccess.Name = "btnAccess";
            this.btnAccess.Size = new System.Drawing.Size(128, 20);
            this.btnAccess.TabIndex = 29;
            this.btnAccess.Text = "Microsoft Access";
            this.btnAccess.Click += new System.EventHandler(this.btnAccess_Click);
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
					this.btnAccess.Enabled=false;
					this.btnViewLog.Enabled=false;
					this.m_bUserCancel=false;
					this.m_oRunCore = new RunCoreOld(this);
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

		private void btnSaveAuditTables_Click(object sender, System.EventArgs e)
		{
			int x;
			string strInsertSQL;
			string strUpdateType="insert";
			System.Data.DataTable p_dt;
			this.btnCancel.Text = "Cancel";
			this.btnViewAuditTables.Enabled=false;
			try
			{
                p_dt = this.m_oRunCore.m_ado.m_DataSet.Tables[Tables.Audit.DefaultCondRxAuditTableName].GetChanges();
			

				if (p_dt.Rows.Count > 0)
				{
					this.progressBar1.Visible=true;
					this.progressBar1.Minimum = 1;
					this.progressBar1.Maximum = p_dt.Rows.Count;
					this.progressBar1.Value = 1;
					this.lblMsg.Visible=true;
					this.lblMsg.Text = "Saving cond_rx_audit records";
					this.lblMsg.Refresh();
                    if (this.m_oRunCore.m_ado.getRecordCount(this.m_oRunCore.m_TempMDBFileConn, "select count(*) from " + Tables.Audit.DefaultCondRxAuditTableName, 
                        Tables.Audit.DefaultCondRxAuditTableName) > 0)
						strUpdateType = "update";
				
				
				
					for (x=0; x<=p_dt.Rows.Count-1;x++)
					{
						this.progressBar1.Value = x+1;
						strInsertSQL = "INSERT INTO plot_cond_rx_audit " +
							"(biosum_cond_id,rx,fvs_ffe_yn,processor_tree_vol_val_yn, harvest_costs_yn) " + 
							" VALUES " + 
							"('" + p_dt.Rows[x]["biosum_cond_id"].ToString().Trim() + "'" + 
							",'" + p_dt.Rows[x]["rx"].ToString().Trim() + "'" + 
							",'" + p_dt.Rows[x]["fvs_ffe_yn"].ToString().Trim() + "'" + 
							",'" + p_dt.Rows[x]["processor_tree_vol_val_yn"].ToString().Trim() + "'" + 
							",'" + p_dt.Rows[x]["harvest_costs_yn"].ToString().Trim() + "');";
						if (strUpdateType.Trim()=="update")
						{
							this.m_strSQL = "SELECT COUNT(*) FROM plot_cond_rx_audit WHERE trim(biosum_cond_id)='" + p_dt.Rows[x]["biosum_cond_id"].ToString().Trim() + "'" + 
								" AND rx = '" + p_dt.Rows[x]["rx"].ToString().Trim() + "';";
							if (this.m_oRunCore.m_ado.getRecordCount(this.m_oRunCore.m_TempMDBFileConn,this.m_strSQL,"plot_cond_rx_audit") > 0)
							{
								this.m_strSQL ="UPDATE plot_cond_rx_audit SET fvs_ffe_yn = " + 
									"'" + p_dt.Rows[x]["fvs_ffe_yn"].ToString().Trim() + "'" + 
									", processor_tree_vol_val_yn = '" + p_dt.Rows[x]["fvs_ffe_yn"].ToString().Trim() + "'" + 
									", harvest_costs_yn = '" + p_dt.Rows[x]["harvest_costs_yn"].ToString().Trim() + "'" + 
									" WHERE trim(biosum_cond_id) = '" + p_dt.Rows[x]["biosum_cond_id"].ToString().Trim() + "'" + 
									"    AND rx = '" + p_dt.Rows[x]["rx"].ToString().Trim() + "';";
							}
							else
							{
								this.m_strSQL = strInsertSQL;
							}

						}
						else
						{
							this.m_strSQL = strInsertSQL;
						}
						this.m_oRunCore.m_ado.SqlNonQuery(this.m_oRunCore.m_TempMDBFileConn,this.m_strSQL);
						/**********************************
						 **check to see if user cancelled
						 **********************************/
						System.Windows.Forms.Application.DoEvents();
						if (this.btnCancel.Text.Trim().ToUpper()=="START")
							break;
					}
					/********************************************************
					 **if the did not complete do not accept changes
					 ********************************************************/
					if (x >=p_dt.Rows.Count)
                        this.m_oRunCore.m_ado.m_DataSet.Tables[Tables.Audit.DefaultCondRxAuditTableName].AcceptChanges();
				}
			}
			catch
			{

			}
			
			strUpdateType="insert";
			try
			{
				p_dt = this.m_oRunCore.m_ado.m_DataSet.Tables["plot_cond_audit"].GetChanges();
				if (p_dt.Rows.Count > 0)
				{
					this.progressBar1.Visible=true;
					this.progressBar1.Minimum = 1;
					this.progressBar1.Maximum = p_dt.Rows.Count;
					this.progressBar1.Value = 1;
					this.lblMsg.Visible=true;
					this.lblMsg.Text = "Saving plot_cond_audit records";
					this.lblMsg.Refresh();
					if (this.m_oRunCore.m_ado.getRecordCount(this.m_oRunCore.m_TempMDBFileConn,"select count(*) from plot_cond_audit","plot_cond_audit") > 0)
						strUpdateType = "update";
					for (x=0; x<=p_dt.Rows.Count-1;x++)
					{
						this.progressBar1.Value = x+1;
						strInsertSQL = "INSERT INTO plot_cond_audit " +
							"(biosum_cond_id,gis_travel_times_yn,fvs_ffe_yn,processor_tree_vol_val_yn, harvest_costs_yn) " + 
							" VALUES " + 
							"('" + p_dt.Rows[x]["biosum_cond_id"].ToString().Trim() + "'" + 
							",'" + p_dt.Rows[x]["gis_travel_times_yn"].ToString().Trim() + "'" + 
							",'" + p_dt.Rows[x]["fvs_ffe_yn"].ToString().Trim() + "'" + 
							",'" + p_dt.Rows[x]["processor_tree_vol_val_yn"].ToString().Trim() + "'" + 
							",'" + p_dt.Rows[x]["harvest_costs_yn"].ToString().Trim() + "');";
						if (strUpdateType.Trim()=="update")
						{
							this.m_strSQL = "SELECT COUNT(*) FROM plot_cond_audit WHERE trim(biosum_cond_id)='" + p_dt.Rows[x]["biosum_cond_id"].ToString().Trim() + "';";
							if (this.m_oRunCore.m_ado.getRecordCount(this.m_oRunCore.m_TempMDBFileConn,this.m_strSQL,"plot_cond_audit") > 0)
							{
								this.m_strSQL ="UPDATE plot_cond_audit SET fvs_ffe_yn = " + 
									"'" + p_dt.Rows[x]["fvs_ffe_yn"].ToString().Trim() + "'" + 
									", gis_travel_times_yn = '" + p_dt.Rows[x]["gis_travel_times_yn"].ToString().Trim() + "'" +
									", processor_tree_vol_val_yn = '" + p_dt.Rows[x]["fvs_ffe_yn"].ToString().Trim() + "'" + 
									", harvest_costs_yn = '" + p_dt.Rows[x]["harvest_costs_yn"].ToString().Trim() + "'" + 
									" WHERE trim(biosum_cond_id) = '" + p_dt.Rows[x]["biosum_cond_id"].ToString().Trim() + "';";
							}
							else
							{
								this.m_strSQL = strInsertSQL;
							}

						}
						this.m_oRunCore.m_ado.SqlNonQuery(this.m_oRunCore.m_TempMDBFileConn,this.m_strSQL);

						/**********************************
						 **check to see if user cancelled
						 **********************************/
						System.Windows.Forms.Application.DoEvents();
						if (this.btnCancel.Text.Trim().ToUpper()=="START") break;
					}
					/********************************************************
					 **if the did not complete do not accept changes
					 ********************************************************/
					if (x >=p_dt.Rows.Count)
						this.m_oRunCore.m_ado.m_DataSet.Tables["plot_cond_audit"].AcceptChanges();
				}
			}
			catch
			{
			}
			this.btnCancel.Text = "Start";
			this.btnViewAuditTables.Enabled=true;
			this.progressBar1.Visible=false;
			this.lblMsg.Visible=false;


		}

		/// <summary>
		/// start microsoft access application and open the 
		/// temporary mdb file that contains the links and 
		/// working tables used in the most recent run
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnAccess_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo.UseShellExecute = true;
			try
			{
				proc.StartInfo.FileName = this.m_oRunCore.m_strTempMDBFile;
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
				this.btnAccess.Left = this.btnViewLog.Left - this.btnAccess.Width;
				this.btnViewScenarioTables.Left = this.btnViewLog.Left - this.btnViewScenarioTables.Width;
				this.btnCancel.Left = (int)(this.progressBar1.Width * .50) - (int)(this.btnCancel.Width * .50);
			}
			catch
			{
			}
		}
	}

	/// <summary>
	/// main class used for running the core analysis scenario
	/// </summary>
	public class RunCoreOld
	{
		FIA_Biosum_Manager.frmRunOptimizerScenario m_frmRunCoreScenario;
		private int m_intError;
		private string m_strSQL;
		private string m_strConn;
		public string m_strTempMDBFile;
		public ado_data_access m_ado;
		public System.Data.OleDb.OleDbConnection m_TempMDBFileConn;
		private env m_oEnv;
		public string m_strPlotTable;
		public string m_strRxTable;
		public string m_strTravelTimeTable;
		public string m_strCondTable;
		public string m_strFFETable;
		public string m_strHvstCostsTable;
		public string m_strPSiteTable;
		public string m_strTreeVolValBySpcDiamGroupsTable;
		public string m_strTreeVolValSumTable = "tree_vol_val_sum_by_rx";
        public string m_strUserDefinedPlotSQL;
		private System.IO.FileStream m_txtFileStream;
		private System.IO.StreamWriter m_txtStreamWriter;
		private string m_strLine;
		
		

		public RunCoreOld(FIA_Biosum_Manager.frmRunOptimizerScenario p_form)
		{
			
			this.m_intError=0;
                                       
			this.m_frmRunCoreScenario = p_form;

			try
			{
				this.m_txtFileStream = new System.IO.FileStream(this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\runlog.txt", System.IO.FileMode.Create, 
					System.IO.FileAccess.Write);
				this.m_txtStreamWriter = new System.IO.StreamWriter(this.m_txtFileStream);
				this.m_txtStreamWriter.WriteLine(" ");
				this.m_txtStreamWriter.WriteLine(" ");
				this.m_strLine = "Core Analysis Run Log " + System.DateTime.Now.ToString();
				this.m_txtStreamWriter.WriteLine("{0}{1}","        ", this.m_strLine);
				this.m_txtStreamWriter.WriteLine(" ");
				this.m_txtStreamWriter.WriteLine("Project: {0}", ((frmMain)this.m_frmRunCoreScenario.m_frmScenario.ParentForm).frmProject.uc_project1.txtProjectId.Text);
				this.m_txtStreamWriter.WriteLine("Project Directory: {0}", ((frmMain)this.m_frmRunCoreScenario.m_frmScenario.ParentForm).frmProject.uc_project1.txtRootDirectory.Text);
                this.m_txtStreamWriter.WriteLine("Scenario Directory: {0}", this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioPath.Text);
				this.m_strLine = "---------------------------------------------------------------";
				this.m_txtStreamWriter.WriteLine(this.m_strLine);
			}
			catch (Exception caught)
			{
				MessageBox.Show(caught.ToString());
				this.m_intError=-1;
				return;
			}

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Create A Temporary MDB File With Links To All The Core Tables And Scenario Result Tables");
			this.m_txtStreamWriter.WriteLine("------------------------------------------------------------------------------------------------");


			/**************************************************************************
			 **first lets create a temp mdb with links to all the scenario core 
			 **and result tables
			 **************************************************************************/
			this.m_oEnv = new env();
			this.m_strTempMDBFile = 
				this.m_frmRunCoreScenario.m_frmScenario.uc_datasource1.CreateMDBAndScenarioTableDataSourceLinks(this.m_oEnv.strTempDir);  
			
			this.m_strUserDefinedPlotSQL= 
				this.m_frmRunCoreScenario.m_frmScenario.uc_scenario_filter1.txtCurrentSQL.Text.Trim();

            
			if (this.m_strTempMDBFile.Trim().Length == 0)
			{
				this.m_txtStreamWriter.WriteLine("!!RunCore: Error Creating MDB File Containing Links To All The Tables!!");
				MessageBox.Show("RunCore: Error Creating MDB File Containing Links To All The Core Tables");
			}
			else
			{
				this.m_txtStreamWriter.WriteLine("Delete Scenario Result Table Records");
				this.DeleteScenarioResultRecords();

				this.m_txtStreamWriter.WriteLine("Initialize Air Destruction Tables");
				InitializeAirDestructionTables();


				this.m_txtStreamWriter.WriteLine("Links MDB File: " + this.m_strTempMDBFile);
				this.m_frmRunCoreScenario.btnAccess.Enabled=false;

				getTableNames();				

				
				//CREATE TABLE LINKS
				CreateScenarioResultTableLinks();
				if (this.m_intError != 0) return;
				
				

				this.CreateScenarioTableLinks();
				if (this.m_intError !=0) return;

				//CREATE WORK TABLES
				CreateTableStructureOfHarvestCosts();
				if (this.m_intError != 0) return;

				CreateTableStructureForIntensity();
				if (this.m_intError !=0) return;

				this.CreateTableStructureForScenarioProcessingSites();
				if (this.m_intError !=0) return;

				this.CreateTableStructureForPSiteSum();
				if (this.m_intError !=0) return;

                this.CreateTableStructureForOwnershipSum();
                if (this.m_intError != 0) return;


				this.CreateTableStructureForHaulCosts();
				if (this.m_intError != 0) return;

				this.CreateTableStructureForPlotCondAccessiblity();
				if (this.m_intError != 0) return;

				
				this.m_ado = new ado_data_access();

				this.m_strConn=m_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");

			
				this.m_TempMDBFileConn = new System.Data.OleDb.OleDbConnection();
				this.m_ado.OpenConnection(this.m_strConn,ref this.m_TempMDBFileConn);
				if (this.m_ado.m_intError==0)
				{
					/******************************************************************
					 **create table structure of user defined sql plot filter statement
					 ******************************************************************/
                    CreateTableStructureOfUserDefinedSQL(this.m_strUserDefinedPlotSQL);

					/********************************************************************
					 **create table structure for condition table filters
					 ********************************************************************/
					CreateTableStructureForConditionTable();

					/********************************************************************
					 **filter scenario selected processing sites
					 ********************************************************************/
					FilterPSites();

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


					/***********************************************************
					 **identify the plots that are accessible
					 ***********************************************************/
					if (this.m_intError==0) // && this.m_frmRunCoreScenario.chkProcAccessible.Checked==true)
					{
						this.m_frmRunCoreScenario.lblProcAccessible.ForeColor=System.Drawing.Color.Green;
						this.m_frmRunCoreScenario.lblProcAccessible.Text = "Processing";
						this.m_frmRunCoreScenario.lblProcAccessible.Refresh();
						this.PlotAccessible();
				
					}
					else
					{
						this.m_frmRunCoreScenario.lblProcAccessible.Text = "NA";
						this.m_frmRunCoreScenario.lblProcAccessible.Refresh();
					}

					

					/**************************************************************
					 **get the fastest travel time from plot to processing site
					 **************************************************************/
					if (this.m_intError == 0 && this.m_frmRunCoreScenario.chkProcTravelTimes.Checked==true && this.m_frmRunCoreScenario.m_bUserCancel==false)
					{
						this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Green;
						this.m_frmRunCoreScenario.lblProcTravelTimes.Text="Processing";
						this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
						this.getHaulCosts();

					}
					else
					{
						if (this.m_frmRunCoreScenario.m_bUserCancel==false)
						{
							this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "NA";
							this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
						}
					}
                    
					/***************************************************************************
					 **sum up tree volumes and values by plot+condition, treatment and species
					 ***************************************************************************/
					if (this.m_intError == 0 && this.m_frmRunCoreScenario.chkProcSumTree.Checked==true && this.m_frmRunCoreScenario.m_bUserCancel==false)
					{
						this.m_frmRunCoreScenario.lblProcSumTree.ForeColor = System.Drawing.Color.Green;
						this.m_frmRunCoreScenario.lblProcSumTree.Text="Processing";
						this.m_frmRunCoreScenario.lblProcSumTree.Refresh();
						this.sumTreeVolVal();

					}
					else
					{
						if (this.m_frmRunCoreScenario.m_bUserCancel == false)
						{
							this.m_frmRunCoreScenario.lblProcSumTree.Text = "NA";
							this.m_frmRunCoreScenario.lblProcSumTree.Refresh();
						}
					}
					/***************************************************************************
					 **valid combos
					 ***************************************************************************/
					if (this.m_intError == 0 && this.m_frmRunCoreScenario.m_bUserCancel==false)
					{
						this.m_frmRunCoreScenario.lblProcValidCombos.ForeColor = System.Drawing.Color.Green;
						this.m_frmRunCoreScenario.lblProcValidCombos.Text="Processing";
						this.m_frmRunCoreScenario.lblProcValidCombos.Refresh();
						this.validcombos();

					}
					if (this.m_intError==0 && this.m_frmRunCoreScenario.m_bUserCancel==false)
					{
						this.m_frmRunCoreScenario.lblProcEffective.ForeColor = System.Drawing.Color.Green;
						this.m_frmRunCoreScenario.lblProcEffective.Text="Processing";
						this.m_frmRunCoreScenario.lblProcEffective.Refresh();
						this.effective();

					}
					/***************************************************************
					 **wood product yields net revenue and costs summary table
					 ***************************************************************/
					if (this.m_intError==0 && this.m_frmRunCoreScenario.m_bUserCancel==false)
					{
						this.m_frmRunCoreScenario.lblSumWoodProducts.ForeColor = System.Drawing.Color.Green;
						this.m_frmRunCoreScenario.lblSumWoodProducts.Text="Processing";
						this.m_frmRunCoreScenario.lblSumWoodProducts.Refresh();
						this.product_yields_net_rev_costs_summary();

					}
					/*********************************************************************
					 **find the best treatments for revenue, torch/crown index improvement,
					 **and merch removal
					 *********************************************************************/ 
					if (this.m_intError==0 && this.m_frmRunCoreScenario.m_bUserCancel==false)
					{
						this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Green;
						this.m_frmRunCoreScenario.lblProcBestRx.Text="Processing";
						this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
						this.best_rx_summary();
					}
					/*******************************************************************************
					 **expand acreage for best treatments by plot 
					 *******************************************************************************/
					if (this.m_intError==0 && this.m_frmRunCoreScenario.m_bUserCancel==false)
					{
						this.m_frmRunCoreScenario.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Green;
						this.m_frmRunCoreScenario.lblProcBestRxPlot.Text="Processing";
						this.m_frmRunCoreScenario.lblProcBestRxPlot.Refresh();
						this.BestTreatmentsByPlot();
					}
					/**********************************************************************
					 **sum up the values by processing site
					 **********************************************************************/
					if (this.m_intError == 0 && this.m_frmRunCoreScenario.m_bUserCancel==false)
					{
						this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Green;
						this.m_frmRunCoreScenario.lblProcBestRxPSite.Text="Processing";
						this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
						this.SumPSite();

					}

					/**********************************************************************
					 **sum up the values by ownership
					 **********************************************************************/
					if (this.m_intError == 0 && this.m_frmRunCoreScenario.m_bUserCancel==false)
					{
						this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Green;
						this.m_frmRunCoreScenario.lblProcBestRxOwner.Text="Processing";
						this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
						this.SumOwnership();

					}
					if (this.m_intError==0 && this.m_frmRunCoreScenario.m_bUserCancel==false)
					{
						this.CreateHtml();
					}
					this.m_strLine = "***End Of Core Analysis Run: " + System.DateTime.Now.ToString() + " ***"; 
					this.m_txtStreamWriter.WriteLine(this.m_strLine);

					this.m_frmRunCoreScenario.progressBar1.Visible=false;
					this.m_frmRunCoreScenario.lblMsg.Visible=false;
					this.m_frmRunCoreScenario.btnCancel.Text = "Start";
					this.m_frmRunCoreScenario.btnViewScenarioTables.Enabled=true;
					
                    if (this.m_intError == 0) this.m_frmRunCoreScenario.btnAccess.Enabled=true;
					
					

					if (this.m_intError != 0)
					{
                       MessageBox.Show("Completed With Errors");
					}
					else
					{
					   MessageBox.Show("Successfully Completed");
					}
				}
			}
			this.m_txtStreamWriter.Close();
			this.m_txtFileStream.Close();
			this.m_frmRunCoreScenario.btnViewLog.Enabled=true;
			if (this.m_frmRunCoreScenario.chkAuditTables.Enabled==true)
				this.m_frmRunCoreScenario.btnViewAuditTables.Enabled=true;
			



		}
		public RunCoreOld()
		{

		}
		~RunCoreOld()
		{
			this.m_txtFileStream.Close();
			this.m_txtStreamWriter.Close();
			this.m_txtFileStream=null;
			this.m_txtStreamWriter=null;
		}
		public FIA_Biosum_Manager.frmRunOptimizerScenario ReferenceRunCoreScenarioForm 
		{
			set {this.m_frmRunCoreScenario=value;}
			get {return m_frmRunCoreScenario;}
		}

		/// <summary>
		/// create a table structure that will hold
		/// the plot data that results when running the user 
		/// defined sql
		/// </summary>
		/// <param name="strUserDefinedSQL"></param>
		private void CreateTableStructureOfUserDefinedSQL(string strUserDefinedSQL)
		{
			
			
			/*******************************************************************
			 ** get scenario_results.mdb path
			 *******************************************************************/
			string strMDBPathAndFile = this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";
			/********************************************************
			 **get the user defined PLOT filter sql
			 ********************************************************/
			this.m_strSQL = this.m_strUserDefinedPlotSQL;
			/****************************************************************
			 **get the table structure that results from executing the sql
			 ****************************************************************/
			System.Data.DataTable p_dt = this.m_ado.getTableSchema(this.m_TempMDBFileConn,this.m_strSQL);

			/*****************************************************************
			 **create the table structure in the scenario_results.mdb file
			 **and give it the name of userdefinedplotfilter
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine("Create userdefinedplotfilter Table Schema From User Defined Plot Filter SQL");
			dao_data_access p_dao = new dao_data_access();
			p_dao.CreateMDBTableFromDataSetTable(this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioPath.Text + "\\db\\scenario_results.mdb","userdefinedplotfilter",p_dt,true);
			p_dt.Dispose();
			this.m_ado.m_OleDbDataReader.Close();
			if (p_dao.m_intError!=0)
			{
				this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
				this.m_intError=p_dao.m_intError;
				p_dao=null;
				return;
			}
			/***********************************************************************
			 **create a table link to the user defined plot filter table in the
			 **temporary MDB file located on the hard drive of the user
			 ***********************************************************************/
			this.m_txtStreamWriter.WriteLine("Create link in " + this.m_strTempMDBFile);
			p_dao.CreateTableLink(this.m_strTempMDBFile,"userdefinedplotfilter",strMDBPathAndFile,"userdefinedplotfilter",true);
			this.m_txtStreamWriter.WriteLine("{0}\t{1}",strMDBPathAndFile,"userdefinedplotfilter");

			/***********************************************************************
			 **make a copy of the userdefinedplot filter table and give it the
			 **name ruledefinitionsplotfilter. This will apply the owngrpcd
			 **filters and any other future filters to the userdefinedplotfilter table.
			 ***********************************************************************/
			this.m_txtStreamWriter.WriteLine("Delete table ruledefinitionsplotfilter");
			//first delete the table if it exists already
			p_dao.DeleteTableFromMDB(this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioPath.Text + "\\db\\scenario_results.mdb","ruledefinitionsplotfilter");
			if (p_dao.m_intError !=0)
			{
				this.m_txtStreamWriter.WriteLine("!! Error Deleting alluserdefinedplotfilter Table!!");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
			this.m_txtStreamWriter.WriteLine("Copy table structure userdefinedplotfilter to ruledefinitionsplotfilter");
			p_dao.MoveTableToMDB(this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioPath.Text + "\\db\\scenario_results.mdb","ruledefinitionsplotfilter",this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioPath.Text + "\\db\\scenario_results.mdb","userdefinedplotfilter",false);
			if (p_dao.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!! Error Creating ruledefinitionsplotfilter Table!!");
				this.m_intError = p_dao.m_intError;
				p_dao=null;
				return;
			}
			/***********************************************************************
			 **create a table link to the ruledefinitionsplotfilter table in the
			 **temporary MDB file located on the hard drive of the user
			 ***********************************************************************/
			p_dao.CreateTableLink(this.m_strTempMDBFile,"ruledefinitionsplotfilter",strMDBPathAndFile,"ruledefinitionsplotfilter",true);
			this.m_txtStreamWriter.WriteLine("Create link in " + this.m_strTempMDBFile);
			this.m_txtStreamWriter.WriteLine("{0}\t{1}",strMDBPathAndFile,"ruledefinitionsplotfilter");


			p_dao=null;




		}
		private void InitializeAirDestructionTables()
		{
			
			
			string strMDBPathAndFile = this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";
			ado_data_access oAdo = new ado_data_access();
			string strConn=oAdo.getMDBConnString(strMDBPathAndFile,"admin","");
			oAdo.OpenConnection(strConn);

			if (oAdo.TableExist(oAdo.m_OleDbConnection,"best_rx_summary_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE best_rx_summary_air_dest");

			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,
				  "SELECT * INTO best_rx_summary_air_dest FROM best_rx_summary");

			//MAX CROWN INDEX IMPROVEMENT BY PLOT
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_plots_air_dest"))
		          oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE max_ci_imp_plots_air_dest");

			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,
					"SELECT * INTO max_ci_imp_plots_air_dest FROM max_ci_imp_plots");

			//MAX CROWN INDEX AND POSITIVE NET REVENUE IMPROVEMENT BY PLOT
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_pnr_plots_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE max_ci_imp_pnr_plots_air_dest");


			//MAX CROWN INDEX AND POSITIVE NET REVENUE IMPROVEMENT SUMMED BY OWNERSHIP GROUP
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_pnr_sum_own_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE max_ci_imp_pnr_sum_own_air_dest");


			
			//MAX CROWN INDEX IMPROVEMENT SUMMED BY OWNERSHIP GROUP
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ci_imp_sum_own_air_dest"))
					oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE max_ci_imp_sum_own_air_dest");

			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,
				"SELECT * INTO max_ci_imp_sum_own_air_dest FROM max_ci_imp_sum_own");


			//MAX TORCH INDEX IMPROVEMENT BY PLOT
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_plots_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE max_ti_imp_plots_air_dest");

			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,
				"SELECT * INTO max_ti_imp_plots_air_dest FROM max_ti_imp_plots");

			//MAX TORCH INDEX AND POSITIVE NET REVENUE IMPROVEMENT BY PLOT
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_pnr_plots_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE max_ti_imp_pnr_plots_air_dest");


			//MAX TORCH INDEX AND POSITIVE NET REVENUE SUMMED BY OWNERSHIP GROUP
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_pnr_sum_own_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE max_ti_imp_pnr_sum_own_air_dest");


			//MAX TORCH INDEX IMPROVEMENT SUMMED BY OWNERSHIP GROUP
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_ti_imp_sum_own_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE max_ti_imp_sum_own_air_dest");

			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,
				"SELECT * INTO max_ti_imp_sum_own_air_dest FROM max_ti_imp_sum_own");
			
			//MAX NET REVENUE BY PLOT
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_nr_plots_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE max_nr_plots_air_dest");


			//MAX NET REVENUE SUMMED BY OWNERSHIP GROUP
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_nr_sum_own_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE max_nr_sum_own_air_dest");


			//MAX POSITIVE NET REVENUE BY PLOT
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_pnr_plots_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE max_pnr_plots_air_dest");


			//MAX POSITIVE NET REVENUE SUMMED BY OWNERSHIP GROUP
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"max_pnr_sum_own_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE max_pnr_sum_own_air_dest");


			//MIN MERCHANTABLE WOOD BY PLOTS
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_plots_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE min_merch_plots_air_dest");

			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,
				"SELECT * INTO min_merch_plots_air_dest FROM min_merch_plots");

			//MIN MERCHANTABLE WOOD BY OWNERSHIP GROUP
			if (oAdo.TableExist(oAdo.m_OleDbConnection,"min_merch_sum_own_air_dest"))
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE min_merch_sum_own_air_dest");

			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,
				"SELECT * INTO min_merch_sum_own_air_dest FROM min_merch_sum_own");





			


			
			
		}

		/// <summary>
		/// create links to the tables located in the scenario_results.mdb file
		/// </summary>
		private void CreateScenarioResultTableLinks()
		{
			

			string[] strTableNames;
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();

		
			string strMDBPathAndFile = this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";


			int intCount = p_dao.getTableNames(strMDBPathAndFile,ref strTableNames);
			if (p_dao.m_intError==0)
			{
				if (intCount > 0)
				{
					for (int x=0; x <= intCount-1;x++)
					{
						this.m_txtStreamWriter.WriteLine("{0}\t{1}",strMDBPathAndFile,strTableNames[x]);

						p_dao.CreateTableLink(this.m_strTempMDBFile,strTableNames[x],strMDBPathAndFile,strTableNames[x]);
						if (p_dao.m_intError != 0)
						{
							this.m_txtStreamWriter.WriteLine("!!Error Creating Table Link!!!");
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
		/// create links to the audit tables
		/// </summary>
		private void CreateAuditTableLinks()
		{
			

			string[] strTableNames;
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();
			string strConn="";
			
			string strMDBPathAndFile = ((frmMain)this.m_frmRunCoreScenario.ParentForm).frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\audit.mdb";

			if (strMDBPathAndFile.Substring(strMDBPathAndFile.Trim().Length - 6,6).ToUpper()==".ACCDB")
				strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";
			else
			    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";

			
			int intCount = p_dao.getTableNames(strMDBPathAndFile,ref strTableNames);
			if (p_dao.m_intError==0)
			{
				if (intCount > 0)
				{
					for (int x=0; x <= intCount-1;x++)
					{
						this.m_txtStreamWriter.WriteLine("{0}\t{1}",strMDBPathAndFile,strTableNames[x]);
						p_dao.CreateTableLink(this.m_strTempMDBFile,strTableNames[x],strMDBPathAndFile,strTableNames[x]);
						if (p_dao.m_intError != 0)
						{
							this.m_txtStreamWriter.WriteLine("!!Error Creating Table Link!!!");
							this.m_intError = p_dao.m_intError;
							break;
						}
					}

				}
			}
			else
			{
				this.m_intError = p_dao.m_intError;
			}
			p_dao = null;
			
		}
			
		/// <summary>
		/// create links to the scenario tables that contain the 
		/// rule definitions defined by the user
		/// </summary>
		private void CreateScenarioTableLinks()
		{
			

			string[] strTableNames;
			strTableNames = new string[1];
			dao_data_access p_dao = new dao_data_access();
			string strConn="";
			
			string strMDBPathAndFile = ((frmMain)this.m_frmRunCoreScenario.ParentForm).frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\core\\db\\scenario_core_rule_definitions.mdb";

			if (strMDBPathAndFile.Substring(strMDBPathAndFile.Trim().Length - 6,6).ToUpper()==".ACCDB")
				strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + 
					strMDBPathAndFile + ";User Id=admin;Password=;";
			else
				strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strMDBPathAndFile + ";User Id=admin;Password=;";

			
			int intCount = p_dao.getTableNames(strMDBPathAndFile,ref strTableNames);
			if (p_dao.m_intError==0)
			{
				if (intCount > 0)
				{
					for (int x=0; x <= intCount-1;x++)
					{
						this.m_txtStreamWriter.WriteLine("{0}\t{1}",strMDBPathAndFile,strTableNames[x]);
						p_dao.CreateTableLink(this.m_strTempMDBFile,strTableNames[x],strMDBPathAndFile,strTableNames[x]);
						if (p_dao.m_intError != 0)
						{
							this.m_txtStreamWriter.WriteLine("!!Error Creating Table Link!!!");
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
		/// get the names of the core tables
		/// </summary>
		private void getTableNames()
		{
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Get Core Table Names");
			this.m_txtStreamWriter.WriteLine("---------------------");
			/**************************************************************
			 **get the plot table name
			 **************************************************************/
			this.m_strPlotTable=this.m_frmRunCoreScenario.m_frmScenario.uc_datasource1.getDataSourceTableName("PLOT");
			this.m_txtStreamWriter.WriteLine("Plot:{0}",this.m_strPlotTable);


			/**************************************************************
			 **get the treatment prescriptions table
			 **************************************************************/
			this.m_strRxTable = 
				this.m_frmRunCoreScenario.m_frmScenario.uc_datasource1.getDataSourceTableName("TREATMENT PRESCRIPTIONS");
			this.m_txtStreamWriter.WriteLine("Treatment:{0}",this.m_strRxTable);

			/**************************************************************
			 **get the travel time table name
			 **************************************************************/
			this.m_strTravelTimeTable = 
				this.m_frmRunCoreScenario.m_frmScenario.uc_datasource1.getDataSourceTableName("TRAVEL TIMES");
			this.m_txtStreamWriter.WriteLine("Travel Time:{0}",this.m_strTravelTimeTable);
			/**************************************************************
			 **get the cond table name
			 **************************************************************/
			this.m_strCondTable=this.m_frmRunCoreScenario.m_frmScenario.uc_datasource1.getDataSourceTableName("CONDITION");
			this.m_txtStreamWriter.WriteLine("Condition:{0}",this.m_strCondTable);


			/*************************************************************
			 **get the harvest costs table
			 *************************************************************/
			this.m_strHvstCostsTable=this.m_frmRunCoreScenario.m_frmScenario.uc_datasource1.getDataSourceTableName("HARVEST COSTS");
			this.m_txtStreamWriter.WriteLine("Harvest Costs:{0}",this.m_strHvstCostsTable);

			/**************************************************************
			 **get the processing site table name
			 **************************************************************/

			this.m_strPSiteTable = "scenario_psites_work_table";
			this.m_txtStreamWriter.WriteLine("Processsing Sites:{0}",this.m_strPSiteTable);

			this.m_strTreeVolValBySpcDiamGroupsTable = 
				this.m_frmRunCoreScenario.m_frmScenario.uc_datasource1.getDataSourceTableName("TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS");
			this.m_txtStreamWriter.WriteLine("Tree Volumes And Values By Species And Diameter groups:{0}",this.m_strTreeVolValBySpcDiamGroupsTable);


			this.m_strTreeVolValSumTable = "tree_vol_val_sum_by_rx";
			this.m_txtStreamWriter.WriteLine("Tree Sum Volume And Value:{0}",this.m_strTreeVolValSumTable);

		}
		
		private void FilterPSites()
		{
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Processing Site Filter");
			this.m_txtStreamWriter.WriteLine("-------------------");

			this.m_strSQL="DELETE FROM scenario_psites_work_table";
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_strSQL = "INSERT INTO scenario_psites_work_table (psite_id,trancd,biocd) " + 
				            "SELECT psite_id,trancd,biocd " + 
				            "FROM scenario_psites " + 
				            "WHERE TRIM(scenario_id)='" + this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioId.Text.Trim() + "' AND " + 
				                   "selected_yn='Y';";
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


		}
		/// <summary>
		/// set the plot_accessible_yn by evaluating the field values in
		/// gis_protected_area_yn, gis_moved_from_protected_to_unprotected_yn,
		/// and gis_roadless_yn
		/// </summary>
		private void PlotAccessible()
		{
			/*********************************************************************
			 **set the plot_accessible_yn flag to 'Y' or 'N' based on whether
			 **the plot is in a protected area, roadless area,
			 **has been moded from protected to an unprotected area
			 **when moving the plot to the nearest road, or 
			 **every condition on the plot is unavailable
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Plot and Condition Accessibility");
			this.m_txtStreamWriter.WriteLine("-------------------");


			this.m_frmRunCoreScenario.lblProcAccessible.ForeColor = System.Drawing.Color.Green;
			this.m_frmRunCoreScenario.lblProcAccessible.Text = "Processing";

			

			/********************************************************************
             **update cond_too_far_steep field to Y if slope is 
			 **<= 40% and the yarding distance >= to maximum yarding distance 
			 **for a slope <= 40%
			 ********************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
				            "INNER JOIN " + this.m_strPlotTable + " p " +
				            "ON c.biosum_plot_id = p.biosum_plot_id  " + 
				            "SET c.cond_too_far_steep_yn = 'Y' " + 
				            "WHERE c.slope IS NOT NULL AND " + 
				                  "c.slope <= 40 AND " + 
				                  "p.gis_yard_dist_ft >= " + this.m_frmRunCoreScenario.m_frmScenario.uc_scenario_filter1.strNonSteepYardingDistance.Trim() + ";" ;

			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/********************************************************************
			 **update cond_too_far_steep field to Y if slope is 
			 **> 40% and the yarding distance >= to maximum yarding distance 
			 **for a slope > 40%
			 ********************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
				"INNER JOIN " + this.m_strPlotTable + " p " +
				"ON c.biosum_plot_id = p.biosum_plot_id  " + 
				"SET c.cond_too_far_steep_yn = 'Y' " + 
				"WHERE c.slope IS NOT NULL AND " + 
				"c.slope > 40 AND " + 
				"p.gis_yard_dist_ft >= " + this.m_frmRunCoreScenario.m_frmScenario.uc_scenario_filter1.strSteepYardingDistance.Trim() + ";" ;

			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);



			/*************************************************************
			 **set the remainder of the cond_too_far_steep_yn fields to N
			 *************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
				"SET c.cond_too_far_steep_yn = 'N' " + 
				"WHERE c.cond_too_far_steep_yn IS NULL OR (c.cond_too_far_steep_yn <> 'Y' AND " + 
				                                          "c.cond_too_far_steep_yn <> 'N');";
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/*************************************************************
			 **update the condition accessible flag
			 *************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strCondTable +  " SET cond_accessible_yn = " + 
				"IIF(cond_too_far_steep_yn='Y','N','Y');";



			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/************************************************************************
			 **set the all_cond_not_accessible_yn flag to Y if every condition on the
			 **plot is inaccessible
			 ************************************************************************/

			/********************************************************************
			 **get the current condition record counts for each plot
			 ********************************************************************/
			//insert the condition counts into the work table
			this.m_strSQL = "INSERT INTO plot_cond_accessible_work_table (biosum_plot_id, num_cond) " + 
				" SELECT biosum_plot_id , COUNT(biosum_plot_id) " + 
				" FROM " + this.m_strCondTable + 
				" GROUP BY biosum_plot_id;";
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_strSQL = "INSERT INTO plot_cond_accessible_work_table2 (biosum_plot_id, num_cond, num_cond_not_accessible) " + 
				"SELECT a.biosum_plot_id, a.num_cond,b.cond_not_accessible_count as num_cond_not_accessible FROM plot_cond_accessible_work_table a," + 
				    "(SELECT biosum_plot_id, Count(*) AS cond_not_accessible_count FROM " + this.m_strCondTable + " " + 
				     "WHERE cond_accessible_yn='N' GROUP BY biosum_plot_id) b " + 
				"WHERE a.biosum_plot_id = b.biosum_plot_id;";
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/**********************************************************************
			 **initialize the all_cond_not_accessible_yn flag to N
			 **********************************************************************/
            this.m_strSQL = "UPDATE plot SET all_cond_not_accessible_yn = 'N';";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/**********************************************************************
			 **update the all_cond_not_accessible_yn flag if all conditions 
			 **are not accessible
			 **********************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable + " p " + 
				            "INNER JOIN plot_cond_accessible_work_table2 c " + 
				            "ON p.biosum_plot_id=c.biosum_plot_id " + 
				            "SET p.all_cond_not_accessible_yn=" + 
				              "IIF(c.num_cond=c.num_cond_not_accessible,'Y','N');";
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/*********************************************************************
			 **update the plot accessible flag
			 *********************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable +  " SET plot_accessible_yn = " + 
				"IIF(gis_protected_area_yn='Y' " + 
				" OR gis_roadless_yn='Y' " + 
				" OR all_cond_not_accessible_yn='Y','N','Y');";

			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


            


			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcAccessible) == true) return;
			if (this.m_ado.m_intError == 0)
			{
				this.m_frmRunCoreScenario.lblProcAccessible.ForeColor = System.Drawing.Color.Blue;
				this.m_frmRunCoreScenario.lblProcAccessible.Text = "Completed";

			}
			else
			{
				this.m_txtStreamWriter.WriteLine("!!Error Executing SQL!!");
				this.m_frmRunCoreScenario.lblProcAccessible.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcAccessible.Text = "!!Error!!";
				this.m_intError = this.m_ado.m_intError;

			}
			this.m_frmRunCoreScenario.lblProcAccessible.Refresh();

		}
		/// <summary>
		/// populate the haul_costs table and plot table with 
		/// the cheapest route for hauling merch and chip
		/// </summary>
		private void getHaulCosts()
		{
			
			//int x;
			string strTruckHaulCost="";
			string strRailHaulCost="";
			string strTransferMerchCost="";
			string strTransferBioCost="";

			
			this.m_frmRunCoreScenario.lblMsg.Text="Processing Haul Costs";
			this.m_frmRunCoreScenario.lblMsg.Visible=true;
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_frmRunCoreScenario.progressBar1.Minimum=0;
			this.m_frmRunCoreScenario.progressBar1.Maximum = 27;
			this.m_frmRunCoreScenario.progressBar1.Value=0;
			this.m_frmRunCoreScenario.progressBar1.Visible=true;

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update Plot And Haul Cost Tables With Merch And Chip Haul Costs");
			this.m_txtStreamWriter.WriteLine("-------------------------------------------------------------");

			/********************************************
			 **get the haul cost per green ton per hour
			 ********************************************/

            
			strTruckHaulCost = strTruckHaulCost.Replace(",","");
			
			if (strTruckHaulCost.Trim().Length == 1) strTruckHaulCost = "0.00";
			

			strRailHaulCost = strRailHaulCost.Replace(",","");
			if (strRailHaulCost.Trim().Length == 1) strRailHaulCost = "0.00";
			
			/***********************************************
			 **get the transfer cost per green to per hour
			 ***********************************************/
            strTransferMerchCost = this.m_frmRunCoreScenario.m_frmScenario.uc_scenario_costs1.RailMerchTransferCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTransferMerchCost = strTransferMerchCost.Replace(",","");
			if (strTransferMerchCost.Trim().Length == 1) strTransferMerchCost = "0.00";
			
            strTransferBioCost = this.m_frmRunCoreScenario.m_frmScenario.uc_scenario_costs1.RailChipTransferCostDollarsPerGreenTonPerHour.Replace("$", "").ToString();
			strTransferBioCost = strTransferBioCost.Replace(",","");
			if (strTransferBioCost.Trim().Length == 1) strTransferBioCost = "0.00";
			 



			/*******************************************************************************
			 **zap the haul_costs table
			 *******************************************************************************/
			this.m_strSQL = "DELETE FROM haul_costs;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("delete records in haul_costs table");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 1;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     
			
			
			this.m_frmRunCoreScenario.lblMsg.Text="Null The Plot Table's Haul Cost Fields";
			this.m_frmRunCoreScenario.lblMsg.Refresh();

			/***************************************************************
			 **null the plots merch_haul_cost_psite, merch_haul_cost_id and 
			 **merch_haul_cpa_pt,chip_haul_cost_psite,chip_haul_cost_id 
			 **and chip_haul_cpa_pt  fields
			 ***************************************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable + " " + 
				"SET merch_haul_cost_id = null," + 
				"merch_haul_cost_psite = null, " +
				"merch_haul_cpa_pt = null ," + 
				"chip_haul_cost_id = null," + 
				"chip_haul_cost_psite = null," + 
				"chip_haul_cpa_pt = null;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("null the plot table's haul cost fields");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 3;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     
			
			

			/*****************************************************************
			 **delete any records that may exist in the work tables
			 *****************************************************************/

			this.m_frmRunCoreScenario.lblMsg.Text="Delete Records In Work Tables";
			this.m_frmRunCoreScenario.lblMsg.Refresh();

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
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 4;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     

			//MERCH AND CHIP ROAD PROCESSING SITE HAUL COSTS
			this.m_frmRunCoreScenario.lblMsg.Text="Road Haul Costs For Merchantable Wood Processing Sites";
			this.m_frmRunCoreScenario.lblMsg.Refresh();

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Merchantable wood processing site haul costs");
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
				                   "0 AS transfer_cost, s.psite_id," +
                                   "(" + strTruckHaulCost.Trim() + " * t.travel_time) AS road_cost_dpgt," +
                                   "0 AS rail_cost_dpgt, (transfer_cost+road_cost_dpgt+rail_cost_dpgt) AS complete_haul_cost_dpgt," + 
				                   "'M' as materialcd " +
                               "FROM " + this.m_strTravelTimeTable + " t," + 
				                         this.m_strPSiteTable + " s " + 
                               "WHERE t.psite_id=s.psite_id AND " + 
				                     "(s.trancd=1 OR s.trancd =3) AND " + 
				                     "(s.biocd=3 OR s.biocd=1);";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table all travel time records where psite has road access and processes merch");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 5;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
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
				                        "0 AS transfer_cost, a.road_cost, 0 AS rail_cost," + 
				                        "b.min_cost AS total_haul_cost," + 
                                        "'M' AS materialcd " + 
                                 "FROM  all_road_merch_haul_costs_work_table  a," + 
                                    "(SELECT biosum_plot_id,MIN(total_haul_cost) AS min_cost " + 
                                     "FROM all_road_merch_haul_costs_work_table " + 
									 "GROUP BY biosum_plot_id)  b," + 
                                    "(SELECT biosum_plot_id,  psite_id ," + 
											"MIN(total_haul_cost) AS min_cost2 " + 
                                     "FROM all_road_merch_haul_costs_work_table " + 
									 "GROUP BY biosum_plot_id, psite_id) c " + 
								  "WHERE  c.biosum_plot_id = b.biosum_plot_id AND " + 
										 "a.biosum_plot_id = b.biosum_plot_id AND " + 
										 "a.psite_id = c.psite_id AND " + 
										 "b.min_cost = c.min_cost2;";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table. Find the cheapest plot to merch processing site road route.");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 5;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     
			this.m_frmRunCoreScenario.lblMsg.Text="Road Haul Costs For Chip Wood Processing Sites";
			this.m_frmRunCoreScenario.lblMsg.Refresh();

            /***********************************************************************
			 **Append to a table all travel time records where the psite 
			 **has road only or road/rail access and processes 
			 **chip only or merch/chip
			 ***********************************************************************/
            this.m_strSQL = "INSERT INTO all_road_chip_haul_costs_work_table " + 
				               "SELECT t.biosum_plot_id, 0 AS railhead_id," + 
				                      "0 AS transfer_cost,s.psite_id," + 
				                      "(" +  strTruckHaulCost.Trim() + " * t.travel_time) AS road_cost," + 
				                      "0 AS rail_cost, " + 
				                      "(transfer_cost+road_cost+rail_cost) AS total_haul_cost," + 
				                      "'C' AS materialcd " + 
                              "FROM " + this.m_strTravelTimeTable + " t ," + 
				                        this.m_strPSiteTable + " s " + 
				              "WHERE t.psite_id=s.psite_id AND " + 
				                   "(s.trancd=1 OR s.trancd=3) AND " + 
				                   "(s.biocd=3 OR s.biocd=2);";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table all travel time records where psite has road access and processes chips.");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 7;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     


            /******************************************************************
			 **For each plot get the cheapest road route to a psite. 
			 ******************************************************************/
            this.m_strSQL = "INSERT INTO cheapest_road_chip_haul_costs_work_table " + 
				                "SELECT b.biosum_plot_id, c.psite_id, null AS railhead_id," + 
				                        "0 AS transfer_cost,a.road_cost," + 
				                        "0 AS rail_cost, b.min_cost AS total_haul_cost," + 
				                        "'C' AS materialcd " + 
                                "FROM all_road_chip_haul_costs_work_table  a," + 
                                       "(SELECT biosum_plot_id,MIN(total_haul_cost) AS min_cost " + 
										"FROM all_road_chip_haul_costs_work_table " + 
                                        "GROUP BY biosum_plot_id)  b," + 
									   "(SELECT biosum_plot_id,  psite_id ," + 
											    "MIN(total_haul_cost) AS min_cost2 " + 
										"FROM all_road_chip_haul_costs_work_table " + 
										"GROUP BY biosum_plot_id, psite_id)  c " + 
							    "WHERE c.biosum_plot_id=b.biosum_plot_id AND " + 
									  "a.biosum_plot_id=b.biosum_plot_id AND " + 
									  "a.psite_id=c.psite_id AND b.min_cost=c.min_cost2;";



			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table. Find the cheapest plot to chip processing site road route.");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 8;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     

			

			//MERCH AND CHIP RAIL PROCESSING SITE HAUL COSTS
			/*********************************************************
			 **Append to a table all travel time collector_id (psite)
			 **records where the psite has rail access
			 *********************************************************/
			this.m_frmRunCoreScenario.lblMsg.Text="Rail Haul Costs For Merchantable Wood Processing Sites";
			this.m_frmRunCoreScenario.lblMsg.Refresh();

			this.m_strSQL = "INSERT INTO merch_rh_to_collector_haul_costs_work_table " + 
				               "SELECT t.psite_id AS railhead_id," + 
				                      "t.collector_id AS psite_id," + 
				                      "(" + strTransferMerchCost.Trim() + " * t.travel_time)  AS transfer_cost," + 
				                      "0 AS road_cost," + 
				                      "((t.travel_time * 45) * " + strRailHaulCost.Trim() + ") AS rail_cost," + 
				                      "0 AS total_haul_cost,  'M' AS materialcd " + 
				               "FROM " + this.m_strTravelTimeTable + " t " + 
				               "INNER JOIN  " + this.m_strPSiteTable + " s " + 
				               "ON t.collector_id = s.psite_id " + 
				               "WHERE  s.trancd=3 And (s.biocd=3 Or s.biocd=1) AND " + 
							   "EXISTS (SELECT ss.psite_id " + 
				                       "FROM " + this.m_strPSiteTable + " ss " + 
				                       "WHERE t.psite_id=ss.psite_id AND ss.trancd=2 AND (ss.biocd=3 Or ss.biocd=1));";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table all travel time collector_id (psite) records where the psite has rail access and processes merch.");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 9;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
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
				                       "r.transfer_cost," + 
				                       "(" + strTruckHaulCost.Trim() + " * t.travel_time) AS road_cost," + 
				                       "r.rail_cost, (r.transfer_cost + road_cost + r.rail_cost) AS total_haul_cost," +
				                       "'M' AS materialcd " + 
				               "FROM  " + this.m_strTravelTimeTable + " t," + 
				                      "merch_rh_to_collector_haul_costs_work_table r " + 
				               "WHERE r.railhead_id = t.psite_id;";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table travel time plot records and previous work rail/merch table results");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 10;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     


			this.m_strSQL = "UPDATE merch_plot_to_rh_to_collector_haul_costs_work_table " + 
				"SET total_haul_cost = transfer_cost + road_cost + rail_cost;";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("update merch by road and rail total haul cost ");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 11;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
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
				                      "a.transfer_cost, a.road_cost,a.rail_cost," + 
				                      "c.min_cost AS total_haul_cost,'M' as materialcd " +
				               "FROM merch_plot_to_rh_to_collector_haul_costs_work_table a," + 
				                    "(SELECT biosum_plot_id, MIN(total_haul_cost) AS min_cost2 " + 
				                     "FROM merch_plot_to_rh_to_collector_haul_costs_work_table " + 
                                     "GROUP BY biosum_plot_id) b," + 
                                    "(SELECT biosum_plot_id, psite_id," + 
				                            "MIN(total_haul_cost) AS min_cost " + 
				                     "FROM merch_plot_to_rh_to_collector_haul_costs_work_table " + 
				                     "GROUP BY biosum_plot_id,psite_id) c " + 
				               "WHERE  c.biosum_plot_id = b.biosum_plot_id AND " + 
				                      "a.biosum_plot_id = c.biosum_plot_id AND " + 
								      "a.psite_id = c.psite_id AND " + 
				                      "a.total_haul_cost = c.min_cost AND " + 
				                      "min_cost2 = min_cost;";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Find the cheapest plot to merch processing site rail routes");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 12;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     

			this.m_frmRunCoreScenario.lblMsg.Text="Rail Haul Costs For Chip Wood Processing Sites";
			this.m_frmRunCoreScenario.lblMsg.Refresh();

			/***********************************************************************
			 **Append to a table all travel time collector_id (psite) records
			 **where the psite has rail access and processes chips only or 
			 **both merch/chips
			 ***********************************************************************/
			this.m_strSQL = "INSERT INTO chip_rh_to_collector_haul_costs_work_table " + 
				"SELECT  t.psite_id AS railhead_id,"  + 
						"t.collector_id AS psite_id," + 
						"(" + strTransferBioCost.Trim() + " * t.travel_time)  AS transfer_cost," + 
						"0 AS road_cost," + 
						"((t.travel_time * 45) * " + strRailHaulCost.Trim() + ") AS rail_cost," + 
						"0 AS total_haul_cost,  'C' AS materialcd " + 
				"FROM " + this.m_strTravelTimeTable + " t " + 
				"INNER JOIN  " + this.m_strPSiteTable + " s " + 
				"ON t.collector_id = s.psite_id " + 
				"WHERE s.trancd=3 AND  " + 
				       "(s.biocd=3 OR s.biocd=2) AND " + 
				"EXISTS (SELECT ss.psite_id " + 
				"FROM " + this.m_strPSiteTable + " ss " + 
				"WHERE t.psite_id=ss.psite_id AND ss.trancd=2 AND (ss.biocd=3 Or ss.biocd=2));";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table all travel time collector_id (psite) records where the psite has rail access and processes chips.");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 13;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
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
				                        "r.transfer_cost," + 
						                "(" + strTruckHaulCost.Trim() + " * t.travel_time) AS road_cost," + 
				                        "r.rail_cost, " + 
				                        "(r.transfer_cost + road_cost + r.rail_cost) AS total_haul_cost," + 
				                        "'C' AS materialcd " + 
				                "FROM  " + this.m_strTravelTimeTable + " t," + 
				                       "chip_rh_to_collector_haul_costs_work_table r " + 
				                "WHERE r.railhead_id = t.psite_id;";




			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table travel time plot records and previous rail/chips work table results");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 14;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}    
 

			
			
			
			
			
			
			this.m_strSQL = "UPDATE chip_plot_to_rh_to_collector_haul_costs_work_table " + 
				"SET total_haul_cost = transfer_cost + road_cost + rail_cost;";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("update chips by road and rail total haul cost ");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 15;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
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
				                        "a.transfer_cost, a.road_cost, a.rail_cost," + 
				                        "b.min_cost AS total_haul_cost,'C' AS materialcd " + 
				                 "FROM chip_plot_to_rh_to_collector_haul_costs_work_table a," + 
				                    "(SELECT biosum_plot_id, " + 
				                            "MIN(total_haul_cost) AS min_cost2 " + 
				                     "FROM chip_plot_to_rh_to_collector_haul_costs_work_table " +
				                     "GROUP BY biosum_plot_id) c, " +
				                    "(SELECT biosum_plot_id, psite_id," + 
				                            "MIN(total_haul_cost) AS min_cost " + 
				                     "FROM chip_plot_to_rh_to_collector_haul_costs_work_table " + 
				                     "GROUP BY biosum_plot_id,psite_id) b " + 
				                 "WHERE  b.biosum_plot_id = c.biosum_plot_id AND " + 
				                        "a.biosum_plot_id = b.biosum_plot_id AND " + 
				                        "a.psite_id = b.psite_id AND  " + 
				                        "a.total_haul_cost = b.min_cost AND " + 
				                        "min_cost2 = min_cost;";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Find the cheapest plot to chip processing site rail routes");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 16;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     

			this.m_frmRunCoreScenario.lblMsg.Text="Combine Road And Rail Haul Costs For Merchantable Wood Processing Sites";
			this.m_frmRunCoreScenario.lblMsg.Refresh();


            /**************************************************************
			 **combine the cheapest road and rail total cost for each plot
			 **to a merch psite
			 **After the insert there should be two records for each
			 **plot - one with cheapest haul cost by road and another
			 **with cheapest haul cost by rail
			 ***************************************************************/
			this.m_strSQL = "INSERT INTO combine_merch_rail_road_haul_costs_work_table " + 
				               "SELECT * FROM cheapest_road_merch_haul_costs_work_table;";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Cheapest road route to merch psite ");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 17;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     

			this.m_strSQL = "INSERT INTO combine_merch_rail_road_haul_costs_work_table " + 
				"SELECT * FROM cheapest_rail_merch_haul_costs_work_table;";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Cheapest rail route to merch psite ");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 18;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}   
  

            /***************************************************
			 **Get the overall cheapest merch route
			 ***************************************************/
			this.m_frmRunCoreScenario.lblMsg.Text="Get Overall Least Expensive Merch Route";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
            this.m_strSQL = "INSERT INTO cheapest_merch_haul_costs_work_table " + 
				                "SELECT a.biosum_plot_id,b.psite_id, a.railhead_id," + 
				                       "a.transfer_cost, a.road_cost,  a.rail_cost," + 
				                       "b.min_cost AS total_haul_cost,'M' AS materialcd " + 
				                "FROM combine_merch_rail_road_haul_costs_work_table a," + 
				                    "(SELECT biosum_plot_id, MIN(total_haul_cost) AS min_cost2 " + 
				                     "FROM combine_merch_rail_road_haul_costs_work_table " + 
				                     "GROUP BY biosum_plot_id) c, " + 
				                    "(SELECT biosum_plot_id, psite_id," + 
				                            "MIN(total_haul_cost) AS min_cost " + 
				                     "FROM combine_merch_rail_road_haul_costs_work_table " + 
				                     "GROUP BY biosum_plot_id,psite_id) b " + 
				                "WHERE  b.biosum_plot_id = c.biosum_plot_id AND " + 
				                       "a.biosum_plot_id = b.biosum_plot_id AND " + 
				                       "a.psite_id = b.psite_id AND " + 
				                       "a.total_haul_cost = b.min_cost AND " + 
				                       "min_cost2 = min_cost;";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Get the overall cheapest merch route");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 19;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}   

			this.m_frmRunCoreScenario.lblMsg.Text="Combine Road And Rail Haul Costs For Chip Wood Processing Sites";
			this.m_frmRunCoreScenario.lblMsg.Refresh();

			/**************************************************************
			 **combine the cheapest road and rail total cost for each plot
			 **to a chips psite
			 **After the insert there should be two records for each
			 **plot - one with cheapest haul cost by road and another
			 **with cheapest haul cost by rail
			 ***************************************************************/
			this.m_strSQL = "INSERT INTO combine_chip_rail_road_haul_costs_work_table " + 
				"SELECT * FROM cheapest_road_chip_haul_costs_work_table;";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Cheapest road route to chip psite ");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 20;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     

			this.m_strSQL = "INSERT INTO combine_chip_rail_road_haul_costs_work_table " + 
				"SELECT * FROM cheapest_rail_chip_haul_costs_work_table;";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Cheapest rail route to chip psite ");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 21;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}   
  
			this.m_frmRunCoreScenario.lblMsg.Text="Get Overall Least Expensive Chip Route";
			this.m_frmRunCoreScenario.lblMsg.Refresh();


            /******************************************
			 **Get the overall cheapest chip route
			 ******************************************/
			this.m_strSQL = "INSERT INTO cheapest_chip_haul_costs_work_table " + 
				                 "SELECT a.biosum_plot_id,b.psite_id, a.railhead_id," + 
				                        "a.transfer_cost, a.road_cost,  a.rail_cost," + 
				                        "b.min_cost AS total_haul_cost,'C' AS materialcd " + 
				                 "FROM combine_chip_rail_road_haul_costs_work_table a, " + 
				                     "(SELECT biosum_plot_id,MIN(total_haul_cost) AS min_cost2 " + 
				                      "FROM combine_chip_rail_road_haul_costs_work_table " + 
				                      "GROUP BY biosum_plot_id) c, " + 
				                     "(SELECT biosum_plot_id, psite_id," + 
				                             "MIN(total_haul_cost) AS min_cost " + 
				                      "FROM combine_chip_rail_road_haul_costs_work_table " + 
				                      "GROUP BY biosum_plot_id,psite_id) b " + 
				                 "WHERE  b.biosum_plot_id = c.biosum_plot_id AND " + 
				                        "a.biosum_plot_id = b.biosum_plot_id AND " + 
				                        "a.psite_id = b.psite_id AND " + 
				                        "a.total_haul_cost = b.min_cost AND " + 
				                        "min_cost2 = min_cost;";



			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table. Get the overall cheapest chip route");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 22;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}   




			//INSERT INTO HAUL_COSTS TABLE
			this.m_frmRunCoreScenario.lblMsg.Text="Inserting Results Into Haul Costs Table";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_strSQL = "INSERT INTO haul_costs " + 
				                 "SELECT * FROM cheapest_merch_haul_costs_work_table;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into haul_costs table cheapest merch route for each plot");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 23;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}   


			this.m_strSQL = "INSERT INTO haul_costs " + 
				"SELECT * FROM cheapest_chip_haul_costs_work_table;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into haul_costs table cheapest chip route for each plot");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 24;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}   

			//UPDATE PLOT TABLE

			/**************************************************
			 **Update cheapest merch route fields
			 **************************************************/
			this.m_frmRunCoreScenario.lblMsg.Text="Updating Plot Table";
			this.m_frmRunCoreScenario.lblMsg.Refresh();

            this.m_strSQL = "UPDATE " + this.m_strPlotTable + " p " + 
				               "INNER JOIN haul_costs h " + 
				               "ON p.biosum_plot_id = h.biosum_plot_id " + 
				            "SET p.merch_haul_cost_id = h.haul_cost_id," + 
				                "p.merch_haul_cost_psite = h.psite_id," + 
				                "p.merch_haul_cpa_pt=h.total_haul_cost " + 
				            "WHERE h.materialcd='M';";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update plot merch haul cost fields");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 25;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			} 
  

			/*****************************************
			 **Update  cheapest chip routes
			 *****************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable + " p " + 
				              "INNER JOIN haul_costs h " + 
				              "ON p.biosum_plot_id = h.biosum_plot_id " + 
				            "SET p.chip_haul_cost_id = h.haul_cost_id," + 
				                "p.chip_haul_cost_psite = h.psite_id," + 
				                "p.chip_haul_cpa_pt=h.total_haul_cost " + 
				            "WHERE h.materialcd='C';";


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update plot chip haul cost fields");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 26;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			} 
  

            /******************************************
			 **clean up work tables
			 ******************************************/

			this.m_frmRunCoreScenario.lblMsg.Text="Cleaning Up Haul Cost Work Tables...Stand By";
			this.m_frmRunCoreScenario.lblMsg.Refresh();

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Cleaning up haul cost work tables");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 27;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			} 
			if (this.m_intError == 0)
			{
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Blue;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "Completed";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
			}
			this.m_frmRunCoreScenario.lblMsg.Visible=false;
			this.m_frmRunCoreScenario.progressBar1.Visible=false;
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_frmRunCoreScenario.progressBar1.Refresh();



		}


		/// <summary>
		/// populate the plot table's merch_tvltm and chip_tvltm fields with
		/// the fastest time from plot to wood processing site
		/// </summary>
		private void getTravelTimes()
		{
			
			//int x;
			string strTruckHaulCost="";
			string strRailHaulCost="";
			string strTransferMerchCost="";
			string strTransferBioCost="";

			
			this.m_frmRunCoreScenario.lblMsg.Text="Process Travel Times Unique Id Fields";
			this.m_frmRunCoreScenario.lblMsg.Visible=true;
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_frmRunCoreScenario.progressBar1.Minimum=0;
			this.m_frmRunCoreScenario.progressBar1.Maximum = 10;
			this.m_frmRunCoreScenario.progressBar1.Value=0;
			this.m_frmRunCoreScenario.progressBar1.Visible=true;

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update Plot Table With Merch And Chip Haul Costs");
			this.m_txtStreamWriter.WriteLine("-------------------------------------------------------------");

			/********************************************
			 **get the haul cost per green ton per hour
			 ********************************************/
			strTruckHaulCost = strTruckHaulCost.Replace(",","");
			strTruckHaulCost="";
			if (strTruckHaulCost.Trim().Length == 1) strTruckHaulCost = "0.00";
			
			strRailHaulCost = strTruckHaulCost.Replace(",","");
			if (strRailHaulCost.Trim().Length == 1) strRailHaulCost = "0.00";
			
			/***********************************************
			 **get the transfer cost per green to per hour
			 ***********************************************/
			strTransferMerchCost = strTransferMerchCost.Replace(",","");
			if (strTransferMerchCost.Trim().Length == 1) strTransferMerchCost = "0.00";
			
			strTransferBioCost = strTransferMerchCost.Replace(",","");
			if (strTransferBioCost.Trim().Length == 1) strTransferMerchCost = "0.00";
			 


			/*************************************************************
			 **first make sure each travel times biosum_plot_id field
			 **has a value
			 *************************************************************/					   
			this.m_strSQL = "UPDATE " + this.m_strTravelTimeTable  + " t " + 
				"INNER JOIN " + this.m_strPlotTable + " p ON t.plot_id = p.idb_plot_id " + 
				"SET t.biosum_plot_id = p.biosum_plot_id " + 
				"WHERE t.plot_id IS NOT NULL AND t.biosum_plot_id IS NULL" ;
           this.m_txtStreamWriter.WriteLine(" ");
		   this.m_txtStreamWriter.WriteLine("update " + this.m_strTravelTimeTable.Trim() + ".biosum_cond_id field NULL values");
		   this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
           this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
           this.m_frmRunCoreScenario.progressBar1.Value = 1;
		   if (this.m_ado.m_intError != 0)
		   {
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
		   }     
			this.m_frmRunCoreScenario.lblMsg.Text="Null The Plot Table's Travel Time Fields";
			this.m_frmRunCoreScenario.lblMsg.Refresh();

			/***************************************************************
			 **null the plots merch_haul_cost_psite, merch_haul_cost_id and 
			 **merch_haul_cpa_pt,chip_haul_cost_psite,chip_haul_cost_id 
			 **and chip_haul_cpa_pt  fields
			 ***************************************************************/
            this.m_strSQL = "UPDATE " + this.m_strPlotTable + " " + 
                            "SET merch_haul_cost_id = null," + 
				                "merch_haul_cost_psite = null, " +
				                "merch_haul_cpa_pt = null ," + 
				                "chip_haul_cost_id = null," + 
				                "chip_haul_cost_psite = null," + 
							    "chip_haul_cpa_pt = null;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("null the plot table's haul cost fields");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 2;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     
			
			

	   /*****************************************************************
			**delete any records that may exist in the work tables
			*****************************************************************/
		   this.m_strSQL = "delete from haul_costs;";
		   this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
       this.m_strSQL = "delete from haul_costs_work_table;";
		   this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
       this.m_strSQL = "delete from haul_costs_work_table2;";
		   this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
		   		   
    		if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 3;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     
     //MERCH PROCESSING SITE HAUL COSTS
			this.m_frmRunCoreScenario.lblMsg.Text="Haul Costs For Merchantable Wood Processing Sites";
			this.m_frmRunCoreScenario.lblMsg.Refresh();

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Merchantable wood processing site haul costs");
		   /***************************************************************
		    **process the merch travel times first
			  ***************************************************************/
			  
			  
			  
		   /********************************************************************
			**insert all travel times for merch processing sites
			********************************************************************/
           this.m_strSQL = "INSERT INTO plot_haul_cost_work_table " + 
			                 "(biosum_plot_id, merch_haul_cost_psite,merch_tvltm) " + 
			                    "SELECT t.biosum_plot_id, " + 
			                           "s.psite_id, " + 
			                           "t.travel_time " + 
			                    "FROM travel_time t," + 
			                          "processing_site s " + 
			                    "WHERE t.psite_id = s.psite_id AND " + 
			                         "(s.biocd = 1 OR s.biocd=3);";
		   this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table merch processing site travel times");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
		   this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 3;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     

           /************************************************************************
			**insert into work table 2 the fastest merch travel time for each plot 
			************************************************************************/
		   this.m_strSQL = "INSERT INTO plot_fastest_tvltm_work_table2 " + 
			                 "(biosum_plot_id, merch_fastest_tvltm_psite, merch_tvltm) " + 
			                    "SELECT b.biosum_plot_id, " +
			                           "c.merch_fastest_tvltm_psite, " + 
			                           "b.min_travel AS merch_tvltm " + 
			                    "FROM plot_fastest_tvltm_work_table a, " + 
			                     "(SELECT biosum_plot_id, " + 
			                             "MIN(merch_tvltm) as min_travel " +  
			                      "FROM plot_fastest_tvltm_work_table " + 
			                      "GROUP BY biosum_plot_id)  b," + 
			                     "(SELECT biosum_plot_id, " + 
			                             "merch_fastest_tvltm_psite," + 
			                             "MIN(merch_tvltm) as min_travel2 " + 
			                      "FROM plot_fastest_tvltm_work_table " + 
			                      "GROUP BY biosum_plot_id, merch_fastest_tvltm_psite) c " + 
			                    "WHERE  c.biosum_plot_id = b.biosum_plot_id AND " + 
			                           "a.biosum_plot_id = b.biosum_plot_id AND " + 
									   "a.merch_fastest_tvltm_psite = c.merch_fastest_tvltm_psite AND  " + 
			                           "b.min_travel = c.min_travel2;";

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table the fastest merch processing site travel times");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 5;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     

			 /***********************************************************
			  **okay, update the plot table with the fastest merch times 
			  ***********************************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable + " a " +  
                              "INNER JOIN plot_fastest_tvltm_work_table2  b " + 
				                 "ON a.biosum_plot_id = b.biosum_plot_id " + 
				              "SET a.merch_fastest_tvltm_psite = b.merch_fastest_tvltm_psite," + 
				                  "a.merch_tvltm = b.merch_tvltm," + 
				                  "a.merch_haul_cpa_pt = b.merch_tvltm * " + strTruckHaulCost.Trim() + ";";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("update plot table from the work table");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL); 
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 6;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     


			/*****************************************************************
			**delete any records that may exist in the work tables
			*****************************************************************/
			this.m_strSQL = "delete from plot_fastest_tvltm_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			this.m_strSQL = "delete from plot_fastest_tvltm_work_table2";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			//Chip PROCESSING SITE TRAVEL TIMES
			this.m_frmRunCoreScenario.lblMsg.Text="Travel Times For Chip Wood Processing Sites";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Chip wood processing site travel times");
			/***************************************************************
			 **process the Chip travel times
			 ***************************************************************/
			/********************************************************************
			 **insert all travel times for Chips processing sites
			 ********************************************************************/
			this.m_strSQL = "INSERT INTO plot_fastest_tvltm_work_table " + 
				"(biosum_plot_id, chip_fastest_tvltm_psite,chip_tvltm) " + 
				"SELECT t.biosum_plot_id, " + 
				"s.psite_id, " + 
				"t.travel_time " + 
				"FROM travel_time t," + 
				"processing_site s " + 
				"WHERE t.psite_id = s.psite_id AND " + 
				"(s.biocd = 2 OR s.biocd=3);";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table Chip processing site travel times");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 7;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     

			/************************************************************************
			 **insert into work table 2 the fastest merch travel time for each plot 
			 ************************************************************************/
			this.m_strSQL = "INSERT INTO plot_fastest_tvltm_work_table2 " + 
				"(biosum_plot_id, chip_fastest_tvltm_psite, chip_tvltm) " + 
				"SELECT b.biosum_plot_id, " +
				"c.chip_fastest_tvltm_psite, " + 
				"b.min_travel AS chip_tvltm " + 
				"FROM plot_fastest_tvltm_work_table a, " + 
				"(SELECT biosum_plot_id, " + 
				"MIN(chip_tvltm) as min_travel " +  
				"FROM plot_fastest_tvltm_work_table " + 
				"GROUP BY biosum_plot_id)  b," + 
				"(SELECT biosum_plot_id, " + 
				"chip_fastest_tvltm_psite," + 
				"MIN(chip_tvltm) as min_travel2 " + 
				"FROM plot_fastest_tvltm_work_table " + 
				"GROUP BY biosum_plot_id, chip_fastest_tvltm_psite) c " + 
				"WHERE  c.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.biosum_plot_id = b.biosum_plot_id AND " + 
				"a.chip_fastest_tvltm_psite = c.chip_fastest_tvltm_psite AND  " + 
				"b.min_travel = c.min_travel2;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into work table the fastest Chip processing site travel times");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 9;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     

			/***********************************************************
			  **okay, update the plot table with the fastest Chips times 
			  ***********************************************************/
			this.m_strSQL = "UPDATE " + this.m_strPlotTable + " a " +  
				"INNER JOIN plot_fastest_tvltm_work_table2  b " + 
				"ON a.biosum_plot_id = b.biosum_plot_id " + 
				"SET a.chip_fastest_tvltm_psite = b.chip_fastest_tvltm_psite," + 
				"a.chip_tvltm = b.chip_tvltm," + 
				"a.chip_haul_cpa_pt = b.chip_tvltm * " + strTruckHaulCost.Trim() + ";";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("update plot table from the work table");
			this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL); 
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcTravelTimes)) return;
			this.m_frmRunCoreScenario.progressBar1.Value = 10;
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
				return;
			}     

			if (this.m_intError == 0)
			{
				this.m_frmRunCoreScenario.lblProcTravelTimes.ForeColor = System.Drawing.Color.Blue;
				this.m_frmRunCoreScenario.lblProcTravelTimes.Text = "Completed";
				this.m_frmRunCoreScenario.lblProcTravelTimes.Refresh();
			}
			this.m_frmRunCoreScenario.lblMsg.Visible=false;
			this.m_frmRunCoreScenario.progressBar1.Visible=false;
		    this.m_frmRunCoreScenario.lblMsg.Refresh();
            this.m_frmRunCoreScenario.progressBar1.Refresh();



		}
		/// <summary>
		/// sum the tree_vol_val_by_species_diam_groups table values to tree_vol_val_sum_by_rx
		/// </summary>
		private void sumTreeVolVal()
		{
			
			
			/**************************************************************
			 **sum the tree_vol_val_by_species_diam_groups table to
			 **        tree_vol_val_sum_by_rx
			 **************************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Sum Tree Volumes and Values");
			this.m_txtStreamWriter.WriteLine("---------------------------");

			this.m_strSQL = "delete from tree_vol_val_sum_by_rx";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.m_ado.m_intError != 0)
			{
				this.m_frmRunCoreScenario.lblProcSumTree.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcSumTree.Text = "!!Error!!";
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcSumTree.Refresh();
				return;
			}
			this.m_strSQL = "INSERT INTO tree_vol_val_sum_by_rx " + 
				"(biosum_cond_id,rx,chip_vol_cf," + 
				"chip_wt_gt,chip_val_dpa,merch_vol_cf," + 
				"merch_wt_gt,merch_val_dpa) ";
			
			this.m_strSQL += "SELECT biosum_cond_id, " + 
				"rx," +
				"Sum(chip_vol_cf) AS chip_vol_cf," + 
				"Sum(chip_wt_gt) AS chip_wt_gt," + 
				"Sum(chip_val_dpa) AS chip_val_dpa," + 
				"Sum(merch_vol_cf) AS merch_vol_cf," +
				"Sum(merch_wt_gt) AS merch_wt_gt," + 
				"Sum(merch_val_dpa) AS merch_val_dpa ";

			this.m_strSQL += " FROM " + this.m_strTreeVolValBySpcDiamGroupsTable.Trim();
			this.m_strSQL += " GROUP BY biosum_cond_id,rx";
			this.m_strSQL += " ORDER BY biosum_cond_id,rx;";
            this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("insert into tree_vol_val_sum_by_rx table tree volume and value sums");
		    this.m_txtStreamWriter.WriteLine("Execute SQL: " + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.m_ado.m_intError != 0)
			{
				this.m_frmRunCoreScenario.lblProcSumTree.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcSumTree.Text = "!!Error!!";
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcSumTree.Refresh();
				return;
			}
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcSumTree)) return;


			if (this.m_intError == 0)
			{
				this.m_frmRunCoreScenario.lblProcSumTree.ForeColor = System.Drawing.Color.Blue;
				this.m_frmRunCoreScenario.lblProcSumTree.Text = "Completed";
				this.m_frmRunCoreScenario.lblProcSumTree.Refresh();
			}
		}
		private void getHaulCost(string p_strBiosumPlotId, int p_intPSiteId)
		{
			/********************************************
			 **get the haul cost per green ton per hour
			 ********************************************/
            string strTruckHaulCost = "0.00";
			
			strTruckHaulCost = strTruckHaulCost.Replace(",","");
			
			if (strTruckHaulCost.Trim().Length == 1) strTruckHaulCost = "0.00";

            string strRailHaulCost = "0.00";
			
			strRailHaulCost = strRailHaulCost.Replace(",","");
			if (strRailHaulCost.Trim().Length == 1) strRailHaulCost = "0.00";
			
			/***********************************************
			 **get the transfer cost per green to per hour
			 ***********************************************/
            string strTransferMerchCost = "0.00";
			
			strTransferMerchCost = strTransferMerchCost.Replace(",","");
			if (strTransferMerchCost.Trim().Length == 1) strTransferMerchCost = "0.00";

            string strTransferBioCost = "0.00";
			
			strTransferBioCost = strTransferBioCost.Replace(",","");
			if (strTransferBioCost.Trim().Length == 1) strTransferBioCost = "0.00";
			



		}

		/// <summary>
		/// load the validcombos table with biosum_cond_id and rx values
		/// that exist in the user defined plot filters, condition, ffe, travel times, and
		/// harvest cost, and tree volume/value tables
		/// </summary>
		private void validcombos()
		{
			string strRxList="";
			string strGrpCd="";
			

             this.m_txtStreamWriter.WriteLine(" ");
			 this.m_txtStreamWriter.WriteLine(" ");
			 this.m_txtStreamWriter.WriteLine("Valid Combinations");
			 this.m_txtStreamWriter.WriteLine("------------------");

			/*****************************************************************
			 **delete audit tables
			 *****************************************************************/
			
			this.m_strSQL = "delete from plot_cond_audit";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.m_ado.m_intError != 0)
			{
				this.m_frmRunCoreScenario.lblProcValidCombos.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcValidCombos.Text = "!!Error!!";
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcValidCombos.Refresh();
				return;
			}
			this.m_strSQL = "delete from plot_cond_rx_audit";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.m_ado.m_intError != 0)
			{
				this.m_frmRunCoreScenario.lblProcValidCombos.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcValidCombos.Text = "!!Error!!";
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcValidCombos.Refresh();
				return;
			}

			/**************************
			 **get the treatment list
			 **************************/
			this.m_strSQL = "SELECT rx FROM " + this.m_strRxTable + ";"; // WHERE trim(ucase(scenario_id)) = '" + this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
			this.m_ado.SqlQueryReader(this.m_TempMDBFileConn,this.m_strSQL);
			if (!this.m_ado.m_OleDbDataReader.HasRows)
			{
				this.m_frmRunCoreScenario.lblProcValidCombos.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcValidCombos.Text = "!!Error!!";
				this.m_intError = -1;
				this.m_frmRunCoreScenario.lblProcValidCombos.Refresh();
				MessageBox.Show("No Treatments Found In The Treatment Table");
				return;
			}
			while (this.m_ado.m_OleDbDataReader.Read())
			{
				strRxList+=this.m_ado.m_OleDbDataReader["rx"].ToString().Trim();
			}
			this.m_ado.m_OleDbDataReader.Close();


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Execute User Defined SQL And Insert Resulting Records Into Table userdefinedplotfilter");
			this.m_strSQL = "INSERT INTO userdefinedplotfilter " + this.m_strUserDefinedPlotSQL;
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			
			
			this.m_strSQL = "INSERT INTO ruledefinitionscondfilter SELECT * FROM " + this.m_strCondTable + " c ";
			this.m_strSQL += " WHERE c.owngrpcd IN (";

			//usfs ownnership
			if (this.m_frmRunCoreScenario.m_frmScenario.uc_scenario_owner_groups1.chkOwnGrp10.Checked==true)
			{
				strGrpCd = "10,1";
			}
			//other federal ownership
			if (this.m_frmRunCoreScenario.m_frmScenario.uc_scenario_owner_groups1.chkOwnGrp20.Checked==true)
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
			if (this.m_frmRunCoreScenario.m_frmScenario.uc_scenario_owner_groups1.chkOwnGrp30.Checked==true)
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
			if (this.m_frmRunCoreScenario.m_frmScenario.uc_scenario_owner_groups1.chkOwnGrp40.Checked==true)
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
			this.m_strSQL +=  strGrpCd + ") AND c.cond_accessible_yn = 'Y';";

			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Execute SQL That Includes  Rule Definitions Defined By The User Into Table ruledefinitionsplotfilter");
			this.m_strSQL = "INSERT INTO ruledefinitionsplotfilter SELECT DISTINCT userdefinedplotfilter.* from userdefinedplotfilter INNER JOIN ruledefinitionscondfilter ON userdefinedplotfilter.biosum_plot_id = ruledefinitionscondfilter.biosum_plot_id" ;
			this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

            this.m_strSQL = "DELETE FROM validcombos;";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/**********************************************************************
			 **create valid combiniations of biosum_cond_id and treatment and 
			 **also user defined filters by owngrpcd
			 **********************************************************************/
			this.m_strSQL = "INSERT INTO validcombos (biosum_cond_id,rx) SELECT DISTINCT ruledefinitionscondfilter.biosum_cond_id," + this.m_strFFETable + ".rx " +
				"FROM (((ruledefinitionsplotfilter INNER JOIN ruledefinitionscondfilter ON ruledefinitionsplotfilter.biosum_plot_id = ruledefinitionscondfilter.biosum_plot_id) " + 
				" INNER JOIN " + this.m_strFFETable + " ON ruledefinitionscondfilter.biosum_cond_id = " + this.m_strFFETable + ".biosum_cond_id) " +  
				" INNER JOIN " + this.m_strHvstCostsTable + " ON (" + this.m_strFFETable + ".rx = " + this.m_strHvstCostsTable + ".rx) AND (" + this.m_strFFETable + ".biosum_cond_id = " + this.m_strHvstCostsTable + ".biosum_cond_id)) " + 
				" INNER JOIN " + this.m_strTreeVolValSumTable + " ON (" + this.m_strHvstCostsTable + ".biosum_cond_id = " + this.m_strTreeVolValSumTable + ".biosum_cond_id) AND (" + this.m_strHvstCostsTable + ".rx = " + this.m_strTreeVolValSumTable + ".rx) AND (" + this.m_strFFETable + ".biosum_cond_id = " + this.m_strTreeVolValSumTable + ".biosum_cond_id) AND (" + this.m_strFFETable + ".rx = " + this.m_strTreeVolValSumTable + ".rx);";


            this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Create Valid Combinations");
            this.m_txtStreamWriter.WriteLine("Execute SQL:{0}",this.m_strSQL);
			
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos) == true) return;

			if (this.m_frmRunCoreScenario.chkAuditTables.Checked == true)
			{
				this.m_frmRunCoreScenario.lblMsg.Text="Creating Audit Data";
				this.m_frmRunCoreScenario.lblMsg.Visible=true;
				this.m_frmRunCoreScenario.lblMsg.Refresh();
				this.m_frmRunCoreScenario.progressBar1.Minimum=0;
				this.m_frmRunCoreScenario.progressBar1.Maximum = 16;
				this.m_frmRunCoreScenario.progressBar1.Value=0;
				this.m_frmRunCoreScenario.progressBar1.Visible=true;

				//BIOSUM_COND_ID RECORD AUDIT
				/******************************************************************************************
				 **insert all the plots that are being processed into the plot audit table
				 ******************************************************************************************/
				this.m_strSQL = "INSERT INTO plot_cond_audit (biosum_cond_id) SELECT ruledefinitionscondfilter.biosum_cond_id FROM ruledefinitionscondfilter INNER JOIN userdefinedplotfilter ON ruledefinitionscondfilter.biosum_plot_id = userdefinedplotfilter.biosum_plot_id";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				if (this.m_ado.m_intError != 0)
				{
					this.m_frmRunCoreScenario.lblProcValidCombos.ForeColor = System.Drawing.Color.Red;
					this.m_frmRunCoreScenario.lblProcValidCombos.Text = "!!Error!!";
					this.m_intError = this.m_ado.m_intError;
					this.m_frmRunCoreScenario.lblProcValidCombos.Refresh();
					return;
				}
                this.m_frmRunCoreScenario.progressBar1.Value=1;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;

				/************************************************************************
				 **check to see if the plot record exists in the frcs harvest cost table
				 ************************************************************************/
				this.m_strSQL = "UPDATE plot_cond_audit SET harvest_costs_yn = 'Y' " + 
					             "WHERE plot_cond_audit.biosum_cond_id " + 
					             "IN (SELECT biosum_cond_id FROM " + this.m_strHvstCostsTable + ");";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=2;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;

				this.m_strSQL = "UPDATE plot_cond_audit SET harvest_costs_yn = 'N' " + 
					"WHERE plot_cond_audit.harvest_costs_yn IS NULL ;" ;
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
                this.m_frmRunCoreScenario.progressBar1.Value=3;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;

				/*******************************************************************
				 **check to see if the plot record exists in the fvs ffe  table
				 *******************************************************************/
				this.m_strSQL = "UPDATE plot_cond_audit SET fvs_ffe_yn = 'Y' " + 
					"WHERE plot_cond_audit.biosum_cond_id " + 
					"IN (SELECT biosum_cond_id FROM " + this.m_strFFETable + ");";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=4;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;

				this.m_strSQL = "UPDATE plot_cond_audit SET fvs_ffe_yn = 'N' " + 
					"WHERE plot_cond_audit.fvs_ffe_yn IS NULL OR plot_cond_audit.fvs_ffe_yn <> 'Y' ;"; 
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=5;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;

				/********************************************************************************************************
				 **check to see if the plot record exists in the processor tree volume and value tableharvest cost table
				 ********************************************************************************************************/
				this.m_strSQL = "UPDATE plot_cond_audit SET processor_tree_vol_val_yn = 'Y' " + 
					"WHERE plot_cond_audit.biosum_cond_id " + 
					"IN (SELECT biosum_cond_id FROM " + this.m_strTreeVolValSumTable + ");";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=6;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;

				this.m_strSQL = "UPDATE plot_cond_audit SET processor_tree_vol_val_yn = 'N' " + 
					"WHERE plot_cond_audit.processor_tree_vol_val_yn IS NULL OR  plot_cond_audit.processor_tree_vol_val_yn<>'Y' ;"; 
                this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=7;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;


				/**********************************************************************
				 **check to see if the plot record exists in the gis travel times table
				 **********************************************************************/
				this.m_strSQL = "UPDATE plot_cond_audit SET gis_travel_times_yn = 'Y' " + 
					"WHERE plot_cond_audit.biosum_cond_id " + 
					"IN (SELECT biosum_cond_id FROM ruledefinitionscondfilter "  +
					   " WHERE ruledefinitionscondfilter.biosum_plot_id " + 
					   " IN (SELECT biosum_plot_id FROM " + this.m_strTravelTimeTable + "));";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=8;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;
				
				this.m_strSQL = "UPDATE plot_cond_audit SET gis_travel_times_yn = 'N' " + 
					"WHERE plot_cond_audit.gis_travel_times_yn IS NULL OR  plot_cond_audit.gis_travel_times_yn<>'Y' ;"; 
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=9;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;

				//BIOSUM_COND_ID + RX RECORD AUDIT
				/**********************************************************************************
				**Insert all the biosum_cond_id + rx combinations into the plot_cond_rx_audit table
				***********************************************************************************/
				this.m_strSQL = "INSERT INTO plot_cond_rx_audit (biosum_cond_id,rx)  " + 
					" SELECT a.biosum_cond_id, b.rx FROM plot_cond_audit a, " + 
                     "(SELECT DISTINCT rx FROM " + this.m_strRxTable + ") b ;";  //+ 
					
			    this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=10;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;

				/**********************************************************************
				 **check to see if the plot + rx record exists in the fvs ffe table
				 **********************************************************************/
               this.m_strSQL="UPDATE plot_cond_rx_audit SET fvs_ffe_yn = 'Y' " + 
                             "WHERE EXISTS (SELECT biosum_cond_id,rx " + 
				                           "FROM "  + this.m_strFFETable + " " + 
				                           "WHERE plot_cond_rx_audit.biosum_cond_id = " + 
				                                  this.m_strFFETable.Trim() + ".biosum_cond_id AND " + 
				                                 "plot_cond_rx_audit.rx = " + 
				                                  this.m_strFFETable.Trim() + ".rx);";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=11;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;

				this.m_strSQL = "UPDATE plot_cond_rx_audit SET fvs_ffe_yn = 'N' " + 
					"WHERE plot_cond_rx_audit.fvs_ffe_yn IS NULL ;" ;
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=12;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;

				/****************************************************************************
				 **check to see if the plot + rx record exists in the frcs harves costs table
				 ****************************************************************************/
				this.m_strSQL="UPDATE plot_cond_rx_audit SET harvest_costs_yn = 'Y' " + 
					"WHERE EXISTS (SELECT biosum_cond_id,rx " + 
					"FROM "  + this.m_strHvstCostsTable + " " + 
					"WHERE plot_cond_rx_audit.biosum_cond_id = " + 
					this.m_strHvstCostsTable.Trim() + ".biosum_cond_id AND " + 
					"plot_cond_rx_audit.rx = " + 
					this.m_strHvstCostsTable.Trim() + ".rx);";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=13;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;

				this.m_strSQL = "UPDATE plot_cond_rx_audit SET harvest_costs_yn = 'N' " + 
					"WHERE plot_cond_rx_audit.harvest_costs_yn IS NULL ;" ;
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=14;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;

				/*********************************************************************************
				 **check to see if the plot + rx record exists in the processor tree vol val table
				 *********************************************************************************/
				this.m_strSQL="UPDATE plot_cond_rx_audit SET processor_tree_vol_val_yn = 'Y' " + 
					"WHERE EXISTS (SELECT biosum_cond_id,rx " + 
					"FROM "  + this.m_strTreeVolValSumTable + " " + 
					"WHERE plot_cond_rx_audit.biosum_cond_id = " + 
					this.m_strTreeVolValSumTable.Trim() + ".biosum_cond_id AND " + 
					"plot_cond_rx_audit.rx = " + 
					this.m_strTreeVolValSumTable.Trim() + ".rx);";
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=15;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;

				this.m_strSQL = "UPDATE plot_cond_rx_audit SET processor_tree_vol_val_yn = 'N' " + 
					"WHERE plot_cond_rx_audit.processor_tree_vol_val_yn IS NULL ;" ;
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
				this.m_frmRunCoreScenario.progressBar1.Value=16;
				if (this.UserCancel(this.m_frmRunCoreScenario.lblProcValidCombos)) return;


				this.m_frmRunCoreScenario.progressBar1.Visible=false;
				this.m_frmRunCoreScenario.lblMsg.Visible=false;
				this.m_frmRunCoreScenario.progressBar1.Refresh();
				this.m_frmRunCoreScenario.lblMsg.Text="";
				this.m_frmRunCoreScenario.lblMsg.Refresh();

			}
			if (this.m_intError == 0)
			{
				this.m_frmRunCoreScenario.lblProcValidCombos.ForeColor = System.Drawing.Color.Blue;
				this.m_frmRunCoreScenario.lblProcValidCombos.Text = "Completed";
				this.m_frmRunCoreScenario.lblProcValidCombos.Refresh();
			}
			
		}

		/// <summary>
		/// evaluate the effectiveness of fire and fuel treatment data 
		/// by loading the effective table with 
		/// results from user defined expressions 
		/// </summary>
		private void effective()
		{
			string strPreTIFFEWindSpeedSQL="";
			string strPreCIFFEWindSpeedSQL="";
			string strPostTIFFEWindSpeedSQL="";
			string strPostCIFFEWindSpeedSQL="";
			string strTIChangeSQL="";
			string strCIChangeSQL="";
			string strTIEffectiveSQL="";
			string strCIEffectiveSQL="";
			string strTIBackslideSQL="";
			string strCIBackslideSQL="";
			string strAtHazardSQL="";
			string strOverallEffectiveSQL="";
			string strSelectSQL;
			int intCount=0;
			int x;
			string strScenarioMDB="";
			string strScenarioConn="";
            
			this.m_txtStreamWriter.WriteLine(" ");
            this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Effective Treatments");
			this.m_txtStreamWriter.WriteLine("--------------------");

            /********************************************
			 **delete all records in the effective table
			 ********************************************/
			this.m_strSQL = "delete from effective";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			ado_data_access p_ado = new ado_data_access();
			p_ado.getScenarioConnStringAndMDBFile(ref strScenarioMDB,ref strScenarioConn, ((frmMain)this.m_frmRunCoreScenario.m_frmScenario.ParentForm).frmProject.uc_project1.txtRootDirectory.Text);
			p_ado.OpenConnection(strScenarioConn);
			if (p_ado.m_intError != 0)
			{
				this.m_intError = p_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcEffective.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcEffective.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcEffective.Refresh();
				p_ado = null;
				return;
			}     

			/**********************************************************
			 **create the torch index wind speed class sql expression
			 **********************************************************/

			//MessageBox.Show(strPreTIFFEWindSpeedSQL);


			/**********************************************************
			 **create the crown index wind speed class sql expression
			 **********************************************************/

			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcEffective) == true) return;
             
            
			if (this.m_intError == 0)
			{
				this.m_frmRunCoreScenario.lblProcEffective.ForeColor = System.Drawing.Color.Blue;
				this.m_frmRunCoreScenario.lblProcEffective.Text = "Completed";
				this.m_frmRunCoreScenario.lblProcEffective.Refresh();
			}


		}

		/// <summary>
		/// get the wood product yields,
		/// revenue, and costs of an applied
		/// treatment on a plot 
		/// </summary>
		private void product_yields_net_rev_costs_summary()
		{
			string[]  strUpdateSQL = new string[25];
			int intUpdateCount=0;
			string strSum="";
			string strSumFields="";
			int x=0;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Wood Product Yields,Revenue, And Costs Table");
			this.m_txtStreamWriter.WriteLine("-----------------------------------------------");
			/**********************************************
			 **complete harvest cost per acre
			 **********************************************/
            
			/*****************************************************
			 **get the Chips market value per green ton
			 *****************************************************/
			ado_data_access p_ado = new ado_data_access();
			string strScenarioMDB="";
			string strScenarioConn="";
			p_ado.getScenarioConnStringAndMDBFile(ref strScenarioMDB,ref strScenarioConn, ((frmMain)this.m_frmRunCoreScenario.m_frmScenario.ParentForm).frmProject.uc_project1.txtRootDirectory.Text);
			p_ado.OpenConnection(strScenarioConn);
			if (p_ado.m_intError != 0)
			{
				this.m_intError = p_ado.m_intError;
				this.m_frmRunCoreScenario.lblSumWoodProducts.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblSumWoodProducts.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblSumWoodProducts.Refresh();
				p_ado = null;
				return;
			}     

			/**********************************************************
			 **create the at hazard expression
			 **********************************************************/
			this.m_strSQL = "SELECT * FROM scenario_costs WHERE trim(ucase(scenario_id))= '" + this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
            
			p_ado.SqlQueryReader(p_ado.m_OleDbConnection,this.m_strSQL);
			double dblChipMktValPgt=0;
			double dblDriverCostPgtPerHour=0;
			if (p_ado.m_OleDbDataReader.HasRows==true)
			{
				while (p_ado.m_OleDbDataReader.Read())
				{
					if (p_ado.m_OleDbDataReader["chip_mkt_val_pgt"] != System.DBNull.Value)
					{
						dblChipMktValPgt = Convert.ToDouble(p_ado.m_OleDbDataReader["chip_mkt_val_pgt"].ToString().Trim());
						dblDriverCostPgtPerHour= Convert.ToDouble(p_ado.m_OleDbDataReader["road_haul_cost_pgt_per_hour"].ToString().Trim());
					}																																  
				}
			}
			p_ado.m_OleDbDataReader.Close();
            p_ado.m_OleDbConnection.Close();
			p_ado = null;


			/**********************************************
			 **sum all the expense columns to get complete
			 ** cost for each condition + treatment
			 **********************************************/
            this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Sum All Harvest Costs Per Acre");

			System.Data.DataTable p_dt = this.m_ado.getTableSchema(this.m_TempMDBFileConn,"select * from " + this.m_strHvstCostsTable.Trim());

			for (x=3; x<=p_dt.Rows.Count-1; x++)
			{
				strSumFields = this.m_strHvstCostsTable.Trim() + "." + p_dt.Rows[x]["columnname"].ToString().Trim();
				strUpdateSQL[intUpdateCount] = " UPDATE " + this.m_strHvstCostsTable.Trim() + " SET " + strSumFields + " = 0  WHERE (((" + strSumFields.Trim() + ") Is Null));";
				if (strSum.Trim().Length == 0)
				{
					strSum =  p_dt.Rows[x]["columnname"].ToString().Trim();
				}
				else
				{
					strSum += "+"  + p_dt.Rows[x]["columnname"].ToString().Trim();
				}
				intUpdateCount++;
			}
			this.m_ado.m_OleDbDataReader.Close();

			/******************************************************
			 **update the sum fields to 0 for those that are null
			 ******************************************************/
			for (x=0;x<=intUpdateCount-1;x++)
			{
				this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,strUpdateSQL[x]);
				if (this.m_ado.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
					this.m_intError = this.m_ado.m_intError;
					this.m_frmRunCoreScenario.lblSumWoodProducts.ForeColor = System.Drawing.Color.Red;
					this.m_frmRunCoreScenario.lblSumWoodProducts.Text = "!!Error!!";
					this.m_frmRunCoreScenario.lblSumWoodProducts.Refresh();
					return;
				}
			}
			
            this.m_txtStreamWriter.WriteLine(" ");
            this.m_strSQL = "INSERT INTO harvest_costs_sum " + 
				                " SELECT biosum_cond_id,rx, " + strSum + " AS complete_cpa FROM " + this.m_strHvstCostsTable + ";";

			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.m_ado.m_intError !=0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblSumWoodProducts.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblSumWoodProducts.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblSumWoodProducts.Refresh();
				return;

			}



           // MessageBox.Show(this.m_strSQL);
			/****************************************************************
			 **sum all columns beginning in column 4 and write value to the
			 **harvest_cpa column
			 ****************************************************************/
			
			this.m_strSQL = "UPDATE " + this.m_strHvstCostsTable.Trim() + 
                             " INNER JOIN harvest_costs_sum " + 
				               " ON (" + this.m_strHvstCostsTable.Trim() +  ".biosum_cond_id=" + 
				                             "harvest_costs_sum.biosum_cond_id) AND (" +
			                             this.m_strHvstCostsTable.Trim() + ".rx=" + 
				                             "harvest_costs_sum.rx) " + 
				             "SET " + this.m_strHvstCostsTable.Trim() + ".complete_cpa=harvest_costs_sum.complete_cpa;";
				
				
			//MessageBox.Show(this.m_strSQL);
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblSumWoodProducts.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblSumWoodProducts.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblSumWoodProducts.Refresh();
				return;
			}


			/********************************************
			 **delete all records in the table
			 ********************************************/
			this.m_strSQL = "delete from product_yields_net_rev_costs_summary";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblSumWoodProducts.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblSumWoodProducts.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblSumWoodProducts.Refresh();
				return;
			}
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert Records");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_strSQL = "INSERT INTO product_yields_net_rev_costs_summary " + 
				                 "(biosum_cond_id,rx,merch_yield_cf,chip_yield_cf,chip_yield_gt,chip_val_dpa," + 
				                  "merch_yield_gt,merch_val_dpa,harvest_onsite_cpa," +
				                  "haul_merch_cpa,haul_chip_cpa,merch_chip_nr_dpa," +
				                  "merch_nr_dpa,usebiomass_yn,max_nr_dpa) ";

			string strSelectSQL = "SELECT validcombos.biosum_cond_id,validcombos.rx," + 
                                   this.m_strTreeVolValSumTable.Trim() + ".merch_vol_cf AS merch_yield_cf," +
				                   this.m_strTreeVolValSumTable.Trim() + ".chip_vol_cf AS chip_yield_cf," +
				                   this.m_strTreeVolValSumTable.Trim() + ".chip_wt_gt AS chip_yield_gt," +
				                   this.m_strTreeVolValSumTable.Trim() + ".chip_val_dpa AS chip_val_dpa," + 
				                   this.m_strTreeVolValSumTable.Trim() + ".merch_wt_gt AS merch_yield_gt," +
				                   this.m_strTreeVolValSumTable.Trim() + ".merch_val_dpa AS merch_val_dpa," +
				                   this.m_strHvstCostsTable.Trim() + ".complete_cpa AS harvest_onsite_cpa," + 
				                   this.m_strPlotTable.Trim() + ".merch_haul_cpa_pt * " + this.m_strTreeVolValSumTable.Trim() + ".merch_wt_gt AS haul_merch_cpa," + 
				                   this.m_strPlotTable.Trim() + ".chip_haul_cpa_pt * chip_yield_gt AS haul_chip_cpa," + 
				                   "merch_val_dpa + chip_val_dpa - harvest_onsite_cpa - haul_merch_cpa - haul_chip_cpa AS merch_chip_nr_dpa," + 
				                   "merch_val_dpa - harvest_onsite_cpa - haul_merch_cpa AS merch_nr_dpa," + 
				                   "IIF(merch_chip_nr_dpa > merch_nr_dpa,'Y','N') AS usebiomass_yn," + 
				                   "IIF(usebiomass_yn = 'Y', merch_chip_nr_dpa,merch_nr_dpa) AS max_nr_dpa " + 
				                   "FROM ((((validcombos INNER JOIN " + this.m_strCondTable.Trim() + " ON validcombos.biosum_cond_id = " + this.m_strCondTable.Trim() + ".biosum_cond_id) " +  
				                                        "INNER JOIN " + this.m_strPlotTable.Trim() + " ON " + this.m_strCondTable.Trim() + ".biosum_plot_id=" + this.m_strPlotTable.Trim() + ".biosum_plot_id) " +
				                                        "INNER JOIN " + this.m_strTreeVolValSumTable.Trim() + " ON (validcombos.biosum_cond_id=" + this.m_strTreeVolValSumTable.Trim() + ".biosum_cond_id) AND " +
				                                                                                                  "(validcombos.rx=" + this.m_strTreeVolValSumTable.Trim() + ".rx)) " + 
				                                        "INNER JOIN " + this.m_strHvstCostsTable.Trim() + " ON (validcombos.biosum_cond_id=" + this.m_strHvstCostsTable.Trim() + ".biosum_cond_id) AND " +
				                                                                                              "(validcombos.rx=" + this.m_strHvstCostsTable.Trim() + ".rx)); ";
			this.m_strSQL+=strSelectSQL;
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblSumWoodProducts.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblSumWoodProducts.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblSumWoodProducts.Refresh();
				return;
			}

			if (this.UserCancel(this.m_frmRunCoreScenario.lblSumWoodProducts) == true) return;

			if (this.m_intError == 0)
			{
				this.m_frmRunCoreScenario.lblSumWoodProducts.ForeColor = System.Drawing.Color.Blue;
				this.m_frmRunCoreScenario.lblSumWoodProducts.Text = "Completed";
				this.m_frmRunCoreScenario.lblSumWoodProducts.Refresh();
			}
     

		}

		/// <summary>
		/// create a temporary work table for summing harvest costs
		/// </summary>
		private void CreateTableStructureOfHarvestCosts()
		{
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn = p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
				 **get the harvest_costs structure
				 *********************************************/
				this.m_strSQL = "SELECT biosum_cond_id,rx, complete_cpa FROM " + this.m_strHvstCostsTable.Trim() + ";";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of harvest_costs_sum
				 *****************************************************************/
				this.m_txtStreamWriter.WriteLine("Create harvest_costs_sum Table Schema From harvest_costs table");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"harvest_costs_sum",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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

				this.m_txtStreamWriter.WriteLine("Create scenario psites work table from  the scenario_psites table");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"scenario_psites_work_table",p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");

			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
				 **get the haul costs table structure
				 *********************************************/
				this.m_strSQL = "SELECT biosum_plot_id,railhead_id,psite_id,transfer_cost,road_cost,rail_cost,total_haul_cost,materialcd FROM haul_costs;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

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
				 *****************************************************************/
				this.m_txtStreamWriter.WriteLine("Create haul costs work tables from  the haul_costs table");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"all_road_merch_haul_costs_work_table",p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"all_road_chip_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cheapest_road_merch_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cheapest_road_chip_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cheapest_rail_merch_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cheapest_rail_chip_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cheapest_merch_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"cheapest_chip_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"merch_plot_to_rh_to_collector_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"chip_plot_to_rh_to_collector_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"combine_chip_rail_road_haul_costs_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"combine_merch_rail_road_haul_costs_work_table",p_dt,true);
				p_dt.Clear();

				

				/*****************************************************************
				 **create the table structures in the temp mdb file
				 **and give them the name OF merch_rh_to_collector_haul_costs_work_table
				 **                          chip_rh_to_collector_haul_costs_work_table
				 *****************************************************************/
				/*********************************************
				 **get the haul cost table structure
				 *********************************************/
				this.m_strSQL = "SELECT railhead_id,psite_id,transfer_cost,road_cost,rail_cost,total_haul_cost,materialcd FROM haul_costs;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"merch_rh_to_collector_haul_costs_work_table",p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"chip_rh_to_collector_haul_costs_work_table",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
		/// a plot's fastest travel time to a processing site
		/// </summary>
		private void CreateTableStructureForFastestTravelTimes()
		{
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");

			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
				 **get the plot table structure
				 *********************************************/
				this.m_strSQL = "SELECT biosum_plot_id, merch_haul_cost_id, merch_haul_cost_psite, merch_haul_cpa_pt, chip_haul_cost_id,chip_haul_cost_psite,chip_haul_cpa_pt FROM " + this.m_strPlotTable.Trim() + ";";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structures in the temp mdb file
				 **and give them the name OF plot_fastest_tvltm_work_table and 
				 **                          plot_fastest_tvltm_work_table2
				 *****************************************************************/
				this.m_txtStreamWriter.WriteLine("Create travel times work table (plot_fastest_tvltm_work_table) from  the plot table");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_tvltm_work_table",p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_tvltm_work_table2",p_dt,true);
				p_dt.Clear();
				
				/*********************************************
				 **get the haul cost table structure
				 *********************************************/
				this.m_strSQL = "SELECT haul_cost_id, biosum_plot_id,railhead_id,psite_id,transfer_cost,road_cost,rail_cost,total_haul_cost FROM haul_costs;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structures in the temp mdb file
				 **and give them the name OF haul_costs_work_table and 
				 **                          haul_costs_work_table2
				 *****************************************************************/
				this.m_txtStreamWriter.WriteLine("Create haul costs work table (haul_costs_work_table) from  the haul_costs table");
				
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"haul_costs_work_table",p_dt,true);
				if (p_dao.m_intError!=0)
				{
					p_dt.Dispose();
					p_ado.m_OleDbDataReader.Close();
					p_ado.m_OleDbConnection.Close();
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
					this.m_intError=p_dao.m_intError;
					p_ado=null;
					p_dao=null;
					return;
				}
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"haul_costs_work_table2",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
		/// create a temporary work tables for finding best treatments
		/// </summary>
		private void CreateTableStructureForIntensity()
		{
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");

			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
				 **get the harvest_costs structure
				 *********************************************/
				this.m_strSQL = "SELECT p.biosum_cond_id,p.rx, p.chip_yield_cf AS number_value,p.chip_yield_cf AS number_value2, i.rx_intensity AS min_intensity FROM product_yields_net_rev_costs_summary p,scenario_rx_intensity i;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of rx_intensity_work_table
				 *****************************************************************/
				this.m_txtStreamWriter.WriteLine("Create rx_intensity_work_table Schema");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"rx_intensity_duplicates_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"rx_intensity_duplicates_work_table2",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"rx_intensity_unique_work_table",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
		/// create a temporary work tables for identifying inaccessible plots and conditions
		/// </summary>
		private void CreateTableStructureForPlotCondAccessiblity()
		{
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
				this.m_txtStreamWriter.WriteLine("Create plot_cond_accessible_work_table Schema");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_cond_accessible_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"plot_cond_accessible_work_table2",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
		private void CreateTableStructureForConditionTable()
		{
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");

			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/*********************************************
				 **get fields from the plot table
				 *********************************************/
				this.m_strSQL = "SELECT * FROM " + this.m_strCondTable;

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of plot_cond_accessible_work_table
				 *****************************************************************/
				this.m_txtStreamWriter.WriteLine("Create ruledefinitionscondfilter Schema");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"ruledefinitionscondfilter",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
		private void CreateTableStructureForPSiteSum()
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
					            "FROM max_nr_plots p," + this.m_strPSiteTable.Trim() + " s;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of rx_intensity_work_table
				 *****************************************************************/
				this.m_txtStreamWriter.WriteLine("Create max_nr_plots Schema");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"psite_sum_work_table",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
		/// by land ownership
		/// </summary>
		private void CreateTableStructureForOwnershipSum()
		{
			ado_data_access p_ado = new ado_data_access();
			this.m_strConn= p_ado.getMDBConnString(this.m_strTempMDBFile,"admin","");
			
			p_ado.OpenConnection(this.m_strConn);
			if (p_ado.m_intError==0)
			{
				/**********************************************************
				 **create the max_nr_plots structure with sum_ field names
				 **********************************************************/
				this.m_strSQL = "SELECT p.owngrpcd, " +
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
					"FROM max_nr_plots p;";

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dt = p_ado.getTableSchema(p_ado.m_OleDbConnection,this.m_strSQL);

				/*****************************************************************
				 **create the table structure in the temp mdb file
				 **and give it the name of rx_intensity_work_table
				 *****************************************************************/
				this.m_txtStreamWriter.WriteLine("Create max_nr_plots Schema");
				dao_data_access p_dao = new dao_data_access();
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"own_sum_work_table",p_dt,true);
				p_dao.CreateMDBTableFromDataSetTable(this.m_strTempMDBFile,"own_sum_work_table_air_dest",p_dt,true);
				p_dt.Dispose();
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbConnection.Close();
				if (p_dao.m_intError!=0)
				{
					this.m_txtStreamWriter.WriteLine("!! Error Creating Table Schema!!");
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
		/// maximum net revenue;  maximum positive net revenue;
		/// maximum torch/crown index improvement; maximum 
		/// torch /crown index improvement with positive net revenue;
		/// minimum merchantable removal; minimum merchantable removal
		/// with positive net revenue 
		/// </summary>
		private void best_rx_summary()
		{
			this.m_frmRunCoreScenario.lblMsg.Text="Finding Best Treatments: Maximum Net Revenue";
            this.m_frmRunCoreScenario.lblMsg.Visible=true;
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_frmRunCoreScenario.progressBar1.Minimum=0;
			this.m_frmRunCoreScenario.progressBar1.Maximum = 8;
			this.m_frmRunCoreScenario.progressBar1.Value=0;
			this.m_frmRunCoreScenario.progressBar1.Visible=true;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Rx Summary");
			this.m_txtStreamWriter.WriteLine("--------------------\r\n");
			

			/*************************************************
			 **delete all records in the best_rx_summary table
			 *************************************************/
			this.m_strSQL = "delete from best_rx_summary";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/*************************************************
			 **delete all records in the best_rx_summary table
			 *************************************************/
			this.m_strSQL = "delete from best_rx_summary_air_dest";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            
			/****************************************************************
			 **delete all records in the rx_intensity_duplicates_work_table
			 ****************************************************************/
			this.m_strSQL = "delete from rx_intensity_duplicates_work_table";
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			/****************************************************************
			 **delete all records in the rx_intensity_unique_work_table
			 ****************************************************************/
			this.m_strSQL = "delete from rx_intensity_unique_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);


			/**********************************************
			 **insert unique biosum_cond_id's into the
			 **best_rx_summary table so we dont have
			 **to worry about whether the biosum_cond_id 
			 **record is in the table or not
			 **********************************************/

			this.m_strSQL = "INSERT INTO best_rx_summary " + 
				                   "SELECT DISTINCT c.biosum_cond_id,c.acres,c.owngrpcd " + 
				                   "FROM " + this.m_strCondTable.Trim() + " c, " + 
				                             this.m_strPlotTable.Trim() + " p "  + 
				                   "WHERE c.biosum_plot_id = p.biosum_plot_id  AND " + 
				                         "(p.merch_haul_cost_id IS NOT NULL OR " + 
										  "p.chip_haul_cost_id IS NOT NULL);";

			this.m_txtStreamWriter.WriteLine("insert condition records that have haul costs into best_rx_summary ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}
			/**************************************************************************************
			 **insert air destruction plots with no haul costs
			 **but they do have ci and ti treatment
			 **************************************************************************************/
			this.m_strSQL = "INSERT INTO best_rx_summary_air_dest " + 
				"SELECT DISTINCT c.biosum_cond_id,c.acres,c.owngrpcd " + 
				"FROM " + this.m_strCondTable.Trim() + " c, " + 
				this.m_strPlotTable.Trim() + " p "  + 
				"WHERE c.biosum_plot_id = p.biosum_plot_id  AND " + 
				"(p.merch_haul_cost_id IS NULL);";

			this.m_txtStreamWriter.WriteLine("insert condition records that have no MERCH haul costs into best_rx_summary_air_dest ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}



			//NET REVENUE
			this.m_txtStreamWriter.WriteLine("\r\n--NET REVENUE--");
            /*************************************************
			 **best treatment for max net revenue
			 *************************************************/

			/**********************************************************
			 **could have 2 different treatments that produce
			 **the same net revenue for a plot so insert into duplicates
			 **work table
			 ***********************************************************/
			 this.m_txtStreamWriter.WriteLine(" ");
			 this.m_txtStreamWriter.WriteLine("Max Net Revenue Processing");
             this.m_txtStreamWriter.WriteLine("                           ");

			if (m_ado.TableExist(this.m_TempMDBFileConn,"effective_product_yields_net_rev_costs_summary"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE effective_product_yields_net_rev_costs_summary");
			}

			//write effective treatments to the effective_product_yields_net_rev_costs_summary
			m_ado.m_strSQL = "SELECT p.*,i.rx_intensity " + 
				"INTO effective_product_yields_net_rev_costs_summary " + 
				"FROM product_yields_net_rev_costs_summary p,effective e, scenario_rx_intensity i " + 
				"WHERE p.biosum_cond_id = e.biosum_cond_id AND " + 
					  "p.rx=e.rx  AND e.effective_yn='Y' AND i.rx=p.rx AND " + 
				      "trim(ucase(i.scenario_id))='" + this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";

			this.m_txtStreamWriter.WriteLine("write effective to the effective_product_yields_net_rev_costs_summary ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}


			if (m_ado.TableExist(this.m_TempMDBFileConn,"max_revenue_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE max_revenue_min_intensity_work_table");
				
			}
           

			m_ado.m_strSQL="CREATE TABLE max_revenue_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(3)," + 
				"max_nr_dpa DOUBLE," + 
				"rx_intensity SHORT)";

			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			
			m_ado.m_strSQL = "INSERT INTO max_revenue_min_intensity_work_table " + 
                "SELECT a.biosum_cond_id,a.rx,a.max_nr_dpa,a.rx_intensity " + 
				"FROM effective_product_yields_net_rev_costs_summary a, " + 
				"(SELECT biosum_cond_id,MAX(max_nr_dpa) AS max_revenue " + 
				"FROM  effective_product_yields_net_rev_costs_summary " + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id,  max_nr_dpa, MIN(rx_intensity) AS min_intensity  " + 
				"FROM effective_product_yields_net_rev_costs_summary " + 
				"GROUP BY biosum_cond_id,max_nr_dpa) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity AND " + 
				"b.max_revenue=a.max_nr_dpa";
		
			this.m_txtStreamWriter.WriteLine("Insert Rx With Maximum Net Revenue ");
			this.m_txtStreamWriter.WriteLine("get treatment with maximum net revenue at minimal intensity to max_revenue_min_intensity_work_table");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}


            
			/*******************************************************
			 **update the best_rx_summary from the 
			 **max_revenue_min_intensity_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				            "INNER JOIN max_revenue_min_intensity_work_table source " +
				                "ON target.biosum_cond_id = source.biosum_cond_id " +
				            "SET max_nr_rx = source.rx, max_nr_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with best net revenue ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}


			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}
			this.m_frmRunCoreScenario.progressBar1.Value=1;
            /****************************************
			 **finished with max net revenue
			 ****************************************/

			/****************************************
			 **start max positive net revenue
			 ****************************************/

            
            this.m_frmRunCoreScenario.lblMsg.Text="Finding Best Treatments: Maximum Positive Net Revenue";
			this.m_frmRunCoreScenario.lblMsg.Refresh();


			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Max Positive Net Revenue Processing");
			this.m_txtStreamWriter.WriteLine(" ");

			if (m_ado.TableExist(this.m_TempMDBFileConn,"positive_max_revenue_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE positive_max_revenue_min_intensity_work_table");
			}
			m_ado.m_strSQL="CREATE TABLE positive_max_revenue_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"max_nr_dpa DOUBLE," + 
				"rx_intensity SHORT)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			m_ado.m_strSQL = "INSERT INTO positive_max_revenue_min_intensity_work_table " + 
                "SELECT a.biosum_cond_id,a.rx,a.max_nr_dpa,a.rx_intensity " + 
				"FROM effective_product_yields_net_rev_costs_summary a, " + 
				"(SELECT biosum_cond_id,MAX(max_nr_dpa) AS max_revenue " + 
				"FROM  effective_product_yields_net_rev_costs_summary WHERE max_nr_dpa > 0 " + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id,  max_nr_dpa, MIN(rx_intensity) AS min_intensity  " + 
				"FROM effective_product_yields_net_rev_costs_summary " + 
				"GROUP BY biosum_cond_id,max_nr_dpa) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity AND " + 
				"b.max_revenue=a.max_nr_dpa";

			this.m_txtStreamWriter.WriteLine("Insert Rx With Maximum Positive Net Revenue ");
			this.m_txtStreamWriter.WriteLine("get treatment with positive maximum net revenue at minimal intensity to max_revenue_min_intensity_work_table");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}



			/*******************************************************
			 **update the best_rx_summary from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN positive_max_revenue_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_pnr_rx = source.rx, max_pnr_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with best net revenue ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}

            this.m_frmRunCoreScenario.progressBar1.Value=2;
			/****************************************
			 **finished with max positive net revenue
			 ****************************************/

			//TORCH INDEX
			this.m_txtStreamWriter.WriteLine("\r\n--TORCH INDEX--");
			/****************************************
			 **start maximum torch index improvement
			 ****************************************/
			
			this.m_frmRunCoreScenario.lblMsg.Text="Finding Best Treatments: Maximum Torch Index Improvement";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			


			
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Max Torch Index Improvement Processing");
			this.m_txtStreamWriter.WriteLine(" ");

			if (m_ado.TableExist(this.m_TempMDBFileConn,"effective_tici_change_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE effective_tici_change_work_table");
			}
			m_ado.m_strSQL="SELECT e.biosum_cond_id, e.rx, i.rx_intensity,e.ti_change,e.ci_change,p.max_nr_dpa " + 
				"INTO effective_tici_change_work_table  " + 
				"FROM effective e, scenario_rx_intensity i,effective_product_yields_net_rev_costs_summary p " + 
				"WHERE e.biosum_cond_id=p.biosum_cond_id AND " +
				      "e.rx=p.rx AND " + 
					  "i.rx=p.rx AND " + 
				"TRIM(UCASE(i.scenario_id)) ='" + this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioId.Text.Trim().ToUpper() + "';";
			this.m_txtStreamWriter.WriteLine("write effective torch index treatments to effective_tici_change_work_table ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}
			//get max torch index improvement with the least intensive treatment
			if (m_ado.TableExist(this.m_TempMDBFileConn,"max_ti_change_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE max_ti_change_min_intensity_work_table");
			}
			m_ado.m_strSQL="CREATE TABLE max_ti_change_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"ti_change DOUBLE," + 
				"rx_intensity SHORT)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);
			m_ado.m_strSQL="INSERT INTO max_ti_change_min_intensity_work_table " + 
                           "SELECT a.biosum_cond_id,a.rx,a.ti_change,a.rx_intensity " + 
				           "FROM effective_tici_change_work_table a," + 
								"(SELECT biosum_cond_id,MAX(ti_change) AS max_ti_change  " + 
				                 "FROM  effective_tici_change_work_table " + 
				                 "GROUP BY biosum_cond_id) b," + 
									"(SELECT biosum_cond_id,  ti_change, MIN(rx_intensity) AS min_intensity " + 
				                     "FROM effective_tici_change_work_table " + 
				                     "GROUP BY biosum_cond_id,ti_change) c " + 
						  "WHERE b.biosum_cond_id=a.biosum_cond_id  AND " + 
				                "c.biosum_cond_id=a.biosum_cond_id AND " + 
				                "c.min_intensity=a.rx_intensity AND " + 
				                "b.max_ti_change=a.ti_change";

			this.m_txtStreamWriter.WriteLine("Insert Rx With Maximum Torch Index Improvement ");
			this.m_txtStreamWriter.WriteLine("get treatment with the greatest torch index improvement at minimal intensity to max_ti_change_min_intensity_work_table ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}


			/*******************************************************
			 **update the best_rx_summary from the 
			 **max_ti_change_min_intensity_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN max_ti_change_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_ti_imp_rx = source.rx, max_ti_imp_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with most torch index improvement");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}
			/*******************************************************
			 **update the best_rx_summary_air_dest from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary_air_dest target " + 
				"INNER JOIN max_ti_change_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_ti_imp_rx = source.rx, max_ti_imp_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary_air_dest with most torch index improvement");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}

    		this.m_frmRunCoreScenario.progressBar1.Value=3;
			/***********************************************
			 **finished with maximum torch index improvement
			 ***********************************************/


			/*******************************************************************
			 **start maximum torch index improvement with positive net revenue
			 *******************************************************************/
			
			this.m_frmRunCoreScenario.lblMsg.Text="Finding Best Treatments: Maximum Torch Index Improvement With Positive Net Revenue";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			


			/**********************************************************
			 **get all records that have positive net revenue
			 **and torch index improvement
			 ***********************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Max Torch Index Improvement With Positive Net Revenue Processing");
			this.m_txtStreamWriter.WriteLine(" ");

			//get max torch index improvement with positive net revenue and the least intensive treatment
			if (m_ado.TableExist(this.m_TempMDBFileConn,"max_ti_change_pnr_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE max_ti_change_pnr_min_intensity_work_table");
			}
			m_ado.m_strSQL="CREATE TABLE max_ti_change_pnr_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"ti_change DOUBLE," + 
				"rx_intensity SHORT," + 
				"max_nr_dpa DOUBLE)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			m_ado.m_strSQL="INSERT INTO max_ti_change_pnr_min_intensity_work_table " + 
                           "SELECT a.biosum_cond_id,a.rx,a.ti_change,a.rx_intensity,a.max_nr_dpa " + 
				"FROM effective_tici_change_work_table a," + 
				"(SELECT biosum_cond_id,MAX(ti_change) AS max_ti_change  " + 
				"FROM  effective_tici_change_work_table WHERE max_nr_dpa > 0 "  + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id,  ti_change, MIN(rx_intensity) AS min_intensity " + 
				"FROM effective_tici_change_work_table " + 
				"GROUP BY biosum_cond_id,ti_change) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id  AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity AND " + 
				"b.max_ti_change=a.ti_change";

			this.m_txtStreamWriter.WriteLine("Insert Rx With Maximum Torch Index Improvement And Positive Net Revenue ");
			this.m_txtStreamWriter.WriteLine("get treatment with the greatest torch index improvement and minimal intensity that has a positive net revenue to max_ti_change_pnr_min_intensity_work_table ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}

			/*******************************************************
			 **update the best_rx_summary from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN max_ti_change_pnr_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_ti_imp_pnr_rx = source.rx, max_ti_imp_pnr_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with most torch index improvement that has positive net revenue");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}


			this.m_frmRunCoreScenario.progressBar1.Value=4;
			/*************************************************************************
			 **finished with maximum torch index improvement with positive net revenue
			 *************************************************************************/

			//CROWN INDEX
			this.m_txtStreamWriter.WriteLine("\r\n--CROWN INDEX--");
			/****************************************
			 **start maximum crown index improvement
			 ****************************************/
			this.m_frmRunCoreScenario.lblMsg.Text="Finding Best Treatments: Maximum Crown Index Improvement";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			

			/**********************************************************
			 **could have 2 different treatments that produce
			 **the same net revenue for a plot so insert into duplicates
			 **work table
			 ***********************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Max Crown Index Improvement Processing");
			this.m_txtStreamWriter.WriteLine(" ");

			//get max crown index improvement with the least intensive treatment
			if (m_ado.TableExist(this.m_TempMDBFileConn,"max_ci_change_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE max_ci_change_min_intensity_work_table");
			}
			m_ado.m_strSQL="CREATE TABLE max_ci_change_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"ci_change DOUBLE," + 
				"rx_intensity SHORT)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			m_ado.m_strSQL="INSERT INTO max_ci_change_min_intensity_work_table " + 
                "SELECT a.biosum_cond_id,a.rx,a.ci_change,a.rx_intensity " + 
				"FROM effective_tici_change_work_table a," + 
				"(SELECT biosum_cond_id,MAX(ci_change) AS max_ci_change  " + 
				"FROM  effective_tici_change_work_table " + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id,  ci_change, MIN(rx_intensity) AS min_intensity " + 
				"FROM effective_tici_change_work_table " + 
				"GROUP BY biosum_cond_id,ci_change) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id  AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity AND " + 
				"b.max_ci_change=a.ci_change";

			this.m_txtStreamWriter.WriteLine("Insert Rx With Maximum Crown Index Improvement ");
			this.m_txtStreamWriter.WriteLine("get treatment with the greatest crown index improvement at minimal intensity to max_ci_change_min_intensity_work_table ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}

			/*******************************************************
			 **update the best_rx_summary from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN max_ci_change_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_ci_imp_rx = source.rx, max_ci_imp_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with most crown index improvement");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}
			/*******************************************************
			 **update the best_rx_summary_air_dest from the 
			 **max_ci_change_min_intensity_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary_air_dest target " + 
				"INNER JOIN max_ci_change_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_ci_imp_rx = source.rx, max_ci_imp_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary_air_dest with most crown index improvement");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}
			this.m_frmRunCoreScenario.progressBar1.Value=5;
			/***********************************************
			 **finished with maximum crown index improvement
			 ***********************************************/
			/*******************************************************************
			 **start maximum crown index improvement with positive net revenue
			 *******************************************************************/
			
			this.m_frmRunCoreScenario.lblMsg.Text="Finding Best Treatments: Maximum Crown Index Improvement With Positive Net Revenue";
			this.m_frmRunCoreScenario.lblMsg.Refresh();


			/**********************************************************
			 **get all records that have positive net revenue
			 **and most crown index improvement
			 ***********************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Max Crown Index Improvement With Positive Net Revenue Processing");
			this.m_txtStreamWriter.WriteLine(" ");
			//get max torch index improvement with positive net revenue and the least intensive treatment
			if (m_ado.TableExist(this.m_TempMDBFileConn,"max_ci_change_pnr_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE max_ci_change_pnr_min_intensity_work_table");
			}
			m_ado.m_strSQL="CREATE TABLE max_ci_change_pnr_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"ci_change DOUBLE," + 
				"rx_intensity SHORT," + 
				"max_nr_dpa DOUBLE)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			m_ado.m_strSQL="INSERT INTO max_ci_change_pnr_min_intensity_work_table " + 
                "SELECT a.biosum_cond_id,a.rx,a.ci_change,a.rx_intensity,a.max_nr_dpa " + 
				"FROM effective_tici_change_work_table a," + 
				"(SELECT biosum_cond_id,MAX(ci_change) AS max_ci_change  " + 
				"FROM  effective_tici_change_work_table WHERE max_nr_dpa > 0 " + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id,  ci_change, MIN(rx_intensity) AS min_intensity " + 
				"FROM effective_tici_change_work_table " + 
				"GROUP BY biosum_cond_id,ci_change) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id  AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity AND " + 
				"b.max_ci_change=a.ci_change";

			this.m_txtStreamWriter.WriteLine("Insert Rx With Maximum Crown Index Improvement And Positive Net Revenue ");
			this.m_txtStreamWriter.WriteLine("get treatment with the greatest crown index improvement and minimal intensity that has a positive net revenue to max_ci_change_pnr_min_intensity_work_table ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}


			/*******************************************************
			 **update the best_rx_summary from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN max_ci_change_pnr_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET max_ci_imp_pnr_rx = source.rx, max_ci_imp_pnr_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with most crown index improvement that has positive net revenue");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}

			this.m_frmRunCoreScenario.progressBar1.Value=6;
			/*************************************************************************
			 **finished with maximum crown index improvement with positive net revenue
			 *************************************************************************/

			//MERCH
			this.m_txtStreamWriter.WriteLine("\r\n--MERCH--");
			/******************************************
			 **start minimum merchantable wood removal
			 ******************************************/
			this.m_frmRunCoreScenario.lblMsg.Text="Finding Best Treatments: Minimum Merchantable Wood Removal";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine("Minimum Merchantable Wood Removable Processing");
			this.m_txtStreamWriter.WriteLine(" ");
			if (m_ado.TableExist(this.m_TempMDBFileConn,"min_merch_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE min_merch_min_intensity_work_table");
			}

			m_ado.m_strSQL="CREATE TABLE min_merch_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"merch_yield_cf DOUBLE," + 
				"rx_intensity SHORT)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			m_ado.m_strSQL = "INSERT INTO min_merch_min_intensity_work_table " + 
				"SELECT a.biosum_cond_id,a.rx,a.merch_yield_cf,a.rx_intensity " + 
				"FROM effective_product_yields_net_rev_costs_summary a, " + 
				"(SELECT biosum_cond_id,MIN(merch_yield_cf) AS min_merch_removal " + 
				"FROM  effective_product_yields_net_rev_costs_summary " + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id, merch_yield_cf, MIN(rx_intensity) AS min_intensity  " + 
				"FROM effective_product_yields_net_rev_costs_summary " + 
				"GROUP BY biosum_cond_id,merch_yield_cf) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id AND " + 
				"b.min_merch_removal=a.merch_yield_cf AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity";
				

			this.m_txtStreamWriter.WriteLine("Insert Rx With Minimum Merch Removal ");
			this.m_txtStreamWriter.WriteLine("get treatment with minimum merch removal at minimal intensity to min_merch_min_intensity_work_table");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}


			/*******************************************************
			 **update the best_rx_summary from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN min_merch_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET min_merch_rx = source.rx, min_merch_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with minimum merch removal");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}
			/*******************************************************
			 **update the best_rx_summary_air_dest from the 
			 **min_merch_min_intensity_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary_air_dest target " + 
				"INNER JOIN min_merch_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET min_merch_rx = source.rx, min_merch_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary_air_dest with minimum merch removal");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}
			this.m_frmRunCoreScenario.progressBar1.Value=7;
			/**************************************************
			 **finished with minimum merchantable wood removal 
			 **************************************************/
			/*******************************************************************
			 **start minimum merchantable wood removal with positive net revenue
			 *******************************************************************/
			
			this.m_frmRunCoreScenario.lblMsg.Text="Finding Best Treatments: Minimum Merchantable Wood Removal With Positive Net Revenue";
			this.m_frmRunCoreScenario.lblMsg.Refresh();

			/**********************************************************
			 **get all records that have positive net revenue
			 **and minimum merchantable wood removal
			 ***********************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Minimum Merchantable Wood Removal With Positive Net Revenue Processing");
			this.m_txtStreamWriter.WriteLine(" ");
			if (m_ado.TableExist(this.m_TempMDBFileConn,"min_merch_pnr_min_intensity_work_table"))
			{
				m_ado.SqlNonQuery(m_TempMDBFileConn,"DROP TABLE min_merch_pnr_min_intensity_work_table");
			}
			m_ado.m_strSQL="CREATE TABLE min_merch_pnr_min_intensity_work_table " + 
				"(biosum_cond_id CHAR(25)," + 
				"rx CHAR(1)," + 
				"merch_yield_cf DOUBLE," + 
				"rx_intensity SHORT," + 
				"max_nr_dpa DOUBLE)";
			m_ado.SqlNonQuery(m_TempMDBFileConn,m_ado.m_strSQL);

			m_ado.m_strSQL = "INSERT INTO min_merch_pnr_min_intensity_work_table " + 
                "SELECT a.biosum_cond_id,a.rx,a.merch_yield_cf,a.rx_intensity,a.max_nr_dpa " + 
				"FROM effective_product_yields_net_rev_costs_summary a, " + 
				"(SELECT biosum_cond_id,MIN(merch_yield_cf) AS min_merch_removal " + 
				"FROM  effective_product_yields_net_rev_costs_summary WHERE max_nr_dpa > 0 " + 
				"GROUP BY biosum_cond_id) b," + 
				"(SELECT biosum_cond_id, merch_yield_cf, MIN(rx_intensity) AS min_intensity  " + 
				"FROM effective_product_yields_net_rev_costs_summary " + 
				"GROUP BY biosum_cond_id,merch_yield_cf) c " + 
				"WHERE b.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.biosum_cond_id=a.biosum_cond_id AND " + 
				"c.min_intensity=a.rx_intensity AND " + 
				"b.min_merch_removal=a.merch_yield_cf";

			this.m_txtStreamWriter.WriteLine("Insert Rx With Minimum Merch Removal And Positive Net Revenue ");
			this.m_txtStreamWriter.WriteLine("get treatment with minimum merch removal and positive net revenue at minimal intensity to max_ti_change_pnr_min_intensity_work_table ");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + m_ado.m_strSQL);
			m_ado.SqlNonQuery(this.m_TempMDBFileConn,m_ado.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}

			/*******************************************************
			 **update the best_rx_summary from the 
			 **rx_intensity_unique_work_table
			 *******************************************************/
			this.m_strSQL = "UPDATE best_rx_summary target " + 
				"INNER JOIN min_merch_pnr_min_intensity_work_table source " +
				"ON target.biosum_cond_id = source.biosum_cond_id " +
				"SET min_merch_pnr_rx = source.rx, min_merch_pnr_rxint = source.rx_intensity;";
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Update best_rx_summary with minimum merchantable wood removal that has positive net revenue");
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRx) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
				return;
			}

			this.m_frmRunCoreScenario.progressBar1.Value=8;
			/****************************************************************************
			 **finished with minimum merchantable wood removal with positive net revenue
			 ****************************************************************************/


			if (this.m_intError == 0)
			{
				this.m_frmRunCoreScenario.lblProcBestRx.ForeColor = System.Drawing.Color.Blue;
				this.m_frmRunCoreScenario.lblProcBestRx.Text = "Completed";
				this.m_frmRunCoreScenario.lblProcBestRx.Refresh();
			}

		}

		/// <summary>
		/// expand the wood volume,value,and costs by plot acreage for the 
		/// best treatments and data found in the best_rx_summary
		/// product_yields_net_rev_costs_summary, and effective tables
		/// </summary>
		private void BestTreatmentsByPlot()
		{
			this.m_frmRunCoreScenario.lblMsg.Text="Best Treatment Acreage Expansion By Plot";
			this.m_frmRunCoreScenario.lblMsg.Visible=true;
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_frmRunCoreScenario.progressBar1.Minimum=0;
			this.m_frmRunCoreScenario.progressBar1.Maximum = 8;
			this.m_frmRunCoreScenario.progressBar1.Value=0;
			this.m_frmRunCoreScenario.progressBar1.Visible=true;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Treatment Acreage Expansion By Plot");
			this.m_txtStreamWriter.WriteLine("-----------------------------------------");

			            
			/*************************************************
			 **best maximum net revenue by plot
			 *************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Net Revenue By Plot");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine("Insert net revenue totals by plot");
			this.BestRxAcreageExpansionTableInsert("max_nr_plots","biosum_cond_id","max_nr_rx","");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPlot) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Refresh();
				return;
			}

			this.m_frmRunCoreScenario.progressBar1.Value=1;


			/*************************************************
			 **Finished with best maximum net revenue by plot
			 *************************************************/

			/*************************************************
			 **best positive net revenue by plot
			 *************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Positive Net Revenue By Plot");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine("Insert positive net revenue totals by plot");
			this.BestRxAcreageExpansionTableInsert("max_pnr_plots","biosum_cond_id","max_pnr_rx","effective_product_yields_net_rev_costs_summary.max_nr_dpa * best_rx_summary.acres >=  0");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPlot) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Refresh();
				return;
			}

			this.m_frmRunCoreScenario.progressBar1.Value=2;
			/*************************************************
			 **Finished with best maximum net revenue by plot
			 *************************************************/

			/*************************************************
			 **best maximum torch index improvement by plot
			 *************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Torch Index Improvement By Plot");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert torch index improvement by plot");
			this.BestRxAcreageExpansionTableInsert("max_ti_imp_plots","biosum_cond_id","max_ti_imp_rx","");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPlot) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Refresh();
				return;
			}

			/**********************************************************************************
			 **best maximum torch index improvement by plot for air curtain destruction plots
			 **********************************************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Torch Index Improvement By Plot For Air Curtain Destruction Plots");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert torch index improvement by air curtain destruction plot");
			this.BestRxAcreageExpansionTableInsertForAirCurtainDestruction("max_ti_imp_plots_air_dest","biosum_cond_id","max_ti_imp_rx","");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPlot) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Refresh();
				return;
			}




			this.m_frmRunCoreScenario.progressBar1.Value=3;
			/****************************************************
			 **Finished with best torch index improvement by plot
			 ****************************************************/


			/*************************************************************************
			 **best maximum torch index improvement with positive net revenue by plot
			 *************************************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Torch Index Improvement With Positive Net Revenue By Plot");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert Torch Index Improvement With Positive Net Revenue By plot");
			this.BestRxAcreageExpansionTableInsert("max_ti_imp_pnr_plots","biosum_cond_id","max_ti_imp_pnr_rx","effective_product_yields_net_rev_costs_summary.max_nr_dpa * best_rx_summary.acres >=  0");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPlot) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Refresh();
				return;
			}

			this.m_frmRunCoreScenario.progressBar1.Value=4;
			/*******************************************************************************
			 **Finished with best torch index improvement with positive net revenue by plot
			 *******************************************************************************/
            
			/*************************************************
			 **best maximum crown index improvement by plot
			 *************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Crown Index Improvement By Plot");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert Crown index improvement by plot");
			this.BestRxAcreageExpansionTableInsert("max_ci_imp_plots","biosum_cond_id","max_ci_imp_rx","");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPlot) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Refresh();
				return;
			}

			/**********************************************************************************
			 **best maximum crown index improvement by plot for air curtain destruction plots
			 **********************************************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Crown Index Improvement By Plot For Air Curtain Destruction Plots");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert crown index improvement by air curtain destruction plot");
			this.BestRxAcreageExpansionTableInsertForAirCurtainDestruction("max_ci_imp_plots_air_dest","biosum_cond_id","max_ci_imp_rx","");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPlot) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Refresh();
				return;
			}

			this.m_frmRunCoreScenario.progressBar1.Value=5;
			/****************************************************
			 **Finished with best crown index improvement by plot
			 ****************************************************/


			/*************************************************************************
			 **best maximum crown index improvement with positive net revenue by plot
			 *************************************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Crown Index Improvement With Positive Net Revenue By Plot");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert Crown Index Improvement With Positive Net Revenue By plot");
			this.BestRxAcreageExpansionTableInsert("max_ci_imp_pnr_plots","biosum_cond_id","max_ci_imp_pnr_rx","effective_product_yields_net_rev_costs_summary.max_nr_dpa * best_rx_summary.acres >=  0");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPlot) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Refresh();
				return;
			}

			this.m_frmRunCoreScenario.progressBar1.Value=6;
			/*******************************************************************************
			 **Finished with best crown index improvement with positive net revenue by plot
			 *******************************************************************************/

			/*************************************************
			 **best minimum merchantable wood removal by plot
			 *************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Minimum Merchantable Wood Removal By Plot");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert records");
			this.BestRxAcreageExpansionTableInsert("min_merch_plots","biosum_cond_id","min_merch_rx","");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPlot) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Refresh();
				return;
			}

			/*************************************************************************************
			 **best minimum merchantable wood removal by plot for air curtain destruction plots
			 *************************************************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Minimum Merchantable Wood Removal By Plot For Air Curtain Destruction Plots");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert records");
			this.BestRxAcreageExpansionTableInsertForAirCurtainDestruction("min_merch_plots_air_dest","biosum_cond_id","min_merch_rx","");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPlot) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Refresh();
				return;
			}

			this.m_frmRunCoreScenario.progressBar1.Value=7;
			/****************************************************
			 **Finished with best crown index improvement by plot
			 ****************************************************/


			/*************************************************************************
			 **best maximum crown index improvement with positive net revenue by plot
			 *************************************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Best Minimum Merchantable Wood Removal With Positive Net Revenue By Plot");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert Records");
			this.BestRxAcreageExpansionTableInsert("min_merch_pnr_plots","biosum_cond_id","min_merch_pnr_rx","effective_product_yields_net_rev_costs_summary.max_nr_dpa * best_rx_summary.acres >=  0");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPlot) == true) return;

			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Refresh();
				return;
			}

			this.m_frmRunCoreScenario.progressBar1.Value=8;
			/*******************************************************************************
			 **Finished with best crown index improvement with positive net revenue by plot
			 *******************************************************************************/




			if (this.m_intError == 0)
			{
				this.m_frmRunCoreScenario.lblProcBestRxPlot.ForeColor = System.Drawing.Color.Blue;
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Text = "Completed";
				this.m_frmRunCoreScenario.lblProcBestRxPlot.Refresh();
			}

		}
		private void DeleteScenarioResultRecords()
		{

			string strMDBPathAndFile = this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\scenario_results.mdb";
			ado_data_access oAdo = new ado_data_access();
			string strConn=oAdo.getMDBConnString(strMDBPathAndFile,"admin","");
			oAdo.OpenConnection(strConn);

			/*************************************************
			 **delete all records in the by plot tables
			 *************************************************/
			this.m_strSQL = "delete from max_nr_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_pnr_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ti_imp_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ti_imp_pnr_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ci_imp_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ci_imp_pnr_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from min_merch_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from min_merch_pnr_plots";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);

			/*************************************************
			 **delete all records in the by ownership tables
			 *************************************************/
			this.m_strSQL = "delete from max_nr_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_pnr_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ti_imp_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ti_imp_pnr_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ci_imp_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ci_imp_pnr_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from min_merch_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from min_merch_pnr_sum_own";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);

			/******************************************************
			 **delete all records in the by processing site tables
			 ******************************************************/
			this.m_strSQL = "delete from max_nr_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_pnr_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ti_imp_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ti_imp_pnr_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ci_imp_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from max_ci_imp_pnr_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from min_merch_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);
			this.m_strSQL = "delete from min_merch_pnr_sum_psite";
			oAdo.SqlNonQuery(oAdo.m_OleDbConnection,this.m_strSQL);

			oAdo.m_OleDbConnection.Close();
			while (oAdo.m_OleDbConnection.State==System.Data.ConnectionState.Open)
				System.Threading.Thread.Sleep(1000);

			oAdo.m_OleDbConnection.Dispose();
			oAdo=null;



		}
		private void BestRxAcreageExpansionTableInsert(string strTable,string strTypeField, string strRxField, string strWhereExpression)
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
				"best_rx_summary.biosum_cond_id," + 
				"best_rx_summary.owngrpcd," + 
				"best_rx_summary.acres," +
				p + ".merch_haul_cost_psite," + 
				p + ".merch_haul_cpa_pt * best_rx_summary.acres AS merch_haul_cents," + 
				"effective_product_yields_net_rev_costs_summary.merch_yield_cf * best_rx_summary.acres AS merch_vol," + 
				"effective_product_yields_net_rev_costs_summary.merch_val_dpa * best_rx_summary.acres AS merch_dollars_val," + 
				p + ".chip_haul_cost_psite," + 
				p + ".chip_haul_cpa_pt * best_rx_summary.acres AS chip_haul_cents," + 
				"effective_product_yields_net_rev_costs_summary.chip_yield_gt * best_rx_summary.acres AS chip_yield," + 
				"effective_product_yields_net_rev_costs_summary.chip_val_dpa * best_rx_summary.acres AS chip_dollars_val," + 
				"effective_product_yields_net_rev_costs_summary.max_nr_dpa * best_rx_summary.acres AS net_rev," + 
				"effective_product_yields_net_rev_costs_summary.harvest_onsite_cpa * best_rx_summary.acres AS harv_costs," + 
				"(effective_product_yields_net_rev_costs_summary.haul_merch_cpa + effective_product_yields_net_rev_costs_summary.haul_chip_cpa) * best_rx_summary.acres AS haul_costs," + 
				"effective.ti_change * best_rx_summary.acres  AS ti_chg_acres," +
				"effective.ci_change * best_rx_summary.acres  AS ci_chg_acres " + 
				"FROM (((best_rx_summary  " + 
				"INNER JOIN effective_product_yields_net_rev_costs_summary  " + 
				"ON (best_rx_summary.biosum_cond_id = effective_product_yields_net_rev_costs_summary.biosum_cond_id) AND " + 
				"(best_rx_summary." + strRxField + " = effective_product_yields_net_rev_costs_summary.rx)) " + 
				"INNER JOIN effective " + 
				"ON (best_rx_summary.biosum_cond_id = effective.biosum_cond_id) AND " +
				"(best_rx_summary." + strRxField + "= effective.rx)) " + 
				"INNER JOIN (" + p + 
				" INNER JOIN " + c + 
				" ON " + p + ".biosum_plot_id = " + c + ".biosum_plot_id) " + 
				"ON " + c + ".biosum_cond_id = best_rx_summary.biosum_cond_id)" ;
				
			if (strWhereExpression.Trim().Length > 0)
			{
				this.m_strSQL += " WHERE " + strWhereExpression + ";";
			}
			else
			{
				this.m_strSQL += ";";
			}

			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}

		private void BestRxAcreageExpansionTableInsertForAirCurtainDestruction(string strTable,string strTypeField, string strRxField, string strWhereExpression)
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

			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}


		/// <summary>
		/// sum wood volume, values, and costs by processing sites
		/// </summary>
		private void SumPSite()
		{
			//int x=0;
			this.m_frmRunCoreScenario.lblMsg.Text="Maximum Net Revenue By Processing Site";
			this.m_frmRunCoreScenario.lblMsg.Visible=true;
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_frmRunCoreScenario.progressBar1.Minimum=0;
			this.m_frmRunCoreScenario.progressBar1.Maximum = 8;
			this.m_frmRunCoreScenario.progressBar1.Value=0;
			this.m_frmRunCoreScenario.progressBar1.Visible=true;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Processing Site Summary");
			this.m_txtStreamWriter.WriteLine("-----------------------");

			/*************************************************
			 **delete all records in the by plot tables
			 *************************************************/
			
			this.m_strSQL = "delete from psite_sum_work_table";
            this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            
			//MAX_NR_SUM_PSITE
			/*************************************************
			 **net revenue by processing site
			 *************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Net Revenue By Psite");
			this.m_txtStreamWriter.WriteLine(" ");
			/*************************************************
			 **process Chip only processing sites
			 *************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert chip only psite sums");
			this.SumPSiteWorkTableInsert("max_nr_plots","chip_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.chip_haul_cost_psite AND a.biocd=2");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************************
			 **process merchantable only processing sites
			 *************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch only psite sums");
			this.SumPSiteWorkTableInsert("max_nr_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=1");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the merch_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from merch_haul_cost_psite column");
			this.SumPSiteWorkTableInsert("max_nr_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=3");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/**********************************************************
			 **get a list of all the psites currently processed
			 **********************************************************/
			string strPSiteList = this.m_ado.CreateSQLNOTINString(this.m_TempMDBFileConn,"select psite_id from psite_sum_work_table");
			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the chip_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from chip_haul_cost_psite column");
			if (strPSiteList.Trim().Length > 0)
			{
			   this.SumPSiteWorkTableInsert("max_nr_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				      "a.psite_id=c.chip_haul_cost_psite AND " + 
				      "a.biocd=3 AND " + 
				      "b.cheapest_psite NOT IN (" + strPSiteList.Trim() + ")");
			}
			else
			{
				  this.SumPSiteWorkTableInsert("max_nr_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				      "a.psite_id=c.chip_haul_cost_psite AND " + 
				      "a.biocd=3");

			}
				    
      if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_nr_sum_psite
			 *****************************************************************/
            this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert from psite values from work table to max_nr_sum_psite");
			this.SumPSiteTableInsert("max_nr_sum_psite");
			this.m_frmRunCoreScenario.progressBar1.Value=1;
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************
			 **Finished net revenue by processsing site
			 *************************************************/

			//MAX_PNR_SUM_PSITE
			/*************************************************
			 **Positive Net Revenue By Processing site
			 *************************************************/
			this.m_strSQL = "delete from psite_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            

            this.m_frmRunCoreScenario.lblMsg.Text="Maximum Positive Net Revenue By Processing Site";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Positive Net Revenue By Psite");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert chip only psite sums");
			this.SumPSiteWorkTableInsert("max_pnr_plots","chip_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.chip_haul_cost_psite AND a.biocd=2");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************************
			 **process merchantable only processing sites
			 *************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch only psite sums");
			this.SumPSiteWorkTableInsert("max_pnr_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=1");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the merch_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from merch_haul_cost_psite column");
			this.SumPSiteWorkTableInsert("max_pnr_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=3");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/**********************************************************
			 **get a list of all the psites currently processed
			 **********************************************************/
			strPSiteList = this.m_ado.CreateSQLNOTINString(this.m_TempMDBFileConn,"select psite_id from psite_sum_work_table");
			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the chip_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from chip_haul_cost_psite column");
			if (strPSiteList.Trim().Length > 0)
			{
			   this.SumPSiteWorkTableInsert("max_pnr_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				   "a.psite_id=c.chip_haul_cost_psite AND " + 
				   "a.biocd=3 AND " + 
				   "b.cheapest_psite NOT IN (" + strPSiteList.Trim() + ")");
		  }
		  else
		  {
		  	this.SumPSiteWorkTableInsert("max_pnr_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				   "a.psite_id=c.chip_haul_cost_psite AND " + 
				   "a.biocd=3");
		  }
      if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_pnr_sum_psite
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert from psite values from work table to max_pnr_sum_psite");
			this.SumPSiteTableInsert("max_pnr_sum_psite");
			this.m_frmRunCoreScenario.progressBar1.Value=2;
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************
			 **Finished positive net revenue by processsing site
			 *************************************************/

            //MAX_TI_IMP_SUM_PSITE
			/*************************************************
			 **most ti improvement by processing site
			 *************************************************/
			this.m_strSQL = "delete from psite_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_frmRunCoreScenario.lblMsg.Text="Maximum Torch Index Improvement By Processing Site";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Torch Index Improvement By Psite");
			this.m_txtStreamWriter.WriteLine(" ");
			/*************************************************
			 **process chip only processing sites
			 *************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert chip only psite sums");
			this.SumPSiteWorkTableInsert("max_ti_imp_plots","chip_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.chip_haul_cost_psite AND a.biocd=2");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************************
			 **process merchantable only processing sites
			 *************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch only psite sums");
			this.SumPSiteWorkTableInsert("max_ti_imp_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=1");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the merch_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from merch_haul_cost_psite column");
			this.SumPSiteWorkTableInsert("max_ti_imp_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=3");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/**********************************************************
			 **get a list of all the psites currently processed
			 **********************************************************/
			strPSiteList = this.m_ado.CreateSQLNOTINString(this.m_TempMDBFileConn,"select psite_id from psite_sum_work_table");
			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the chip_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from chip_haul_cost_psite column");
			if (strPSiteList.Trim().Length > 0)
			{
			   this.SumPSiteWorkTableInsert("max_ti_imp_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				   "a.psite_id=c.chip_haul_cost_psite AND " + 
				   "a.biocd=3 AND " + 
				   "b.cheapest_psite NOT IN (" + strPSiteList.Trim() + ")");
		  }
		  else
		  {
			   this.SumPSiteWorkTableInsert("max_ti_imp_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				   "a.psite_id=c.chip_haul_cost_psite AND " + 
				   "a.biocd=3");

		  }
      if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_ti_imp_sum_psite
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert from psite values from work table to max_ti_imp_sum_psite");
			this.SumPSiteTableInsert("max_ti_imp_sum_psite");
			this.m_frmRunCoreScenario.progressBar1.Value=3;
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************
			 **Finished torch index improvement processsing site
			 *************************************************/

			//MAX_TI_IMP_PNR_SUM_PSITE
			/*************************************************
			 **Torch Index Positive Net Revenue By Processing site
			 *************************************************/
			this.m_strSQL = "delete from psite_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            

			this.m_frmRunCoreScenario.lblMsg.Text="Maximum Torch Index Improvement And Positive Net Revenue By Processing Site";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Torch Index And Positive Net Revenue By Psite");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert chip only psite sums");
			this.SumPSiteWorkTableInsert("max_ti_imp_pnr_plots","chip_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.chip_haul_cost_psite AND a.biocd=2");
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************************
			 **process merchantable only processing sites
			 *************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch only psite sums");
			this.SumPSiteWorkTableInsert("max_ti_imp_pnr_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=1");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the merch_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from merch_haul_cost_psite column");
			this.SumPSiteWorkTableInsert("max_ti_imp_pnr_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=3");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/**********************************************************
			 **get a list of all the psites currently processed
			 **********************************************************/
			strPSiteList = this.m_ado.CreateSQLNOTINString(this.m_TempMDBFileConn,"select psite_id from psite_sum_work_table");
			/*********************************************************************
			 **process psites that process both chiop and merchantable
			 **from the chip_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from chip_haul_cost_psite column");
			if (strPSiteList.Trim().Length > 0)
			{
 			   this.SumPSiteWorkTableInsert("max_ti_imp_pnr_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				   "a.psite_id=c.chip_haul_cost_psite AND " + 
				   "a.biocd=3 AND " + 
			     "b.cheapest_psite NOT IN (" + strPSiteList.Trim() + ")");
		  }
		  else
		  {
		  	this.SumPSiteWorkTableInsert("max_ti_imp_pnr_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				   "a.psite_id=c.chip_haul_cost_psite AND " + 
				   "a.biocd=3");

		  }		
      if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_ti_imp_pnr_sum_psite
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert from psite values from work table to max_ti_imp_pnr_sum_psite");
			this.SumPSiteTableInsert("max_ti_imp_pnr_sum_psite");
			this.m_frmRunCoreScenario.progressBar1.Value=4;
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/**************************************************************
			 **Finished torch index postive net revenue by processsing site
			 ***************************************************************/

			//MAX_CI_IMP_SUM_PSITE
			/*************************************************
			 **most ci improvement by processing site
			 *************************************************/
			this.m_strSQL = "delete from psite_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_frmRunCoreScenario.lblMsg.Text="Maximum Crown Index Improvement By Processing Site";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Crown Index Improvement By Psite");
			this.m_txtStreamWriter.WriteLine(" ");
			/*************************************************
			 **process chip only processing sites
			 *************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert chip only psite sums");
			this.SumPSiteWorkTableInsert("max_ci_imp_plots","chip_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.chip_haul_cost_psite AND a.biocd=2");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************************
			 **process merchantable only processing sites
			 *************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch only psite sums");
			this.SumPSiteWorkTableInsert("max_ci_imp_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=1");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the merch_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from merch_haul_cost_psite column");
			this.SumPSiteWorkTableInsert("max_ci_imp_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=3");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/**********************************************************
			 **get a list of all the psites currently processed
			 **********************************************************/
			strPSiteList = this.m_ado.CreateSQLNOTINString(this.m_TempMDBFileConn,"select psite_id from psite_sum_work_table");
			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the chip_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from chip_haul_cost_psite column");
			if (strPSiteList.Trim().Length > 0)
			{
			   this.SumPSiteWorkTableInsert("max_ci_imp_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				   "a.psite_id=c.chip_haul_cost_psite AND " + 
				   "a.biocd=3 AND " + 
			   	 "b.cheapest_psite NOT IN (" + strPSiteList.Trim() + ")");
			}
			else
			{
			   this.SumPSiteWorkTableInsert("max_ci_imp_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				   "a.psite_id=c.chip_haul_cost_psite AND " + 
				   "a.biocd=3");

			}
      if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_ci_imp_sum_psite
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert from psite values from work table to max_ti_imp_sum_psite");
			this.SumPSiteTableInsert("max_ci_imp_sum_psite");
			this.m_frmRunCoreScenario.progressBar1.Value=5;
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************
			 **Finished torch index improvement processsing site
			 *************************************************/

			//MAX_CI_IMP_PNR_SUM_PSITE
			/*************************************************
			 **Torch Index Positive Net Revenue By Processing site
			 *************************************************/
			this.m_strSQL = "delete from psite_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            

			this.m_frmRunCoreScenario.lblMsg.Text="Maximum Crown Index Improvement And Positive Net Revenue By Processing Site";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Crown Index And Positive Net Revenue By Psite");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert chip only psite sums");
			this.SumPSiteWorkTableInsert("max_ci_imp_pnr_plots","chip_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.chip_haul_cost_psite AND a.biocd=2");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************************
			 **process merchantable only processing sites
			 *************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch only psite sums");
			this.SumPSiteWorkTableInsert("max_ci_imp_pnr_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=1");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the merch_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from merch_haul_cost_psite column");
			this.SumPSiteWorkTableInsert("max_ci_imp_pnr_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=3");
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/**********************************************************
			 **get a list of all the psites currently processed
			 **********************************************************/
			strPSiteList = this.m_ado.CreateSQLNOTINString(this.m_TempMDBFileConn,"select psite_id from psite_sum_work_table");
			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the chip_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from chip_haul_cost_psite column");
			if (strPSiteList.Trim().Length > 0)
			{
			   this.SumPSiteWorkTableInsert("max_ci_imp_pnr_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
			     "a.psite_id=c.chip_haul_cost_psite AND " + 
			   	 "a.biocd=3 AND " + 
				   "b.cheapest_psite NOT IN (" + strPSiteList.Trim() + ")");
			}
			else
			{
			   this.SumPSiteWorkTableInsert("max_ci_imp_pnr_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
			     "a.psite_id=c.chip_haul_cost_psite AND " + 
			   	 "a.biocd=3");
			}
      if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_ci_imp_pnr_sum_psite
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert from psite values from work table to max_ti_imp_pnr_sum_psite");
			this.SumPSiteTableInsert("max_ci_imp_pnr_sum_psite");
			this.m_frmRunCoreScenario.progressBar1.Value=6;
            if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/**************************************************************
			 **Finished torch index postive net revenue by processsing site
			 ***************************************************************/

			//MIN_MERCH_SUM_PSITE
			/*************************************************
			 **Minimum Merchantable Removal By processing site
			 *************************************************/
			this.m_strSQL = "delete from psite_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_frmRunCoreScenario.lblMsg.Text="Minimum Merchantable Removal Treatments By Processing Site";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Minimum Merchantable Removal Treatments By Psite");
			this.m_txtStreamWriter.WriteLine(" ");
			/*************************************************
			 **process chip only processing sites
			 *************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert chip only psite sums");
			this.SumPSiteWorkTableInsert("min_merch_plots","chip_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.chip_haul_cost_psite AND a.biocd=2");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************************
			 **process merchantable only processing sites
			 *************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch only psite sums");
			this.SumPSiteWorkTableInsert("min_merch_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=1");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the merch_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from merch_haul_cost_psite column");
			this.SumPSiteWorkTableInsert("min_merch_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=3");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/**********************************************************
			 **get a list of all the psites currently processed
			 **********************************************************/
			strPSiteList = this.m_ado.CreateSQLNOTINString(this.m_TempMDBFileConn,"select psite_id from psite_sum_work_table");
			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the chip_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from chip_haul_cost_psite column");
			if (strPSiteList.Trim().Length > 0)
			{
			   this.SumPSiteWorkTableInsert("min_merch_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				   "a.psite_id=c.chip_haul_cost_psite AND " + 
				   "a.biocd=3 AND " + 
				   "b.cheapest_psite NOT IN (" + strPSiteList.Trim() + ")");
			}
			else
			{
			   this.SumPSiteWorkTableInsert("min_merch_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				   "a.psite_id=c.chip_haul_cost_psite AND " + 
				   "a.biocd=3");
			}
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into min_merch_sum_psite
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert from psite values from work table to min_merch_sum_psite");
			this.SumPSiteTableInsert("min_merch_sum_psite");
			this.m_frmRunCoreScenario.progressBar1.Value=7;
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************************
			 **Finished Minimum Merchantable Removal By  processsing site
			 *************************************************************/

			//MIN_MERCH_PNR_SUM_PSITE
			/****************************************************************************
			 **Minimum Merchantable Removal And Positive Net Revenue By Processing site
			 ****************************************************************************/
			this.m_strSQL = "delete from psite_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            

			this.m_frmRunCoreScenario.lblMsg.Text="Minimum Merchantable Removal And Positive Net Revenue By Processing Site";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Minumum Merchantable Removal And Positive Net Revenue By Psite");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert chip only psite sums");
			this.SumPSiteWorkTableInsert("min_merch_pnr_plots","chip_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.chip_haul_cost_psite AND a.biocd=2");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/*************************************************************
			 **process merchantable only processing sites
			 *************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch only psite sums");
			this.SumPSiteWorkTableInsert("min_merch_pnr_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=1");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the merch_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from merch_haul_cost_psite column");
			this.SumPSiteWorkTableInsert("min_merch_pnr_plots","merch_haul_cost_psite",
				"WHERE a.psite_id=b.cheapest_psite And a.psite_id=c.merch_haul_cost_psite AND a.biocd=3");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/**********************************************************
			 **get a list of all the psites currently processed
			 **********************************************************/
			strPSiteList = this.m_ado.CreateSQLNOTINString(this.m_TempMDBFileConn,"select psite_id from psite_sum_work_table");
			/*********************************************************************
			 **process psites that process both chip and merchantable
			 **from the chip_haul_cost_psite field
			 *********************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert merch and chip psites sums from chip_haul_cost_psite column");
			if (strPSiteList.Trim().Length > 0)
			{
			   this.SumPSiteWorkTableInsert("min_merch_pnr_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				   "a.psite_id=c.chip_haul_cost_psite AND " + 
				   "a.biocd=3 AND " + 
			   	 "b.cheapest_psite NOT IN (" + strPSiteList.Trim() + ")");
			}
			else
			{
			   this.SumPSiteWorkTableInsert("min_merch_pnr_plots","chip_haul_cost_psite",
				   "WHERE a.psite_id=b.cheapest_psite AND " + 
				   "a.psite_id=c.chip_haul_cost_psite AND " + 
				   "a.biocd=3");
			}
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into min_merch_pnr_sum_psite
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert from psite values from work table to min_merch_pnr_sum_psite");
			this.SumPSiteTableInsert("min_merch_pnr_sum_psite");
			this.m_frmRunCoreScenario.progressBar1.Value=8;
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxPSite) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
				return;
			}
			/**************************************************************
			 **Finished min merch postive net revenue by processsing site
			 ***************************************************************/
			if (this.m_intError == 0)
			{
				this.m_frmRunCoreScenario.lblProcBestRxPSite.ForeColor = System.Drawing.Color.Blue;
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Text = "Completed";
				this.m_frmRunCoreScenario.lblProcBestRxPSite.Refresh();
			}
		}
		private void SumPSiteWorkTableInsert(string strTable,string strPSiteField, string strWhereExpression)
		{
			this.m_strSQL = "INSERT INTO psite_sum_work_table " + 
                             "(psite_id,biocd,sum_acres,sum_chip_yield," + 
                              "sum_chip_haul_cents, sum_chip_dollars_val,sum_merch_haul_cents," + 
				              "sum_merch_vol,sum_net_rev,sum_merch_dollars_val,sum_harv_costs," + 
				              "sum_haul_costs,sum_ti_chg_acres,sum_ci_chg_acres) " + 
                              "SELECT DISTINCT a.psite_id," + 
				                            "a.biocd," +  
				                            "c.sum_acres," + 
				                            "c.sum_chip_yield," + 
				                            "c.sum_chip_haul_cents," +
				                            "c.sum_chip_dollars_val, " + 
				                            "c.sum_merch_haul_cents," + 
				                            "c.sum_merch_vol," + 
				                            "c.sum_net_rev," + 
				                            "c.sum_merch_dollars_val," +
				                            "c.sum_harv_costs," + 
				                            "c.sum_haul_costs," + 
				                            "c.sum_ti_chg_acres," +
				                            "c.sum_ci_chg_acres " +
                            "FROM " + this.m_strPSiteTable.Trim() + " AS a," + 
				              "(SELECT biosum_cond_id," + 
				                      "MIN(" + strPSiteField.Trim() + ") as cheapest_psite " + 
                               "FROM " + strTable + " GROUP BY biosum_cond_id)  b," + 
				              "(SELECT " + strPSiteField.Trim() + "," + 
                                    "SUM(chip_yield)  AS sum_chip_yield," + 
				                    "SUM(chip_haul_cents) AS sum_chip_haul_cents," +
				                    "SUM(chip_dollars_val) AS sum_chip_dollars_val, " +
				                    "SUM(acres) AS sum_acres, " +
				                    "SUM(merch_haul_cents) AS sum_merch_haul_cents," +
				                    "SUM(merch_vol) AS sum_merch_vol," + 
				                    "SUM(net_rev) AS sum_net_rev," + 
				                    "SUM(merch_dollars_val) AS sum_merch_dollars_val," + 
				                    "SUM(harv_costs) AS sum_harv_costs," + 
				                    "SUM(haul_costs) AS sum_haul_costs," + 
				                    "SUM(ti_chg_acres) AS sum_ti_chg_acres," + 
				                    "SUM(ci_chg_acres) AS sum_ci_chg_acres " + 
                               "FROM " + strTable + " GROUP BY " + strPSiteField.Trim() + ") c " +
							   strWhereExpression.Trim() + ";";

			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}
		private void SumPSiteTableInsert(string strTable)
		{
			this.m_strSQL = "INSERT INTO " + strTable + " " + 
				"(psite_id,acres,merch_haul_cents,chip_haul_cents," + 
				"merch_vol,chip_yield,net_rev,merch_dollars_val," + 
				"chip_dollars_val,harv_costs,haul_costs," + 
				"ti_chg_acres,ci_chg_acres) " + 
				"SELECT psite_id,sum_acres AS acres," + 
				"sum_merch_haul_cents AS merch_haul_cents," + 
				"sum_chip_haul_cents AS chip_haul_cents," + 
				"sum_merch_vol AS merch_vol," + 
				"sum_chip_yield AS chip_yield," + 
				"sum_net_rev AS net_rev," + 
				"sum_merch_dollars_val AS merch_dollars_val," + 
				"sum_chip_dollars_val AS chip_dollars_val," + 
				"sum_harv_costs AS harv_costs," + 
				"sum_haul_costs AS haul_costs," + 
				"sum_ti_chg_acres AS ti_chg_acres," + 
				"sum_ci_chg_acres AS ci_chg_acres " + 
				"FROM psite_sum_work_table;";
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}

		/// <summary>
		/// sum wood volumes, values, and costs by land ownership
		/// </summary>
		private void SumOwnership()
		{
			//int x=0;
			this.m_frmRunCoreScenario.lblMsg.Text="Maximum Net Revenue By Ownership";
			this.m_frmRunCoreScenario.lblMsg.Visible=true;
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_frmRunCoreScenario.progressBar1.Minimum=0;
			this.m_frmRunCoreScenario.progressBar1.Maximum = 8;
			this.m_frmRunCoreScenario.progressBar1.Value=0;
			this.m_frmRunCoreScenario.progressBar1.Visible=true;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Ownership Summary");
			this.m_txtStreamWriter.WriteLine("-----------------------");

			/*************************************************
			 **delete all records in the by plot tables
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
			


			
            
			//MAX_NR_SUM_OWN
			/*************************************************
			 **net revenue by ownership
			 *************************************************/

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Net Revenue By Onwership");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table");
			this.SumOwnershipWorkTableInsert("max_nr_plots","own_sum_work_table");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into max_nr_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_nr_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","max_nr_sum_own");
			this.m_frmRunCoreScenario.progressBar1.Value=1;
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}
			/*************************************************
			 **Finished net revenue by ownership
			 *************************************************/

			//MAX_PNR_SUM_OWNERSHIP
			/*************************************************
			 **Positive Net Revenue By Ownership
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            

			this.m_frmRunCoreScenario.lblMsg.Text="Maximum Positive Net Revenue By Ownership";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Positive Net Revenue By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into work table");
			this.SumOwnershipWorkTableInsert("max_pnr_plots","own_sum_work_table");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into max_pnr_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert from ownership values from work table to max_pnr_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","max_pnr_sum_own");
			this.m_frmRunCoreScenario.progressBar1.Value=2;
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}
			/*************************************************
			 **Finished positive net revenue by ownership
			 *************************************************/

			//MAX_TI_IMP_SUM_OWN
			/*************************************************
			 **most ti improvement by ownership
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_frmRunCoreScenario.lblMsg.Text="Maximum Torch Index Improvement By Ownership";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Torch Index Improvement By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("max_ti_imp_plots","own_sum_work_table");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_ti_imp_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_ti_imp_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","max_ti_imp_sum_own");
			
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}

			
			/********************************************************************
			 **most ti improvement by ownership for air curtain destruction plots
			 ********************************************************************/
			this.m_strSQL = "delete from own_sum_work_table_air_dest";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_frmRunCoreScenario.lblMsg.Text="Maximum Torch Index Improvement By Ownership For Air Curtain Destruction Plots";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Torch Index Improvement By Ownership For Air Curtain Destruction Plots");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("max_ti_imp_plots_air_dest","own_sum_work_table_air_dest");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_ti_imp_sum_own_air_dest
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_ti_imp_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table_air_dest","max_ti_imp_sum_own_air_dest");
			
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}

			this.m_frmRunCoreScenario.progressBar1.Value=3;
			/*************************************************
			 **Finished torch index improvement ownership
			 *************************************************/

			//MAX_TI_IMP_PNR_SUM_OWN
			/*************************************************
			 **Torch Index Positive Net Revenue By Ownership
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            

			this.m_frmRunCoreScenario.lblMsg.Text="Maximum Torch Index Improvement And Positive Net Revenue By Ownership";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Torch Index And Positive Net Revenue By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("max_ti_imp_pnr_plots","own_sum_work_table");
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_ti_imp_pnr_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_ti_imp_pnr_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","max_ti_imp_pnr_sum_own");
			this.m_frmRunCoreScenario.progressBar1.Value=4;
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}
			/**************************************************************
			 **Finished torch index postive net revenue by ownership
			 ***************************************************************/

			//MAX_CI_IMP_SUM_OWN
			/*************************************************
			 **most ci improvement by processing site
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_frmRunCoreScenario.lblMsg.Text="Maximum Crown Index Improvement By Ownership";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Crown Index Improvement By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("max_ci_imp_plots","own_sum_work_table");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into max_ci_imp_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_ci_imp_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","max_ci_imp_sum_own");

			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}

			/********************************************************************
			 **most ci improvement by ownership for air curtain destruction plots
			 ********************************************************************/
			this.m_strSQL = "delete from own_sum_work_table_air_dest";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_frmRunCoreScenario.lblMsg.Text="Maximum Crown Index Improvement By Ownership For Air Curtain Destruction Plots";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Crown Index Improvement By Ownership For Air Curtain Destruction Plots");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("max_ci_imp_plots_air_dest","own_sum_work_table_air_dest");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into max_ti_imp_sum_own_air_dest
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_ci_imp_sum_own_air_dest");
			this.SumOwnershipTableInsert("own_sum_work_table_air_dest","max_ci_imp_sum_own_air_dest");
			
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}

			this.m_frmRunCoreScenario.progressBar1.Value=5;
			/*************************************************
			 **Finished crown index improvement by ownership
			 *************************************************/

			//MAX_CI_IMP_PNR_SUM_OWN
			/*************************************************************
			 **Crown Index Improvement Positive Net Revenue By Ownership
			 *************************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            

			this.m_frmRunCoreScenario.lblMsg.Text="Maximum Crown Index Improvement And Positive Net Revenue By Ownership";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Crown Index And Positive Net Revenue By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("max_ci_imp_pnr_plots","own_sum_work_table");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into max_ci_imp_pnr_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to max_ti_imp_pnr_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","max_ci_imp_pnr_sum_own");
			this.m_frmRunCoreScenario.progressBar1.Value=6;
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}
			/**************************************************************
			 **Finished crown index postive net revenue by ownership
			 ***************************************************************/

			//MIN_MERCH_SUM_OWN
			/*************************************************
			 **Minimum Merchantable Removal By Ownership
			 *************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_frmRunCoreScenario.lblMsg.Text="Minimum Merchantable Removal Treatments By Ownership";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Minimum Merchantable Removal Treatments By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");
			/*************************************************
			 **process chip only processing sites
			 *************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("min_merch_plots","own_sum_work_table");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into min_merch_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to min_merch_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","min_merch_sum_own");
			this.m_frmRunCoreScenario.progressBar1.Value=7;
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}

			/********************************************************************************
			 **Minimum Merchantable Removal By Ownership For Air Curtain Destruction Plots
			 ********************************************************************************/
			this.m_strSQL = "delete from own_sum_work_table_air_dest";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			this.m_frmRunCoreScenario.lblMsg.Text="Minimum Merchantable Removal Treatments By Ownership For Air Curtain Destruction Plots";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Minimum Merchantable Removal Treatments By Ownership For Air Curtain Destruction Plots");
			this.m_txtStreamWriter.WriteLine(" ");
			/*************************************************
			 **process chip only processing sites
			 *************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert into ownership work table");
			this.SumOwnershipWorkTableInsert("min_merch_plots_air_dest","own_sum_work_table_air_dest");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}
			/*****************************************************************
			 **insert into min_merch_sum_own_air_dest
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to min_merch_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table_air_dest","min_merch_sum_own_air_dest");
			
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}
			this.m_frmRunCoreScenario.progressBar1.Value=7;
			/*************************************************************
			 **Finished Minimum Merchantable Removal By  Ownership
			 *************************************************************/

			//MIN_MERCH_PNR_SUM_OWN
			/****************************************************************************
			 **Minimum Merchantable Removal And Positive Net Revenue By Ownership
			 ****************************************************************************/
			this.m_strSQL = "delete from own_sum_work_table";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);
            

			this.m_frmRunCoreScenario.lblMsg.Text="Minimum Merchantable Removal And Positive Net Revenue By Ownership";
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Minumum Merchantable Removal And Positive Net Revenue By Ownership");
			this.m_txtStreamWriter.WriteLine(" ");

			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership into work table");
			this.SumOwnershipWorkTableInsert("min_merch_pnr_plots","own_sum_work_table");
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}

			/*****************************************************************
			 **insert into min_merch_pnr_sum_own
			 *****************************************************************/
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Insert ownership values from work table to min_merch_pnr_sum_own");
			this.SumOwnershipTableInsert("own_sum_work_table","min_merch_pnr_sum_own");
			this.m_frmRunCoreScenario.progressBar1.Value=8;
			if (this.UserCancel(this.m_frmRunCoreScenario.lblProcBestRxOwner) == true) return;
			if (this.m_ado.m_intError != 0)
			{
				this.m_txtStreamWriter.WriteLine("!!!Error Executing SQL!!!");
				this.m_intError = this.m_ado.m_intError;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Red;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "!!Error!!";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
				return;
			}
			/**************************************************************
			 **Finished min merch postive net revenue by Ownership
			 ***************************************************************/



			if (this.m_intError == 0)
			{
				this.m_frmRunCoreScenario.lblProcBestRxOwner.ForeColor = System.Drawing.Color.Blue;
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Text = "Completed";
				this.m_frmRunCoreScenario.lblProcBestRxOwner.Refresh();
			}
		}

		private void SumOwnershipWorkTableInsert(string strTableSource,string strTableDestination)
		{
			this.m_strSQL = "INSERT INTO " + strTableDestination + " "  + 
				"(owngrpcd,sum_acres,sum_chip_yield," + 
				"sum_chip_haul_cents, sum_chip_dollars_val,sum_merch_haul_cents," + 
				"sum_merch_vol,sum_net_rev,sum_merch_dollars_val,sum_harv_costs," + 
				"sum_haul_costs,sum_ti_chg_acres,sum_ci_chg_acres) " + 
				"SELECT owngrpcd," + 
				"SUM(acres) AS sum_acres, " +
				"SUM(chip_yield)  AS sum_chip_yield," + 
				"SUM(IIF(chip_haul_cents IS NOT NULL,chip_haul_cents,0)) AS sum_chip_haul_cents," +
				"SUM(IIF(chip_dollars_val IS NOT NULL,chip_dollars_val,0)) AS sum_chip_dollars_val, " +
				"SUM(IIF(merch_haul_cents IS NOT NULL,merch_haul_cents,0)) AS sum_merch_haul_cents," +
				"SUM(merch_vol) AS sum_merch_vol," + 
				"SUM(IIF(net_rev IS NOT NULL,net_rev,0)) AS sum_net_rev," + 
				"SUM(IIF(merch_dollars_val IS NOT NULL,merch_dollars_val,0)) AS sum_merch_dollars_val," + 
				"SUM(harv_costs) AS sum_harv_costs," + 
				"SUM(IIF(haul_costs IS NOT NULL,haul_costs,0)) AS sum_haul_costs," + 
				"SUM(ti_chg_acres) AS sum_ti_chg_acres," + 
				"SUM(ci_chg_acres) AS sum_ci_chg_acres " + 
				"FROM " + strTableSource + " GROUP BY owngrpcd ;";
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}
		
		private void SumOwnershipTableInsert(string strTableSource,string strTableDestination)
		{
			this.m_strSQL = "INSERT INTO " + strTableDestination + " " + 
				"(owngrpcd,acres,merch_haul_cents,chip_haul_cents," + 
				"merch_vol,chip_yield,net_rev,merch_dollars_val," + 
				"chip_dollars_val,harv_costs,haul_costs," + 
				"ti_chg_acres,ci_chg_acres) " + 
				"SELECT owngrpcd,sum_acres AS acres," + 
				"sum_merch_haul_cents AS merch_haul_cents," + 
				"sum_chip_haul_cents AS chip_haul_cents," + 
				"sum_merch_vol AS merch_vol," + 
				"sum_chip_yield AS chip_yield," + 
				"sum_net_rev AS net_rev," + 
				"sum_merch_dollars_val AS merch_dollars_val," + 
				"sum_chip_dollars_val AS chip_dollars_val," + 
				"sum_harv_costs AS harv_costs," + 
				"sum_haul_costs AS haul_costs," + 
				"sum_ti_chg_acres AS ti_chg_acres," + 
				"sum_ci_chg_acres AS ci_chg_acres " + 
				"FROM " + strTableSource;
			this.m_txtStreamWriter.WriteLine("Execute SQL:" + this.m_strSQL);
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

		}
		private void SumAirCurtainDestruction()
		{
			this.m_frmRunCoreScenario.lblMsg.Text="Air Curtain Destruction";
			this.m_frmRunCoreScenario.lblMsg.Visible=true;
			this.m_frmRunCoreScenario.lblMsg.Refresh();
			this.m_frmRunCoreScenario.progressBar1.Minimum=0;
			this.m_frmRunCoreScenario.progressBar1.Maximum = 8;
			this.m_frmRunCoreScenario.progressBar1.Value=0;
			this.m_frmRunCoreScenario.progressBar1.Visible=true;
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine(" ");
			this.m_txtStreamWriter.WriteLine("Air Curtain Destruction Summary");
			this.m_txtStreamWriter.WriteLine("-------------------------------");

			this.m_strSQL = "delete from air_curtain_destruction_plots";
			this.m_ado.SqlNonQuery(this.m_TempMDBFileConn,this.m_strSQL);

			string p = this.m_strPlotTable.Trim();
			string c = this.m_strCondTable.Trim();

			string strWhere = "WHERE " + p + ".merch_haul_cost_psite IS NULL OR " + 
							  p + ".chip_haul_cost_psite IS NULL";

		
			/**************************************************************
			 **Finished air curtain destruction
			 ***************************************************************/



		}
		private void CreateHtml()
		{
			System.IO.FileStream oTxtFileStream;
			System.IO.StreamWriter oTxtStreamWriter;

			oTxtFileStream = new System.IO.FileStream(this.m_frmRunCoreScenario.m_frmScenario.uc_scenario1.txtScenarioPath.Text.Trim() + "\\db\\runstats.htm", System.IO.FileMode.Create, 
				System.IO.FileAccess.Write);
			oTxtStreamWriter = new System.IO.StreamWriter(oTxtFileStream);
			oTxtStreamWriter.WriteLine("<html>\r\n");
			oTxtStreamWriter.WriteLine("<head>\r\n");
			oTxtStreamWriter.WriteLine("<title>\r\n");
			oTxtStreamWriter.WriteLine("FIA Biosum Core Analysis Scenario Run Summary Report\r\n");
			oTxtStreamWriter.WriteLine("</title>\r\n");
            oTxtStreamWriter.WriteLine("<body bgcolor='#ffffff' link='#33339a' vlink='#33339a' alink='#33339a'>\r\n");
			oTxtStreamWriter.WriteLine(System.DateTime.Now.ToString() + "\r\n");
			oTxtStreamWriter.WriteLine("<A NAME='GO TOP'></A>\r\n");
			oTxtStreamWriter.WriteLine("<BR> <BR>\r\n");
			oTxtStreamWriter.WriteLine("<b><CENTER><FONT SIZE='+2' >FIA Biosum Core Analysis Scenario Run Summary Report</FONT></b></center><br>\r\n");
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


		/// <summary>
		/// check and see if the user pressed the cancel button
		/// </summary>
		/// <param name="p_oLabel"></param>
		/// <returns></returns>
		private bool UserCancel(System.Windows.Forms.Label p_oLabel)
		{
			System.Windows.Forms.Application.DoEvents();
			if (this.m_frmRunCoreScenario.m_bUserCancel == true)
			{
				p_oLabel.ForeColor = System.Drawing.Color.Red;
				p_oLabel.Text = "Cancelled";
				return true;
			}
			return false;

		}
	}
}

	

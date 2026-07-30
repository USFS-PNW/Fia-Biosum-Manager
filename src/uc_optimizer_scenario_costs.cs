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
	/// Summary description for uc_scenario_costs.
	/// </summary>
	public class uc_optimizer_scenario_costs : System.Windows.Forms.UserControl
	{
        private System.Windows.Forms.ImageList imgSize;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.ComponentModel.IContainer components;
        private System.Windows.Forms.GroupBox grpboxCost;
        private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtHaulCost;

		public FIA_Biosum_Manager.frmOptimizerScenario m_frmScenario;
		public string[] m_strColumnsToEdit;
		public int m_intColumnsToEditCount=0;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtRailHaulCost;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox txtRailMerchTransfer;
        private System.Windows.Forms.TextBox txtRailChipTransfer;
        private System.Windows.Forms.GroupBox grpboxProfitability;
        private System.Windows.Forms.Label lblUseWithoutProfit;
        private System.Windows.Forms.CheckBox chkMerchUseWithoutProfit;
        private System.Windows.Forms.CheckBox chkChipsUseWithoutProfit;
        private System.Windows.Forms.CheckBox chkWood4UseWithoutProfit;
        private System.Windows.Forms.CheckBox chkWood5UseWithoutProfit;
        private System.Windows.Forms.CheckBox chkWood6UseWithoutProfit;
		private env m_oEnv;
		public System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Panel panel1;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario=null;

        private FIA_Biosum_Manager.ValidateNumericValues m_oValidate = new ValidateNumericValues();
        private string m_strTextHaulCostSave="";
        private string m_strTextRailHaulCostSave = "";
        private string m_strTextRailMerchTransferSave = "";
        private string m_strTextRailChipTransferSave = "";
        public bool m_bMerchUseWithoutProfit = false;
        public bool m_bChipsUseWithoutProfit = false;
        public bool m_bWood4UseWithoutProfit = false;
        public bool m_bWood5UseWithoutProfit = false;
        public bool m_bWood6UseWithoutProfit = false;


		public uc_optimizer_scenario_costs()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

            m_oValidate.RoundDecimalLength = 2;
            m_oValidate.Money = true;
            m_oValidate.NullsAllowed = false;
            m_oValidate.TestForMaxMin = false;
            m_oValidate.RoundDecimalLength = 2;
            m_oValidate.MinValue = 0;
            m_oValidate.TestForMin = true;
			this.m_oEnv = new env();


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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_optimizer_scenario_costs));
            this.imgSize = new System.Windows.Forms.ImageList(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.grpboxCost = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtRailChipTransfer = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtRailMerchTransfer = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtRailHaulCost = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtHaulCost = new System.Windows.Forms.TextBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.grpboxProfitability = new System.Windows.Forms.GroupBox();
            this.lblUseWithoutProfit = new System.Windows.Forms.Label();
            this.chkMerchUseWithoutProfit = new System.Windows.Forms.CheckBox();
            this.chkChipsUseWithoutProfit = new System.Windows.Forms.CheckBox();
            this.chkWood4UseWithoutProfit = new System.Windows.Forms.CheckBox();
            this.chkWood5UseWithoutProfit = new System.Windows.Forms.CheckBox();
            this.chkWood6UseWithoutProfit = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.grpboxCost.SuspendLayout();
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
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(672, 424);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.grpboxCost);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 48);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(666, 373);
            this.panel1.TabIndex = 31;
            this.panel1.Resize += new System.EventHandler(this.panel1_Resize);
            // 
            // grpboxCost
            // 
            this.grpboxCost.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxCost.Controls.Add(this.label6);
            this.grpboxCost.Controls.Add(this.txtRailChipTransfer);
            this.grpboxCost.Controls.Add(this.label4);
            this.grpboxCost.Controls.Add(this.txtRailMerchTransfer);
            this.grpboxCost.Controls.Add(this.label3);
            this.grpboxCost.Controls.Add(this.txtRailHaulCost);
            this.grpboxCost.Controls.Add(this.label2);
            this.grpboxCost.Controls.Add(this.txtHaulCost);
            this.grpboxCost.Controls.Add(this.grpboxProfitability);
            this.grpboxCost.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxCost.ForeColor = System.Drawing.Color.Black;
            this.grpboxCost.Location = new System.Drawing.Point(8, 8);
            this.grpboxCost.Name = "grpboxCost";
            this.grpboxCost.Size = new System.Drawing.Size(648, 314);
            this.grpboxCost.TabIndex = 0;
            this.grpboxCost.TabStop = false;
            this.grpboxCost.Text = "Travel Costs";
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.Black;
            this.label6.Location = new System.Drawing.Point(6, 278);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(483, 24);
            this.label6.TabIndex = 30;
            this.label6.Text = "Truck To Rail Transfer Load Cost (Hog Fuel) $/gt :";
            // 
            // txtRailChipTransfer
            // 
            this.txtRailChipTransfer.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRailChipTransfer.Location = new System.Drawing.Point(546, 278);
            this.txtRailChipTransfer.MaxLength = 10;
            this.txtRailChipTransfer.Name = "txtRailChipTransfer";
            this.txtRailChipTransfer.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtRailChipTransfer.Size = new System.Drawing.Size(80, 26);
            this.txtRailChipTransfer.TabIndex = 4;
            this.txtRailChipTransfer.Text = "$0.00";
            this.txtRailChipTransfer.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtRailChipTransfer.Leave += new System.EventHandler(this.txtRailChipTransfer_Leave);
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(6, 242);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(522, 24);
            this.label4.TabIndex = 28;
            this.label4.Text = "Truck To Rail Transfer Load Cost (Merch) $/gt :";
            // 
            // txtRailMerchTransfer
            // 
            this.txtRailMerchTransfer.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRailMerchTransfer.Location = new System.Drawing.Point(546, 242);
            this.txtRailMerchTransfer.MaxLength = 10;
            this.txtRailMerchTransfer.Name = "txtRailMerchTransfer";
            this.txtRailMerchTransfer.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtRailMerchTransfer.Size = new System.Drawing.Size(80, 26);
            this.txtRailMerchTransfer.TabIndex = 3;
            this.txtRailMerchTransfer.Text = "$0.00";
            this.txtRailMerchTransfer.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtRailMerchTransfer.Leave += new System.EventHandler(this.txtRailMerchTransfer_Leave);
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(6, 206);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(512, 24);
            this.label3.TabIndex = 26;
            this.label3.Text = "Rail Haul Cost Per Green Ton Per Mile:";
            // 
            // txtRailHaulCost
            // 
            this.txtRailHaulCost.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRailHaulCost.Location = new System.Drawing.Point(546, 206);
            this.txtRailHaulCost.MaxLength = 10;
            this.txtRailHaulCost.Name = "txtRailHaulCost";
            this.txtRailHaulCost.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtRailHaulCost.Size = new System.Drawing.Size(80, 26);
            this.txtRailHaulCost.TabIndex = 2;
            this.txtRailHaulCost.Text = "$0.00";
            this.txtRailHaulCost.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtRailHaulCost.Leave += new System.EventHandler(this.txtRailHaulCost_Leave);
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(6, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(522, 24);
            this.label2.TabIndex = 20;
            this.label2.Text = "Truck and Driver Haul Cost per Green Ton Hour (One Direction):";
            // 
            // txtHaulCost
            // 
            this.txtHaulCost.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtHaulCost.Location = new System.Drawing.Point(546, 48);
            this.txtHaulCost.MaxLength = 10;
            this.txtHaulCost.Name = "txtHaulCost";
            this.txtHaulCost.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtHaulCost.Size = new System.Drawing.Size(80, 26);
            this.txtHaulCost.TabIndex = 1;
            this.txtHaulCost.Text = "$0.00";
            this.txtHaulCost.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtHaulCost.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtHaulCost_KeyPress);
            this.txtHaulCost.Leave += new System.EventHandler(this.txtHaulCost_Leave);
            //
            // grpboxProfitability
            //
            this.grpboxProfitability.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxProfitability.Controls.Add(this.lblUseWithoutProfit);
            this.grpboxProfitability.Controls.Add(this.chkMerchUseWithoutProfit);
            this.grpboxProfitability.Controls.Add(this.chkChipsUseWithoutProfit);
            this.grpboxProfitability.Controls.Add(this.chkWood4UseWithoutProfit);
            this.grpboxProfitability.Controls.Add(this.chkWood5UseWithoutProfit);
            this.grpboxProfitability.Controls.Add(this.chkWood6UseWithoutProfit);
            this.grpboxProfitability.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxProfitability.ForeColor = System.Drawing.Color.Black;
            this.grpboxProfitability.Location = new System.Drawing.Point(6, 84);
            this.grpboxProfitability.Name = "grpboxProfitability";
            this.grpboxProfitability.Size = new System.Drawing.Size(500, 110);
            this.grpboxProfitability.TabIndex = 0;
            this.grpboxProfitability.TabStop = false;
            this.grpboxProfitability.Text = "Allocate as Processor-specified wood type, even if net revenue is negative";
            //
            // lblUseWithoutProfit
            //
            this.lblUseWithoutProfit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUseWithoutProfit.ForeColor = System.Drawing.Color.Black;
            this.lblUseWithoutProfit.Location = new System.Drawing.Point(2, 24);
            this.lblUseWithoutProfit.Name = "lblUseWithoutProfit";
            this.lblUseWithoutProfit.Size = new System.Drawing.Size(280, 80);
            this.lblUseWithoutProfit.TabIndex = 27;
            this.lblUseWithoutProfit.Text = "For each wood type, tick box if it should route to a facility of that type, provided one exists, " +
                "even if net revenue negative to do so; otherwise revert it to hog fuel or in-forest residue (details on help screen).";
            //
            // chkMerchUseWithoutProfit
            //
            this.chkMerchUseWithoutProfit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkMerchUseWithoutProfit.Location = new System.Drawing.Point(320, 28);
            this.chkMerchUseWithoutProfit.Name = "chkMerchUseWithoutProfit";
            this.chkMerchUseWithoutProfit.Size = new System.Drawing.Size(80, 20);
            this.chkMerchUseWithoutProfit.TabIndex = 5;
            this.chkMerchUseWithoutProfit.Text = "Merch";
            //
            // chkChipsUseWithoutProfit
            //
            this.chkChipsUseWithoutProfit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkChipsUseWithoutProfit.Location = new System.Drawing.Point(320, 50);
            this.chkChipsUseWithoutProfit.Name = "chkChipsUseWithoutProfit";
            this.chkChipsUseWithoutProfit.Size = new System.Drawing.Size(80, 20);
            this.chkChipsUseWithoutProfit.TabIndex = 5;
            this.chkChipsUseWithoutProfit.Text = "Hog Fuel";
            //
            // chkWood4UseWithoutProfit
            //
            this.chkWood4UseWithoutProfit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkWood4UseWithoutProfit.Location = new System.Drawing.Point(400, 28);
            this.chkWood4UseWithoutProfit.Name = "chkWood4UseWithoutProfit";
            this.chkWood4UseWithoutProfit.Size = new System.Drawing.Size(80, 20);
            this.chkWood4UseWithoutProfit.TabIndex = 6;
            this.chkWood4UseWithoutProfit.Text = "Wood 4";
            //
            // chkWood5UseWithoutProfit
            //
            this.chkWood5UseWithoutProfit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkWood5UseWithoutProfit.Location = new System.Drawing.Point(400, 50);
            this.chkWood5UseWithoutProfit.Name = "chkWood5UseWithoutProfit";
            this.chkWood5UseWithoutProfit.Size = new System.Drawing.Size(80, 20);
            this.chkWood5UseWithoutProfit.TabIndex = 7;
            this.chkWood5UseWithoutProfit.Text = "Wood 5";
            //
            // chkWood6UseWithoutProfit
            //
            this.chkWood6UseWithoutProfit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkWood6UseWithoutProfit.Location = new System.Drawing.Point(400, 72);
            this.chkWood6UseWithoutProfit.Name = "chkWood6UseWithoutProfit";
            this.chkWood6UseWithoutProfit.Size = new System.Drawing.Size(80, 20);
            this.chkWood6UseWithoutProfit.TabIndex = 8;
            this.chkWood6UseWithoutProfit.Text = "Wood 6";
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(666, 32);
            this.lblTitle.TabIndex = 29;
            this.lblTitle.Text = "Haul and Transfer Costs";
            // 
            // uc_scenario_costs
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_scenario_costs";
            this.Size = new System.Drawing.Size(672, 424);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.grpboxCost.ResumeLayout(false);
            this.grpboxCost.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		
		private void grpboxCosts_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.grpboxCost.Left = 16;
			}
			catch
			{
			}
		}

		
		public void loadvalues()
		{
            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oTranCosts.RoadHaulCostPerGreenTonPerHour.Trim().Length > 0)
            {
                txtHaulCost.Text = ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oTranCosts.RoadHaulCostPerGreenTonPerHour;
                txtHaulCost_Leave(null, null);
            }
            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oTranCosts.RailHaulCostPerGreenTonPerMile.Trim().Length > 0)
            {
                txtRailHaulCost.Text = ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oTranCosts.RailHaulCostPerGreenTonPerMile;
                txtRailHaulCost_Leave(null, null);
            }
            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oTranCosts.RailChipTransferPerGreenTon.Trim().Length > 0)
            {
                this.txtRailChipTransfer.Text = ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oTranCosts.RailChipTransferPerGreenTon;
                txtRailChipTransfer_Leave(null, null);
                    
            }
            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oTranCosts.RailMerchTransferPerGreenTon.Trim().Length > 0)
            {
                txtRailMerchTransfer.Text = ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oTranCosts.RailMerchTransferPerGreenTon;
                txtRailMerchTransfer_Leave(null, null);
            }

			
		}
       
		public int savevalues()
		{
			int x = 0;

			string strHaulCost;
			string strRailHaulCost;
			string strRailBioTransferCost;
			string strRailMerchTransferCost;

			//strHaulCost = this.txtHaulCost_subclass.Text.Replace("$","");
			strHaulCost = RoadHaulCostDollarsPerGreenTonPerHour.Replace("$", "");
			strHaulCost = strHaulCost.Replace(",", "");
			if (strHaulCost.Trim().Length == 1) strHaulCost = "0.00";

			//ldp strRailHaulCost = this.txtRailHaulCost_subclass.Text.Replace("$","");
			strRailHaulCost = RailHaulCostDollarsPerGreenTonPerMile.Replace("$", "");
			strRailHaulCost = strRailHaulCost.Replace(",", "");
			if (strRailHaulCost.Trim().Length == 1) strRailHaulCost = "0.00";

			//ldp strRailBioTransferCost = this.txtRailChipTransfer_subclass.Text.Replace("$","");
			strRailBioTransferCost = RailChipTransferCostDollarsPerGreenTonPerHour.Replace("$", "");
			strRailBioTransferCost = strRailBioTransferCost.Replace(",", "");
			if (strRailBioTransferCost.Trim().Length == 1) strRailBioTransferCost = "0.00";

			//ldp strRailMerchTransferCost = this.txtRailMerchTransfer_subclass.Text.Replace("$","");
			strRailMerchTransferCost = RailMerchTransferCostDollarsPerGreenTonPerHour.Replace("$", "");
			strRailMerchTransferCost = strRailMerchTransferCost.Replace(",", "");
			if (strRailMerchTransferCost.Trim().Length == 1) strRailMerchTransferCost = "0.00";


			DataMgr oDataMgr = new DataMgr();
			string strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim();
			string strScenarioDB =
				frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory.Trim() + "\\" +
				Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;

			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strScenarioDB)))
			{
				conn.Open();
				if (oDataMgr.m_intError != 0)
				{
					x = oDataMgr.m_intError;
					oDataMgr = null;
					return x;
				}

				//delete all records from the scenario wind speed class table
				oDataMgr.m_strSQL = "DELETE FROM scenario_costs WHERE " +
					"TRIM(UPPER(scenario_id)) = '" + strScenarioId.Trim().ToUpper() + "'";

				oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
				if (oDataMgr.m_intError < 0)
				{
					conn.Close();
					x = oDataMgr.m_intError;
					oDataMgr = null;
					return x;
				}

				oDataMgr.m_strSQL = "INSERT INTO scenario_costs (scenario_id,road_haul_cost_pgt_per_hour,rail_haul_cost_pgt_per_mile,rail_chip_transfer_pgt,rail_merch_transfer_pgt) " +
						"VALUES ('" + strScenarioId + "'," +
						  strHaulCost + "," + strRailHaulCost + "," + strRailBioTransferCost + "," + strRailMerchTransferCost + ")";
				oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
			}
			return 0;
		}

		public int val_scenario_costs()
		{
			int x=0;
			return x;

		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}
		public void SendSingleKeyStrokes(System.Windows.Forms.TextBox p_oTextBox, string strKeyStrokes)
		{
			string strKeyStroke="";
			p_oTextBox.Focus();
			try 
			{
			
				for (int x=0;x<=strKeyStrokes.Length-1;x++)
				{
					
					switch (strKeyStrokes.Substring(x,1))
					{
						case ")":
							strKeyStroke = "{)}";
							break;
						case "(":
							strKeyStroke = "{(}";
							break;
						case "%":
							strKeyStroke = "{%}";
							break;
						case "^":
							strKeyStroke = "{^}";
							break;
						case "+":
							strKeyStroke = "{+}";
							break;
						case "~":
							strKeyStroke = "{~}";
							break;
						case "[":
							strKeyStroke = "{[}";
							break;
						case "]":
							strKeyStroke = "{]}";
							break;
						case "{":
							strKeyStroke = "{{}";
							break;
						case "}":
							strKeyStroke = "{}}";
							break;
						default:
							strKeyStroke = strKeyStrokes.Substring(x,1).ToString();
							break;

					}
					
					System.Windows.Forms.SendKeys.Send(strKeyStroke);
				
				}
				p_oTextBox.Refresh();
			}
			catch  (Exception caught)
			{
				MessageBox.Show("SendKeyStrokes Method Failed With This Message:" + caught.Message);
			}

		}
		
		public int val_costs()
		{

    		return 0;
		}


		private void txtHaulCost_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
            
		}

		private void txtHaulCost_Leave(object sender, System.EventArgs e)
		{
            m_oValidate.ValidateDecimal(txtHaulCost.Text);
            if (m_oValidate.m_intError == 0)
                txtHaulCost.Text = m_oValidate.ReturnValue;
            else
            {
                this.txtHaulCost.Text = this.m_strTextHaulCostSave;
                this.txtHaulCost.Focus();

            }
			
		}

		private void cmdCosts_Click(object sender, System.EventArgs e)
		{
			
		}
      
		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			
		}

        
		private void panel1_Resize(object sender, System.EventArgs e)
		{
			this.grpboxCost.Left = (int)(panel1.ClientSize.Width * .5) - (int)(this.grpboxCost.Width * .5);

		}
	
		public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
        
        public string RoadHaulCostDollarsPerGreenTonPerHour
        {
            set {this.txtHaulCost.Text=value; this.m_strTextHaulCostSave=value;}
            get {return this.txtHaulCost.Text.Trim();}
        }
        public string RailHaulCostDollarsPerGreenTonPerMile
        {
            set {this.txtRailHaulCost.Text=value; this.m_strTextRailHaulCostSave=value;}
            get {return txtRailHaulCost.Text.Trim();}
        }
        public string RailMerchTransferCostDollarsPerGreenTonPerHour
        {
            set {this.txtRailMerchTransfer.Text=value; this.m_strTextRailMerchTransferSave=value;}
            get {return this.txtRailMerchTransfer.Text.Trim();}
        }
        public string RailChipTransferCostDollarsPerGreenTonPerHour
        {
            set {this.txtRailChipTransfer.Text=value; this.m_strTextRailChipTransferSave=value;}
            get {return this.txtRailChipTransfer.Text.Trim();}
        }




        private void txtRailHaulCost_Leave(object sender, EventArgs e)
        {
            m_oValidate.ValidateDecimal(txtRailHaulCost.Text);
            if (m_oValidate.m_intError == 0)
                txtRailHaulCost.Text = m_oValidate.ReturnValue;
            else
            {
                this.txtRailHaulCost.Text = this.m_strTextRailHaulCostSave;
                this.txtRailHaulCost.Focus();

            }
        }

        private void txtRailMerchTransfer_Leave(object sender, EventArgs e)
        {
             m_oValidate.ValidateDecimal(txtRailMerchTransfer.Text);
            if (m_oValidate.m_intError == 0)
                txtRailMerchTransfer.Text = m_oValidate.ReturnValue;
            else
            {
                txtRailMerchTransfer.Text = this.m_strTextRailMerchTransferSave;
                txtRailMerchTransfer.Focus();

            }
            
        }

        private void txtRailChipTransfer_Leave(object sender, EventArgs e)
        {
            m_oValidate.ValidateDecimal(txtRailChipTransfer.Text);
            if (m_oValidate.m_intError == 0)
                txtRailChipTransfer.Text = m_oValidate.ReturnValue;
            else
            {
                txtRailChipTransfer.Text = this.m_strTextRailChipTransferSave;
                txtRailChipTransfer.Focus();

            }
        }

		
	}
}

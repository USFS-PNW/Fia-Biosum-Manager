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
		//private int m_intFullHt=250;
        private System.Windows.Forms.GroupBox grpboxCost;
        private System.Windows.Forms.Label label2;
		//ldp public FIA_Biosum_Manager.txtDollarsAndCents txtRevPerGreenTon_subclass;
		//private FIA_Biosum_Manager.txtDollarsAndCents txtBrushCut_subclass;
		//private FIA_Biosum_Manager.txtDollarsAndCents txtWaterBarring_subclass;
		private System.Windows.Forms.TextBox txtHaulCost;
		//ldp public FIA_Biosum_Manager.txtDollarsAndCents txtHaulCost_subclass;

        public System.Data.OleDb.OleDbConnection m_OleDbConnectionScenario;
		public FIA_Biosum_Manager.frmOptimizerScenario m_frmScenario;
		private FIA_Biosum_Manager.frmGridView m_frmHarvestCosts;
		public string[] m_strColumnsToEdit;
		public int m_intColumnsToEditCount=0;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtRailHaulCost;
		//ldp public FIA_Biosum_Manager.txtDollarsAndCents txtRailHaulCost_subclass;
		private System.Windows.Forms.Label label4;
		//ldp public FIA_Biosum_Manager.txtDollarsAndCents txtRailMerchTransfer_subclass;
		private System.Windows.Forms.Label label6;
		//ldp public FIA_Biosum_Manager.txtDollarsAndCents txtRailChipTransfer_subclass;
		private System.Windows.Forms.TextBox txtRailMerchTransfer;
        private System.Windows.Forms.TextBox txtRailChipTransfer;
        //ldp public FIA_Biosum_Manager.txtDollarsAndCents txtChipsProcessorMktValPgt_subclass;
		private env m_oEnv;
		public System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblRequired;
        private System.Windows.Forms.Panel panel1;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario=null;

        private FIA_Biosum_Manager.ValidateNumericValues m_oValidate = new ValidateNumericValues();
        private string m_strTextHaulCostSave="";
        private string m_strTextRailHaulCostSave = "";
        private string m_strTextRailMerchTransferSave = "";
        private string m_strTextRailChipTransferSave = "";


       

		//private bool bEdit = true;
		
		//private int intDollarMaxLen=4;
		//private int intCentMaxLen=2;
		//private int intDollarCurLen=0;
		//private int intCentCurLen=0;
		//private string strLastKey = "";

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
            this.lblRequired = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
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
            this.panel1.Controls.Add(this.lblRequired);
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
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.Black;
            this.label6.Location = new System.Drawing.Point(6, 259);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(483, 24);
            this.label6.TabIndex = 30;
            this.label6.Text = "* Truck To Rail Transfer Load Cost (Chips) $/gt :";
            // 
            // txtRailChipTransfer
            // 
            this.txtRailChipTransfer.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRailChipTransfer.Location = new System.Drawing.Point(546, 259);
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
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(6, 190);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(522, 24);
            this.label4.TabIndex = 28;
            this.label4.Text = "* Truck To Rail Transfer Load Cost (Merch) $/gt :";
            // 
            // txtRailMerchTransfer
            // 
            this.txtRailMerchTransfer.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRailMerchTransfer.Location = new System.Drawing.Point(546, 185);
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
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(6, 126);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(512, 24);
            this.label3.TabIndex = 26;
            this.label3.Text = "* Rail Haul Cost Per Green Ton Per Mile:";
            // 
            // txtRailHaulCost
            // 
            this.txtRailHaulCost.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRailHaulCost.Location = new System.Drawing.Point(546, 121);
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
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(6, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(522, 24);
            this.label2.TabIndex = 20;
            this.label2.Text = "* Round Trip Truck and Driver Haul Cost per Green Ton Hour:";
            // 
            // txtHaulCost
            // 
            this.txtHaulCost.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
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
            // lblRequired
            // 
            this.lblRequired.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRequired.ForeColor = System.Drawing.Color.Black;
            this.lblRequired.Location = new System.Drawing.Point(8, 325);
            this.lblRequired.Name = "lblRequired";
            this.lblRequired.Size = new System.Drawing.Size(169, 35);
            this.lblRequired.TabIndex = 25;
            this.lblRequired.Text = "* Required";
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

			//string str="";
			string strSQL = "";
			string strConn = "";
			string strRevPerGreenTon = "";

			string strHaulCost;
			string strRailHaulCost;
			string strRailBioTransferCost;
			string strRailMerchTransferCost;


			//ldp strRevPerGreenTon = this.txtRevPerGreenTon_subclass.Text.Replace("$","");
			//ldp strRevPerGreenTon = strRevPerGreenTon.Replace(",","");
			//ldp if (strRevPerGreenTon.Trim().Length == 1) strRevPerGreenTon = "0.00";




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
			string strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioDB =
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
				Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;

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
				strSQL = "DELETE FROM scenario_costs WHERE " +
					" TRIM(UPPER(scenario_id)) = '" + strScenarioId.Trim().ToUpper() + "';";

				oDataMgr.SqlNonQuery(conn, strSQL);
				if (oDataMgr.m_intError < 0)
				{
					conn.Close();
					x = oDataMgr.m_intError;
					oDataMgr = null;
					return x;
				}

				oDataMgr.m_strSQL = "INSERT INTO scenario_costs (scenario_id,road_haul_cost_pgt_per_hour,rail_haul_cost_pgt_per_mile,rail_chip_transfer_pgt,rail_merch_transfer_pgt)" +
						" VALUES ('" + strScenarioId + "'," +
						  strHaulCost + "," + strRailHaulCost + "," + strRailBioTransferCost + "," + strRailMerchTransferCost + ");";
				oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

				conn.Close();
			}
			return 0;
		}

		public int val_scenario_costs()
		{
			int x=0;
			//ldp if (Convert.ToInt32(this.txtRevPerGreenTon.Text) <= 0)
			//ldp {
			//ldp 	x = -1;
			//ldp 	MessageBox.Show("Run Scenario Failed: Cost per green ton must be a value between 1 and 999","Run Scenario",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);		
			//ldp	return x;
			//ldp }
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

		private void btnHarvestCosts_Click(object sender, System.EventArgs e)
		{
			string strScenarioId="";
			//string strScenarioMDB="";
			string strConn="";
			string strSQL="";
			string strRandomPathAndFile="";
			string strHvstCostsTableName="";
			string strCondTableName="";
			
			string[] strColumnsToEditArray;
			string strColumnsToEditList="";
			string[] strAllColumnsArray;
			string  strAllColumnsList="";
			string strScenarioConn="";
			int x,y;

			strScenarioId =  this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();

			/*****************************************************************
			 **lets see if this harvest costs edit form is already open
			 *****************************************************************/
			utils p_oUtils = new utils();
			p_oUtils.m_intLevel=1;
            if (p_oUtils.FindWindowLike(frmMain.g_oFrmMain.Handle, "Treatment Optimizer: Edit Harvest Costs " + " (" + strScenarioId + ")", "*", true, false) > 0)
			{
				MessageBox.Show("!!Harvest Costs Edit Form Is  Already Open!!","Harvest Costs Edit Form",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				if (this.m_frmHarvestCosts.WindowState == System.Windows.Forms.FormWindowState.Minimized)
					this.m_frmHarvestCosts.WindowState = System.Windows.Forms.FormWindowState.Normal;
				this.m_frmHarvestCosts.Focus();
				return;
			}
			

		
			

			ado_data_access p_ado = new ado_data_access();

			strRandomPathAndFile= 
				this.ReferenceOptimizerScenarioForm.uc_datasource1.CreateMDBAndScenarioTableDataSourceLinks(this.m_oEnv.strTempDir);
			if (strRandomPathAndFile.Trim().Length > 0)
			{
				strConn = p_ado.getMDBConnString(strRandomPathAndFile,"admin","");
			
					strHvstCostsTableName = 
						this.ReferenceOptimizerScenarioForm.uc_datasource1.getDataSourceTableName("Harvest Costs");
					if (strHvstCostsTableName.Trim().Length > 0)
					{
						strCondTableName = 
							this.ReferenceOptimizerScenarioForm.uc_datasource1.getDataSourceTableName("Condition");
						if (strCondTableName.Trim().Length > 0)
						{
							strColumnsToEditArray = new string[1];
							strColumnsToEditList="";

							string strScenarioMDB = 
								frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text + "\\" +
								Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;

							strScenarioConn = p_ado.getMDBConnString(strScenarioMDB,"admin","");
							
							strSQL = "SELECT * FROM scenario_harvest_cost_columns WHERE " + 
								" TRIM(scenario_id) = '" + strScenarioId.Trim() + "';";
							p_ado.SqlQueryReader(strScenarioConn, strSQL);
							if (p_ado.m_OleDbDataReader.HasRows)
							{
								while (p_ado.m_OleDbDataReader.Read())
								{
									if (p_ado.m_OleDbDataReader["ColumnName"] != System.DBNull.Value)
									{
										strColumnsToEditList = strColumnsToEditList + p_ado.m_OleDbDataReader["ColumnName"].ToString().Trim() + ",";
									}
								}
							}
							p_ado.m_OleDbDataReader.Close();
							p_ado.m_OleDbDataReader=null;

							if (strColumnsToEditList.Trim().Length > 0)
							{
								strColumnsToEditList = strColumnsToEditList.Substring(0,strColumnsToEditList.Trim().Length-1);
								strColumnsToEditArray = p_oUtils.ConvertListToArray(strColumnsToEditList,",");
							}
							else
							{
							    //strColumnsToEditArray = new string[2];
								//strColumnsToEditArray[0] = "water_barring_roads_cpa";
								//strColumnsToEditArray[1] = "brush_cutting_cpa";
							}
						    
							strAllColumnsList = p_ado.getFieldNames(strConn,"select * from " + strHvstCostsTableName);
							strAllColumnsArray = p_oUtils.ConvertListToArray(strAllColumnsList,",");


							strSQL = "";
							for (x=0;x<=strAllColumnsArray.Length-1;x++)
							{
								if (strAllColumnsArray[x].Trim().ToUpper()=="BIOSUM_COND_ID")
								{
									strSQL = "biosum_cond_id,";
									strSQL = strSQL + "mid(biosum_cond_id,6,2) as statecd,mid(biosum_cond_id,12,3) as countycd,mid(biosum_cond_id,15,7) as plot,mid(biosum_cond_id,25,1) as condid,";
								}
							
								else
								{
									for (y=0;y<=strColumnsToEditArray.Length-1;y++)
									{
										if (strAllColumnsArray[x].Trim().ToUpper()==strColumnsToEditArray[y].Trim().ToUpper())
										{
											strSQL = strSQL + strColumnsToEditArray[y].Trim() + ",";
										}
									}
								}
							}
							strSQL = strSQL.Substring(0,strSQL.Trim().Length - 1);
							
							strSQL = "SELECT DISTINCT " + strSQL + " FROM " + strHvstCostsTableName;

							this.m_strColumnsToEdit = strColumnsToEditArray;
							this.m_intColumnsToEditCount = m_strColumnsToEdit.Length;

							string[] strRecordKeyField = new string[1];
							strRecordKeyField[0] = "biosum_cond_id";


							this.m_frmHarvestCosts = new frmGridView();
							this.m_frmHarvestCosts.HarvestCostColumns=true;
							this.m_frmHarvestCosts.ReferenceOptimizerScenarioForm=this.ReferenceOptimizerScenarioForm;
							this.m_frmHarvestCosts.LoadDataSetToEdit(strConn,strSQL, strHvstCostsTableName,this.m_strColumnsToEdit,this.m_intColumnsToEditCount,strRecordKeyField); 
							if (this.m_frmHarvestCosts.Visible==false)
							{

								this.m_frmHarvestCosts.MdiParent = this.ParentForm.ParentForm;
                                this.m_frmHarvestCosts.Text = "Treatment Optimizer Analysis: Edit Harvest Costs " + " (" + strScenarioId + ")";
								this.m_frmHarvestCosts.Show();
							}
							this.m_frmHarvestCosts.Focus();


				  
							
						}
					}
				
			}
			p_oUtils = null;
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
			this.lblRequired.Left = this.grpboxCost.Left;

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

using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_scenario_escalators.
	/// </summary>
	public class uc_processor_scenario_movein_costs : System.Windows.Forms.UserControl
	{
        public int m_intError = 0;
        public string m_strError = "";
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Label lblYardDistThreshold;
		private RxTools m_oRxTools = new RxTools();
		private ado_data_access m_oAdo = null;
		private string _strScenarioId="";
        private frmProcessorScenario _frmProcessorScenario = null;
        private TextBox txtMoveInAddend;
        private Label label6;
        private Label label7;
        private TextBox txtMoveInTimeMultiplier;
        private Label label4;
        private Label label5;
        private Label label2;
        private TextBox txtAssumedHarvestArea;
        private Label label3;
        private Label label1;
        private TextBox txtYardDistThreshold;
        private Label label9;
        private Label label8;
        private Label label10;
        private FIA_Biosum_Manager.ValidateNumericValues m_oValidate = new ValidateNumericValues();
		
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmProcessorScenario ReferenceProcessorScenarioForm
		{
			get {return this._frmProcessorScenario;}
			set {this._frmProcessorScenario=value;}
		}
		public string ScenarioId
		{
			get {return _strScenarioId;}
			set {_strScenarioId=value;}
		}

		public uc_processor_scenario_movein_costs()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

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

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_processor_scenario_movein_costs));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txtMoveInAddend = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtMoveInTimeMultiplier = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtAssumedHarvestArea = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtYardDistThreshold = new System.Windows.Forms.TextBox();
            this.lblYardDistThreshold = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(900, 448);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.txtMoveInAddend);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.txtMoveInTimeMultiplier);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.txtAssumedHarvestArea);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.txtYardDistThreshold);
            this.panel1.Controls.Add(this.lblYardDistThreshold);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 50);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(894, 395);
            this.panel1.TabIndex = 31;
            // 
            // label10
            // 
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label10.Location = new System.Drawing.Point(329, 283);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(125, 25);
            this.label10.TabIndex = 68;
            this.label10.Text = "Hours";
            // 
            // label9
            // 
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label9.Location = new System.Drawing.Point(333, 112);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(60, 20);
            this.label9.TabIndex = 67;
            this.label9.Text = "Acres";
            // 
            // label8
            // 
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label8.Location = new System.Drawing.Point(329, 8);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(49, 16);
            this.label8.TabIndex = 66;
            this.label8.Text = "Feet";
            // 
            // txtMoveInAddend
            // 
            this.txtMoveInAddend.ForeColor = System.Drawing.Color.Black;
            this.txtMoveInAddend.Location = new System.Drawing.Point(282, 283);
            this.txtMoveInAddend.Name = "txtMoveInAddend";
            this.txtMoveInAddend.Size = new System.Drawing.Size(45, 22);
            this.txtMoveInAddend.TabIndex = 65;
            this.txtMoveInAddend.Text = "1.0";
            this.txtMoveInAddend.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMoveInHoursAddend_KeyPress);
            this.txtMoveInAddend.Leave += new System.EventHandler(this.txtMoveInHoursAddend_Leave);
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label6.Location = new System.Drawing.Point(16, 303);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(847, 45);
            this.label6.TabIndex = 64;
            this.label6.Text = "Move-in Time will be calculated as the sum of Scaled Move-in Time and this Move-i" +
    "n Adjustment";
            // 
            // label7
            // 
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label7.Location = new System.Drawing.Point(16, 283);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(200, 25);
            this.label7.TabIndex = 63;
            this.label7.Text = "Move-in Adjustment:";
            // 
            // txtMoveInTimeMultiplier
            // 
            this.txtMoveInTimeMultiplier.ForeColor = System.Drawing.Color.Black;
            this.txtMoveInTimeMultiplier.Location = new System.Drawing.Point(282, 177);
            this.txtMoveInTimeMultiplier.Name = "txtMoveInTimeMultiplier";
            this.txtMoveInTimeMultiplier.Size = new System.Drawing.Size(45, 22);
            this.txtMoveInTimeMultiplier.TabIndex = 62;
            this.txtMoveInTimeMultiplier.Text = "0.0";
            this.txtMoveInTimeMultiplier.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMoveInTimeMultiplier_KeyPress);
            this.txtMoveInTimeMultiplier.Leave += new System.EventHandler(this.txtMoveInTimeMultiplier_Leave);
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label4.Location = new System.Drawing.Point(16, 199);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(875, 69);
            this.label4.TabIndex = 61;
            this.label4.Text = resources.GetString("label4.Text");
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label5.Location = new System.Drawing.Point(16, 177);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(250, 20);
            this.label5.TabIndex = 60;
            this.label5.Text = "Move-in Time Multiplier:";
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label2.Location = new System.Drawing.Point(16, 133);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(823, 20);
            this.label2.TabIndex = 59;
            this.label2.Text = "Per-acre move-in costs will be calculated based on this assumption";
            // 
            // txtAssumedHarvestArea
            // 
            this.txtAssumedHarvestArea.ForeColor = System.Drawing.Color.Black;
            this.txtAssumedHarvestArea.Location = new System.Drawing.Point(282, 112);
            this.txtAssumedHarvestArea.Name = "txtAssumedHarvestArea";
            this.txtAssumedHarvestArea.Size = new System.Drawing.Size(45, 22);
            this.txtAssumedHarvestArea.TabIndex = 58;
            this.txtAssumedHarvestArea.Text = "80.0";
            this.txtAssumedHarvestArea.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAssumedHarvestArea_KeyPress);
            this.txtAssumedHarvestArea.Leave += new System.EventHandler(this.txtAssumedHarvestArea_Leave);
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label3.Location = new System.Drawing.Point(16, 112);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(259, 16);
            this.label3.TabIndex = 57;
            this.label3.Text = "Assumed Harvest Area:";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label1.Location = new System.Drawing.Point(16, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(856, 55);
            this.label1.TabIndex = 56;
            this.label1.Text = "Distances shorter than this will be set to this value for cost estimation in OpCo" +
    "st. OpCost may apply a larger value for minimum yarding distance, if warranted f" +
    "or the specified harvest system.";
            // 
            // txtYardDistThreshold
            // 
            this.txtYardDistThreshold.ForeColor = System.Drawing.Color.Black;
            this.txtYardDistThreshold.Location = new System.Drawing.Point(282, 7);
            this.txtYardDistThreshold.Name = "txtYardDistThreshold";
            this.txtYardDistThreshold.Size = new System.Drawing.Size(45, 22);
            this.txtYardDistThreshold.TabIndex = 55;
            this.txtYardDistThreshold.Text = "150.0";
            this.txtYardDistThreshold.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYardDistThreshold_KeyPress);
            this.txtYardDistThreshold.Leave += new System.EventHandler(this.txtYardDistThreshold_Leave);
            // 
            // lblYardDistThreshold
            // 
            this.lblYardDistThreshold.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblYardDistThreshold.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblYardDistThreshold.Location = new System.Drawing.Point(16, 8);
            this.lblYardDistThreshold.Name = "lblYardDistThreshold";
            this.lblYardDistThreshold.Size = new System.Drawing.Size(260, 20);
            this.lblYardDistThreshold.TabIndex = 52;
            this.lblYardDistThreshold.Text = "Yarding Distance Threshold:";
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(894, 32);
            this.lblTitle.TabIndex = 30;
            this.lblTitle.Text = "Move-in costs";
            // 
            // uc_processor_scenario_movein_costs
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_scenario_movein_costs";
            this.Size = new System.Drawing.Size(900, 448);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion
		public void loadvalues()
		{
            ScenarioId = this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadMoveInCosts
                    (frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\db\\scenario_processor_rule_definitions.mdb",
                    ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);
            }
            else
            {
                ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadMoveInCostsSqlite
                    (frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile,
                    ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);
            }

            FIA_Biosum_Manager.ProcessorScenarioItem oItem = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem;
            if (ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.m_intError == 0)
            {
                this.txtYardDistThreshold.Text = oItem.m_oMoveInCosts.YardDistThreshold;
                this.txtAssumedHarvestArea.Text = oItem.m_oMoveInCosts.AssumedHarvestAreaAc;
                this.txtMoveInTimeMultiplier.Text = oItem.m_oMoveInCosts.MoveInTimeMultiplier;
                this.txtMoveInAddend.Text = oItem.m_oMoveInCosts.MoveInHoursAddend;
            }              
			
		}
        public void loadvalues_FromProperties()
        {

            if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oMoveInCosts != null)
            {
                FIA_Biosum_Manager.ProcessorScenarioItem oItem = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem;
                if (ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.m_intError == 0)
                {
                    this.txtYardDistThreshold.Text = oItem.m_oMoveInCosts.YardDistThreshold;
                    this.txtAssumedHarvestArea.Text = oItem.m_oMoveInCosts.AssumedHarvestAreaAc;
                    this.txtMoveInTimeMultiplier.Text = oItem.m_oMoveInCosts.MoveInTimeMultiplier;
                    this.txtMoveInAddend.Text = oItem.m_oMoveInCosts.MoveInHoursAddend;
                }
            }
 
        }

		public void savevalues()
		{
            string strFields = "scenario_id,yard_dist_threshold,assumed_harvest_area_ac," +
                 "move_in_time_multiplier," +
                 "move_in_hours_addend";
            string strValues = "";

            //
            //OPEN CONNECTION TO DB FILE CONTAINING PROCESSOR SCENARIO TABLES
            //
            //scenario mdb connection
            ado_data_access oAdo = null;
            SQLite.ADO.DataMgr oDataMgr = null;
            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                oAdo = new ado_data_access();
                string strScenarioMDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\db\\scenario_processor_rule_definitions.mdb";
                oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB, "", ""));
                if (oAdo.m_intError != 0)
                {
                    m_intError = oAdo.m_intError;
                    m_strError = oAdo.m_strError;
                    oAdo = null;
                    return;
                }
                m_intError = 0;
                m_strError = "";

                oAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName + " " +
                                  "WHERE TRIM(UCASE(scenario_id)) = '" + ScenarioId.Trim().ToUpper() + "'";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            }
            else
            {
                oDataMgr = new SQLite.ADO.DataMgr();
                string strScenarioDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                oDataMgr.OpenConnection(oDataMgr.GetConnectionString(strScenarioDB));
                if (oDataMgr.m_intError != 0)
                {
                    m_intError = oDataMgr.m_intError;
                    m_strError = oDataMgr.m_strError;
                    oDataMgr = null;
                    return;
                }
                m_intError = 0;
                m_strError = "";

                oDataMgr.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName + " " +
                                  "WHERE TRIM(UPPER(scenario_id)) = '" + ScenarioId.Trim().ToUpper() + "'";
                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
            }

            //
            //SCENARIOID
            //
            strValues = "'" + ScenarioId + "',";
            //
            //YARDING DISTANCE THRESHOLD
            //
            if (this.txtYardDistThreshold.Text.Trim().Length > 0)
            {
                strValues = strValues + this.txtYardDistThreshold.Text.Trim() + ",";
            }
            else
            {
                strValues = strValues + "null,";
            }
            //
            //ASSUMED HARVEST AREA
            //
            if (this.txtAssumedHarvestArea.Text.Trim().Length > 0)
            {
                strValues = strValues + this.txtAssumedHarvestArea.Text.Trim() + ",";
            }
            else
            {
                strValues = strValues + "null,";
            }
            //
            //MOVE IN TIME MULTIPLIER
            //
            if (this.txtMoveInTimeMultiplier.Text.Trim().Length > 0)
            {
                strValues = strValues + this.txtMoveInTimeMultiplier.Text.Trim() + ",";
            }
            else
            {
                strValues = strValues + "null,";
            }
            //
            //MOVE IN TIME ADDEND
            //
            if (this.txtMoveInAddend.Text.Trim().Length > 0)
            {
                strValues = strValues + this.txtMoveInAddend.Text.Trim();
            }
            else
            {
                strValues = strValues + "null";
            }
            //

            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                oAdo.m_strSQL = Queries.GetInsertSQL(strFields, strValues, Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName);
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                m_intError = oAdo.m_intError;

                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }
            else
            {
                oDataMgr.m_strSQL = Queries.GetInsertSQL(strFields, strValues, Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName);
                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                m_intError = oDataMgr.m_intError;

                oDataMgr.CloseConnection(oDataMgr.m_Connection);
                oDataMgr = null;
            }			
		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

        private void txtYardDistThreshold_Leave(object sender, System.EventArgs e)
        {
            m_oValidate.RoundDecimalLength = 3;
            m_oValidate.MaxValue = 500;
            m_oValidate.MinValue = 0;
            m_oValidate.TestForMaxMin = true;
            m_oValidate.NullsAllowed = false;
            m_oValidate.ValidateDecimal(this.txtYardDistThreshold.Text.Trim());
            if (m_oValidate.m_intError == 0)
            {
                this.txtYardDistThreshold.Text = m_oValidate.ReturnValue;
            }
            else
                this.txtYardDistThreshold.Focus();
        }

        private void txtAssumedHarvestArea_Leave(object sender, EventArgs e)
        {
            m_oValidate.RoundDecimalLength = 4;
            m_oValidate.MaxValue = 1000;
            m_oValidate.MinValue = 0;
            m_oValidate.TestForMaxMin = true;
            m_oValidate.NullsAllowed = false;
            m_oValidate.ValidateDecimal(this.txtAssumedHarvestArea.Text.Trim());
            if (m_oValidate.m_intError == 0)
            {
                this.txtAssumedHarvestArea.Text = m_oValidate.ReturnValue;
            }
            else
                this.txtAssumedHarvestArea.Focus();
        }

        private void txtMoveInTimeMultiplier_Leave(object sender, EventArgs e)
        {
            m_oValidate.RoundDecimalLength = 2;
            m_oValidate.MaxValue = 25;
            m_oValidate.MinValue = 0;
            m_oValidate.TestForMaxMin = true;
            m_oValidate.NullsAllowed = false;
            m_oValidate.ValidateDecimal(this.txtMoveInTimeMultiplier.Text.Trim());
            if (m_oValidate.m_intError == 0)
            {
                this.txtMoveInTimeMultiplier.Text = m_oValidate.ReturnValue;
            }
            else
                this.txtMoveInTimeMultiplier.Focus();
        }
        
        private void txtMoveInHoursAddend_Leave(object sender, EventArgs e)
        {
            m_oValidate.RoundDecimalLength = 2;
            m_oValidate.MaxValue = 25;
            m_oValidate.MinValue = 0;
            m_oValidate.TestForMaxMin = true;
            m_oValidate.NullsAllowed = false;
            m_oValidate.ValidateDecimal(this.txtMoveInAddend.Text.Trim());
            if (m_oValidate.m_intError == 0)
            {
                this.txtMoveInAddend.Text = m_oValidate.ReturnValue;
            }
            else
                this.txtMoveInAddend.Focus();
        }

        private void txtYardDistThreshold_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }
        private void txtAssumedHarvestArea_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }
        private void txtMoveInTimeMultiplier_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }
        private void txtMoveInHoursAddend_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }
	}
}

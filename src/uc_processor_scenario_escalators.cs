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
	public class uc_processor_scenario_escalators : System.Windows.Forms.UserControl
	{
        public int m_intError = 0;
        public string m_strError = "";
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Panel panel1;
		private FIA_Biosum_Manager.uc_processor_scenario_escalators_value uc_processor_scenario_escalators_value3;
		private FIA_Biosum_Manager.uc_processor_scenario_escalators_value uc_processor_scenario_escalators_value2;
		private FIA_Biosum_Manager.uc_processor_scenario_escalators_value uc_processor_scenario_escalators_value1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblCosts;
		private System.Windows.Forms.Label lblCycle3;
		private System.Windows.Forms.Label lblCycle2;
		private System.Windows.Forms.Label lblCycle1;
		private System.Windows.Forms.Label lblDesc;
		private System.Windows.Forms.Label lblCycleLengthDesc;
		private System.Windows.Forms.Label lblCycleLength;
		private RxTools m_oRxTools = new RxTools();
		private ado_data_access m_oAdo = null;
		private string _strScenarioId="";
		private frmProcessorScenario _frmProcessorScenario=null;
        private Label lblNote;
		
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
        public uc_processor_scenario_escalators_value ReferenceUserControlOperatingEscalator
        {
            get { return this.uc_processor_scenario_escalators_value1; }
        }
        public uc_processor_scenario_escalators_value ReferenceUserControlMerchEscalator
        {
            get { return this.uc_processor_scenario_escalators_value2; }
        }
        public uc_processor_scenario_escalators_value ReferenceUserControlChipsEscalator
        {
            get { return this.uc_processor_scenario_escalators_value3; }
        }

		public uc_processor_scenario_escalators()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_processor_scenario_escalators));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblNote = new System.Windows.Forms.Label();
            this.lblCycleLength = new System.Windows.Forms.Label();
            this.lblCycleLengthDesc = new System.Windows.Forms.Label();
            this.uc_processor_scenario_escalators_value3 = new FIA_Biosum_Manager.uc_processor_scenario_escalators_value();
            this.uc_processor_scenario_escalators_value2 = new FIA_Biosum_Manager.uc_processor_scenario_escalators_value();
            this.uc_processor_scenario_escalators_value1 = new FIA_Biosum_Manager.uc_processor_scenario_escalators_value();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblCosts = new System.Windows.Forms.Label();
            this.lblCycle3 = new System.Windows.Forms.Label();
            this.lblCycle2 = new System.Windows.Forms.Label();
            this.lblCycle1 = new System.Windows.Forms.Label();
            this.lblDesc = new System.Windows.Forms.Label();
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
            this.groupBox1.Size = new System.Drawing.Size(696, 448);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.lblNote);
            this.panel1.Controls.Add(this.lblCycleLength);
            this.panel1.Controls.Add(this.lblCycleLengthDesc);
            this.panel1.Controls.Add(this.uc_processor_scenario_escalators_value3);
            this.panel1.Controls.Add(this.uc_processor_scenario_escalators_value2);
            this.panel1.Controls.Add(this.uc_processor_scenario_escalators_value1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.lblCosts);
            this.panel1.Controls.Add(this.lblCycle3);
            this.panel1.Controls.Add(this.lblCycle2);
            this.panel1.Controls.Add(this.lblCycle1);
            this.panel1.Controls.Add(this.lblDesc);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 50);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(690, 395);
            this.panel1.TabIndex = 31;
            // 
            // lblNote
            // 
            this.lblNote.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNote.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblNote.Location = new System.Drawing.Point(16, 288);
            this.lblNote.Name = "lblNote";
            this.lblNote.Size = new System.Drawing.Size(652, 83);
            this.lblNote.TabIndex = 54;
            this.lblNote.Text = resources.GetString("lblNote.Text");
            // 
            // lblCycleLength
            // 
            this.lblCycleLength.BackColor = System.Drawing.Color.White;
            this.lblCycleLength.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCycleLength.ForeColor = System.Drawing.Color.Black;
            this.lblCycleLength.Location = new System.Drawing.Point(224, 8);
            this.lblCycleLength.Name = "lblCycleLength";
            this.lblCycleLength.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.lblCycleLength.Size = new System.Drawing.Size(32, 24);
            this.lblCycleLength.TabIndex = 53;
            this.lblCycleLength.Text = "10";
            // 
            // lblCycleLengthDesc
            // 
            this.lblCycleLengthDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCycleLengthDesc.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblCycleLengthDesc.Location = new System.Drawing.Point(16, 8);
            this.lblCycleLengthDesc.Name = "lblCycleLengthDesc";
            this.lblCycleLengthDesc.Size = new System.Drawing.Size(200, 16);
            this.lblCycleLengthDesc.TabIndex = 52;
            this.lblCycleLengthDesc.Text = "Current Cycle Length (Years):";
            // 
            // uc_processor_scenario_escalators_value3
            // 
            this.uc_processor_scenario_escalators_value3.Cycle1 = "1.00";
            this.uc_processor_scenario_escalators_value3.Cycle2 = "1.00";
            this.uc_processor_scenario_escalators_value3.Cycle3 = "1.00";
            this.uc_processor_scenario_escalators_value3.Location = new System.Drawing.Point(240, 225);
            this.uc_processor_scenario_escalators_value3.Name = "uc_processor_scenario_escalators_value3";
            this.uc_processor_scenario_escalators_value3.ReferenceProcessorScenarioForm = null;
            this.uc_processor_scenario_escalators_value3.Size = new System.Drawing.Size(408, 32);
            this.uc_processor_scenario_escalators_value3.TabIndex = 51;
            // 
            // uc_processor_scenario_escalators_value2
            // 
            this.uc_processor_scenario_escalators_value2.Cycle1 = "1.00";
            this.uc_processor_scenario_escalators_value2.Cycle2 = "1.00";
            this.uc_processor_scenario_escalators_value2.Cycle3 = "1.00";
            this.uc_processor_scenario_escalators_value2.Location = new System.Drawing.Point(240, 182);
            this.uc_processor_scenario_escalators_value2.Name = "uc_processor_scenario_escalators_value2";
            this.uc_processor_scenario_escalators_value2.ReferenceProcessorScenarioForm = null;
            this.uc_processor_scenario_escalators_value2.Size = new System.Drawing.Size(408, 32);
            this.uc_processor_scenario_escalators_value2.TabIndex = 50;
            // 
            // uc_processor_scenario_escalators_value1
            // 
            this.uc_processor_scenario_escalators_value1.Cycle1 = "1.00";
            this.uc_processor_scenario_escalators_value1.Cycle2 = "1.00";
            this.uc_processor_scenario_escalators_value1.Cycle3 = "1.00";
            this.uc_processor_scenario_escalators_value1.Location = new System.Drawing.Point(240, 142);
            this.uc_processor_scenario_escalators_value1.Name = "uc_processor_scenario_escalators_value1";
            this.uc_processor_scenario_escalators_value1.ReferenceProcessorScenarioForm = null;
            this.uc_processor_scenario_escalators_value1.Size = new System.Drawing.Size(408, 32);
            this.uc_processor_scenario_escalators_value1.TabIndex = 49;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label2.Location = new System.Drawing.Point(37, 233);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(152, 24);
            this.label2.TabIndex = 48;
            this.label2.Text = "Energy Wood Revenue";
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label3.Location = new System.Drawing.Point(37, 190);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(152, 24);
            this.label3.TabIndex = 47;
            this.label3.Text = "Merch. Wood Revenue";
            // 
            // lblCosts
            // 
            this.lblCosts.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCosts.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblCosts.Location = new System.Drawing.Point(37, 150);
            this.lblCosts.Name = "lblCosts";
            this.lblCosts.Size = new System.Drawing.Size(112, 24);
            this.lblCosts.TabIndex = 46;
            this.lblCosts.Text = "Operating Costs";
            // 
            // lblCycle3
            // 
            this.lblCycle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCycle3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblCycle3.Location = new System.Drawing.Point(509, 115);
            this.lblCycle3.Name = "lblCycle3";
            this.lblCycle3.Size = new System.Drawing.Size(112, 24);
            this.lblCycle3.TabIndex = 45;
            this.lblCycle3.Text = "End Of Cycle 3";
            // 
            // lblCycle2
            // 
            this.lblCycle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCycle2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblCycle2.Location = new System.Drawing.Point(381, 115);
            this.lblCycle2.Name = "lblCycle2";
            this.lblCycle2.Size = new System.Drawing.Size(112, 24);
            this.lblCycle2.TabIndex = 44;
            this.lblCycle2.Text = "End Of Cycle 2";
            // 
            // lblCycle1
            // 
            this.lblCycle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCycle1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblCycle1.Location = new System.Drawing.Point(253, 115);
            this.lblCycle1.Name = "lblCycle1";
            this.lblCycle1.Size = new System.Drawing.Size(112, 24);
            this.lblCycle1.TabIndex = 43;
            this.lblCycle1.Text = "End Of Cycle 1";
            // 
            // lblDesc
            // 
            this.lblDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDesc.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblDesc.Location = new System.Drawing.Point(16, 48);
            this.lblDesc.Name = "lblDesc";
            this.lblDesc.Size = new System.Drawing.Size(361, 53);
            this.lblDesc.TabIndex = 42;
            this.lblDesc.Text = " Enter escalator that will be applied as a multiplier to costs or revenue that oc" +
    "cur at the end of each cycle\r\n";
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(690, 32);
            this.lblTitle.TabIndex = 30;
            this.lblTitle.Text = "Cost and Revenue Escalators";
            // 
            // uc_processor_scenario_escalators
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_scenario_escalators";
            this.Size = new System.Drawing.Size(696, 448);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion
		public void loadvalues()
		{
            //
            //SCENARIO ID
            //
            ScenarioId = this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
            if (m_oAdo!=null && m_oAdo.m_OleDbConnection != null)
				if (m_oAdo.m_OleDbConnection.State == System.Data.ConnectionState.Open) m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);

            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                //
                //SCENARIO MDB
                //
                string strScenarioMDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\db\\scenario_processor_rule_definitions.mdb";

                //
                //CREATE LINK IN TEMP MDB TO SCENARIO COST REVENUE ESCALATORS TABLE
                //
                dao_data_access oDao = new dao_data_access();
                //link to tree species groups table
                oDao.CreateTableLink(this.ReferenceProcessorScenarioForm.LoadedQueries.m_strTempDbFile,
                    "scenario_cost_revenue_escalators",
                    strScenarioMDB, "scenario_cost_revenue_escalators", true);
                oDao.m_DaoWorkspace.Close();
                oDao = null;
                //
                //OPEN CONNECTION TO TEMP DB FILE
                //
                m_oAdo = new ado_data_access();
                m_oAdo.OpenConnection(m_oAdo.getMDBConnString(this.ReferenceProcessorScenarioForm.LoadedQueries.m_strTempDbFile, "", ""));
                if (m_oAdo.m_intError == 0)
                {
                    ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadEscalators(
                        m_oAdo, m_oAdo.m_OleDbConnection, ReferenceProcessorScenarioForm.LoadedQueries, ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);
                }
            }
            else
            {
                string strScenarioDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadEscalatorsSqlite(ReferenceProcessorScenarioForm.LoadedQueries, 
                    strScenarioDB, ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);
            }

            ProcessorScenarioItem.Escalators oEscalators = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oEscalators;
                //
                //UPDATE CYCLE ESCALATOR TEXT BOXES
                //
                //operating costs cycle 2,3,4
                //cycle2
                if (oEscalators.OperatingCostsCycle2.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value1.Cycle1 =
                         oEscalators.OperatingCostsCycle2.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value1.Cycle1 = "1.00";
                }
                //cycle3
                if (oEscalators.OperatingCostsCycle3.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value1.Cycle2 =
                         oEscalators.OperatingCostsCycle3.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value1.Cycle2 = "1.00";
                }
                if (oEscalators.OperatingCostsCycle4.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value1.Cycle3 =
                         oEscalators.OperatingCostsCycle4.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value1.Cycle3 = "1.00";
                }
                //merch wood revenue cycle 2,3,4
                //cycle2
                if (oEscalators.MerchWoodRevenueCycle2.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value2.Cycle1 =
                         oEscalators.MerchWoodRevenueCycle2.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value2.Cycle1 = "1.00";
                }
                //cycle3
                if (oEscalators.MerchWoodRevenueCycle3.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value2.Cycle2 =
                         oEscalators.MerchWoodRevenueCycle3.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value2.Cycle2 = "1.00";
                }
                if (oEscalators.MerchWoodRevenueCycle4.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value2.Cycle3 =
                         oEscalators.MerchWoodRevenueCycle4.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value2.Cycle3 = "1.00";
                }
                //Energy wood revenue cycle 2,3,4
                //cycle2
                if (oEscalators.EnergyWoodRevenueCycle2.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value3.Cycle1 =
                         oEscalators.EnergyWoodRevenueCycle2.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value3.Cycle1 = "1.00";
                }
                //cycle3
                if (oEscalators.EnergyWoodRevenueCycle3.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value3.Cycle2 =
                         oEscalators.EnergyWoodRevenueCycle3.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value3.Cycle2 = "1.00";
                }
                if (oEscalators.EnergyWoodRevenueCycle4.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value3.Cycle3 =
                         oEscalators.EnergyWoodRevenueCycle4.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value3.Cycle3 = "1.00";
                }

				
            this.uc_processor_scenario_escalators_value1.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
            this.uc_processor_scenario_escalators_value2.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
            this.uc_processor_scenario_escalators_value3.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
            lblCycleLength.Text = Convert.ToString(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oEscalators.CycleLength);			
		}
        public void loadvalues_FromProperties()
        {

            if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oEscalators != null)
            {
                ProcessorScenarioItem.Escalators oEscalators = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oEscalators;
                //
                //UPDATE CYCLE ESCALATOR TEXT BOXES
                //
                //operating costs cycle 2,3,4
                //cycle2
                if (oEscalators.OperatingCostsCycle2.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value1.Cycle1 =
                         oEscalators.OperatingCostsCycle2.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value1.Cycle1 = "1.00";
                }
                //cycle3
                if (oEscalators.OperatingCostsCycle3.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value1.Cycle2 =
                         oEscalators.OperatingCostsCycle3.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value1.Cycle2 = "1.00";
                }
                if (oEscalators.OperatingCostsCycle4.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value1.Cycle3 =
                         oEscalators.OperatingCostsCycle4.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value1.Cycle3 = "1.00";
                }
                //merch wood revenue cycle 2,3,4
                //cycle2
                if (oEscalators.MerchWoodRevenueCycle2.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value2.Cycle1 =
                         oEscalators.MerchWoodRevenueCycle2.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value2.Cycle1 = "1.00";
                }
                //cycle3
                if (oEscalators.MerchWoodRevenueCycle3.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value2.Cycle2 =
                         oEscalators.MerchWoodRevenueCycle3.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value2.Cycle2 = "1.00";
                }
                if (oEscalators.MerchWoodRevenueCycle4.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value2.Cycle3 =
                         oEscalators.MerchWoodRevenueCycle4.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value2.Cycle3 = "1.00";
                }
                //Energy wood revenue cycle 2,3,4
                //cycle2
                if (oEscalators.EnergyWoodRevenueCycle2.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value3.Cycle1 =
                         oEscalators.EnergyWoodRevenueCycle2.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value3.Cycle1 = "1.00";
                }
                //cycle3
                if (oEscalators.EnergyWoodRevenueCycle3.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value3.Cycle2 =
                         oEscalators.EnergyWoodRevenueCycle3.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value3.Cycle2 = "1.00";
                }
                if (oEscalators.EnergyWoodRevenueCycle4.Trim().Length > 0)
                {
                    uc_processor_scenario_escalators_value3.Cycle3 =
                         oEscalators.EnergyWoodRevenueCycle4.Trim();
                }
                else
                {
                    uc_processor_scenario_escalators_value3.Cycle3 = "1.00";
                }


            }
            this.uc_processor_scenario_escalators_value1.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
            this.uc_processor_scenario_escalators_value2.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
            this.uc_processor_scenario_escalators_value3.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
            lblCycleLength.Text = Convert.ToString(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oEscalators.CycleLength);
        }

		public void savevalues()
		{
            m_intError = 0;
            m_strError = "";

			string strValues="";
			string strFields="scenario_id," +
                             "EscalatorOperatingCosts_Cycle2," +
                             "EscalatorOperatingCosts_Cycle3," +
                             "EscalatorOperatingCosts_Cycle4," +
                             "EscalatorMerchWoodRevenue_Cycle2," +
                             "EscalatorMerchWoodRevenue_Cycle3," +
                             "EscalatorMerchWoodRevenue_Cycle4," +
                             "EscalatorEnergyWoodRevenue_Cycle2," +
                             "EscalatorEnergyWoodRevenue_Cycle3," +
                             "EscalatorEnergyWoodRevenue_Cycle4";


            try
            {
                SQLite.ADO.DataMgr oDataMgr = null;
                //
                //DELETE THE CURRENT SCENARIO RECORDS
                //
                if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
                {
                    m_oAdo.m_strSQL = "DELETE FROM scenario_cost_revenue_escalators " +
                        "WHERE TRIM(scenario_id)='" + this.ScenarioId.Trim() + "'";
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
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
                    oDataMgr.m_strSQL = "DELETE FROM scenario_cost_revenue_escalators " +
                        "WHERE TRIM(scenario_id)='" + this.ScenarioId.Trim() + "'";
                    oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                }

                //
                //scenario id
                //
                strValues = "'" + ScenarioId.Trim() + "',";
                //
                //Operating Cost Cycle1
                //
                strValues = strValues + this.uc_processor_scenario_escalators_value1.Cycle1.Trim() + ",";
                //
                //Operating Cost Cycle2
                //
                strValues = strValues + this.uc_processor_scenario_escalators_value1.Cycle2.Trim() + ",";
                //
                //Operating Cost Cycle3
                //
                strValues = strValues + this.uc_processor_scenario_escalators_value1.Cycle3.Trim() + ",";
                //
                //Merch Wood Revenue Cycle1
                //
                strValues = strValues + this.uc_processor_scenario_escalators_value2.Cycle1.Trim() + ",";
                //
                //Merch Wood Revenue Cycle2
                //
                strValues = strValues + this.uc_processor_scenario_escalators_value2.Cycle2.Trim() + ",";
                //
                //Merch Wood Revenue Cycle3
                //
                strValues = strValues + this.uc_processor_scenario_escalators_value2.Cycle3.Trim() + ",";
                //
                //Energy Wood Revenue Cycle1
                //
                strValues = strValues + this.uc_processor_scenario_escalators_value3.Cycle1.Trim() + ",";
                //
                //Energy Wood Revenue Cycle2
                //
                strValues = strValues + this.uc_processor_scenario_escalators_value3.Cycle2.Trim() + ",";
                //
                //Energy Wood Revenue Cycle3
                //
                strValues = strValues + this.uc_processor_scenario_escalators_value3.Cycle3.Trim();

                if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
                {
                    m_oAdo.m_strSQL = Queries.GetInsertSQL(strFields, strValues, "scenario_cost_revenue_escalators");
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                }
                else
                {
                    oDataMgr.m_strSQL = Queries.GetInsertSQL(strFields, strValues, "scenario_cost_revenue_escalators");
                    oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                    m_intError = oDataMgr.m_intError;

                    oDataMgr.CloseConnection(oDataMgr.m_Connection);
                    oDataMgr = null;
                }

                this.uc_processor_scenario_escalators_value1.SaveValues();
                this.uc_processor_scenario_escalators_value2.SaveValues();
                this.uc_processor_scenario_escalators_value3.SaveValues();
            }
            catch (Exception e)
            {
                m_intError = -1;
                m_strError = e.Message;
               
            }


			
		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}

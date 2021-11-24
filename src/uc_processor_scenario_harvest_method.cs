using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_scenario_harvest_method.
	/// </summary>
	public class uc_processor_scenario_harvest_method : System.Windows.Forms.UserControl
	{
		public int m_intError=0;
		public string m_strError="";
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.GroupBox grpboxHarvestMethod;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtDesc;
		private System.Windows.Forms.ComboBox cmbMethod;
		private System.Windows.Forms.Label lblMethod;
		private System.Windows.Forms.Label lblSteepSlopeDesc;
		private System.Windows.Forms.TextBox txtSteepSlopeDesc;
		private System.Windows.Forms.ComboBox cmbSteepSlopeMethod;
		private System.Windows.Forms.Label lblSteepSlopeMethod;
		private System.Windows.Forms.ComboBox cmbSteepSlopePercent;
		private System.Windows.Forms.GroupBox grpboxSteepSlopeHarvestMethod;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtSteepSlopeMinDia;
		private System.Windows.Forms.TextBox txtMinDiaForChips;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox txtMinDiaSmallLogs;
		private System.Windows.Forms.TextBox txtMinDiaLargeLogs;
		//private Queries m_oQueries = new Queries();
		private RxTools m_oRxTools = new RxTools();
		private ado_data_access m_oAdo = new ado_data_access();
		private string _strScenarioId="";
        private frmProcessorScenario _frmProcessorScenario = null;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.ValidateNumericValues m_oValidate = new ValidateNumericValues();
        private Label label7;
        private Label label11;
        private Label label10;
        private Label label9;
        private Label label8;
        private Label label12;
        private TextBox txtWoodlandMerchPct;
        private Label label13;
        private Label label16;
        private Label label17;
        private TextBox txtCullPct;
        private Label label14;
        private Label label15;
        private TextBox txtSaplingMerchPct;
        private RadioButton rdoProcessorSpecified;
        private RadioButton rdoHideLowestCost;
        private RadioButton rdoTreatment;
        private Label label18;
        
		

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_processor_scenario_harvest_method()
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
        public string SelectedHarvestMethod
        {
            get
            {
                if (rdoProcessorSpecified.Checked)
                {
                    return HarvestMethodSelection.SELECTED.Value;
                }
                else
                {
                    return HarvestMethodSelection.RX.Value;
                }
            }
        }
        public string HarvestMethodSteepSlope
        {
            get { return this.cmbSteepSlopeMethod.Text.Trim(); }
        }
        public string HarvestMethodLowSlope
        {
            get { return this.cmbMethod.Text.Trim(); }
        }
        public string MinDiaForChips
        {
            get { return txtMinDiaForChips.Text.Trim(); }
        }
        public string MinDiaForSmallLogs
        {
            get { return txtMinDiaSmallLogs.Text.Trim(); }
        }
        public string MinDiaForLargeLogs
        {
            get { return txtMinDiaLargeLogs.Text.Trim(); }
        }
        public string SteepSlopePercent
        {
            get { return cmbSteepSlopePercent.Text.Trim(); }
        }
        public string MinDiaForAllTreesSteepSlope
        {
            get { return txtSteepSlopeMinDia.Text.Trim(); }
        }
        public string WoodlandMerchPct
        {
            get { return txtWoodlandMerchPct.Text.Trim(); }
        }
        public string SaplingMerchPct
        {
            get { return txtSaplingMerchPct.Text.Trim(); }
        }
        public string CullPct
        {
            get { return txtCullPct.Text.Trim(); }
        }


		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.rdoProcessorSpecified = new System.Windows.Forms.RadioButton();
            this.rdoHideLowestCost = new System.Windows.Forms.RadioButton();
            this.rdoTreatment = new System.Windows.Forms.RadioButton();
            this.label18 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.txtCullPct = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.txtSaplingMerchPct = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txtSteepSlopeMinDia = new System.Windows.Forms.TextBox();
            this.txtMinDiaForChips = new System.Windows.Forms.TextBox();
            this.cmbSteepSlopePercent = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtMinDiaSmallLogs = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.grpboxSteepSlopeHarvestMethod = new System.Windows.Forms.GroupBox();
            this.lblSteepSlopeDesc = new System.Windows.Forms.Label();
            this.txtSteepSlopeDesc = new System.Windows.Forms.TextBox();
            this.cmbSteepSlopeMethod = new System.Windows.Forms.ComboBox();
            this.lblSteepSlopeMethod = new System.Windows.Forms.Label();
            this.txtMinDiaLargeLogs = new System.Windows.Forms.TextBox();
            this.grpboxHarvestMethod = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDesc = new System.Windows.Forms.TextBox();
            this.cmbMethod = new System.Windows.Forms.ComboBox();
            this.lblMethod = new System.Windows.Forms.Label();
            this.txtWoodlandMerchPct = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.grpboxSteepSlopeHarvestMethod.SuspendLayout();
            this.grpboxHarvestMethod.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1000, 569);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(994, 32);
            this.lblTitle.TabIndex = 28;
            this.lblTitle.Text = "Harvest Method";
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.rdoProcessorSpecified);
            this.panel1.Controls.Add(this.rdoHideLowestCost);
            this.panel1.Controls.Add(this.rdoTreatment);
            this.panel1.Controls.Add(this.label18);
            this.panel1.Controls.Add(this.label16);
            this.panel1.Controls.Add(this.label17);
            this.panel1.Controls.Add(this.txtCullPct);
            this.panel1.Controls.Add(this.label14);
            this.panel1.Controls.Add(this.label15);
            this.panel1.Controls.Add(this.txtSaplingMerchPct);
            this.panel1.Controls.Add(this.label12);
            this.panel1.Controls.Add(this.label13);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.txtSteepSlopeMinDia);
            this.panel1.Controls.Add(this.txtMinDiaForChips);
            this.panel1.Controls.Add(this.cmbSteepSlopePercent);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.txtMinDiaSmallLogs);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.grpboxSteepSlopeHarvestMethod);
            this.panel1.Controls.Add(this.txtMinDiaLargeLogs);
            this.panel1.Controls.Add(this.grpboxHarvestMethod);
            this.panel1.Controls.Add(this.txtWoodlandMerchPct);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 18);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(994, 548);
            this.panel1.TabIndex = 0;
            // 
            // rdoProcessorSpecified
            // 
            this.rdoProcessorSpecified.AutoSize = true;
            this.rdoProcessorSpecified.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoProcessorSpecified.ForeColor = System.Drawing.Color.Black;
            this.rdoProcessorSpecified.Location = new System.Drawing.Point(473, 35);
            this.rdoProcessorSpecified.Name = "rdoProcessorSpecified";
            this.rdoProcessorSpecified.Size = new System.Drawing.Size(147, 22);
            this.rdoProcessorSpecified.TabIndex = 43;
            this.rdoProcessorSpecified.TabStop = true;
            this.rdoProcessorSpecified.Text = "Specified below";
            this.rdoProcessorSpecified.UseVisualStyleBackColor = true;
            this.rdoProcessorSpecified.CheckedChanged += new System.EventHandler(this.rdoProcessorSpecified_CheckedChanged);
            // 
            // rdoHideLowestCost
            // 
            this.rdoHideLowestCost.AutoSize = true;
            this.rdoHideLowestCost.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoHideLowestCost.ForeColor = System.Drawing.Color.Black;
            this.rdoHideLowestCost.Location = new System.Drawing.Point(672, 38);
            this.rdoHideLowestCost.Name = "rdoHideLowestCost";
            this.rdoHideLowestCost.Size = new System.Drawing.Size(193, 22);
            this.rdoHideLowestCost.TabIndex = 42;
            this.rdoHideLowestCost.TabStop = true;
            this.rdoHideLowestCost.Text = "Lowest per acre cost ";
            this.rdoHideLowestCost.UseVisualStyleBackColor = true;
            this.rdoHideLowestCost.Visible = false;
            this.rdoHideLowestCost.CheckedChanged += new System.EventHandler(this.rdoHideLowestCost_CheckedChanged);
            // 
            // rdoTreatment
            // 
            this.rdoTreatment.AutoSize = true;
            this.rdoTreatment.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoTreatment.ForeColor = System.Drawing.Color.Black;
            this.rdoTreatment.Location = new System.Drawing.Point(278, 35);
            this.rdoTreatment.Name = "rdoTreatment";
            this.rdoTreatment.Size = new System.Drawing.Size(184, 22);
            this.rdoTreatment.TabIndex = 41;
            this.rdoTreatment.TabStop = true;
            this.rdoTreatment.Text = "Defined by treatment";
            this.rdoTreatment.UseVisualStyleBackColor = true;
            this.rdoTreatment.CheckedChanged += new System.EventHandler(this.rdoTreatment_CheckedChanged);
            // 
            // label18
            // 
            this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label18.ForeColor = System.Drawing.Color.Black;
            this.label18.Location = new System.Drawing.Point(22, 38);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(200, 16);
            this.label18.TabIndex = 40;
            this.label18.Text = "Harvest Method Selection:";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.ForeColor = System.Drawing.Color.Black;
            this.label16.Location = new System.Drawing.Point(707, 430);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(17, 17);
            this.label16.TabIndex = 39;
            this.label16.Text = "=";
            // 
            // label17
            // 
            this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label17.ForeColor = System.Drawing.Color.Black;
            this.label17.Location = new System.Drawing.Point(773, 404);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(209, 88);
            this.label17.TabIndex = 37;
            this.label17.Text = "Cull threshold, above which trees are assumed nonmerchantable and processed inste" +
    "ad as chips";
            // 
            // txtCullPct
            // 
            this.txtCullPct.ForeColor = System.Drawing.Color.Black;
            this.txtCullPct.Location = new System.Drawing.Point(734, 430);
            this.txtCullPct.Name = "txtCullPct";
            this.txtCullPct.Size = new System.Drawing.Size(33, 22);
            this.txtCullPct.TabIndex = 38;
            this.txtCullPct.Text = "50.0";
            this.txtCullPct.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCullPct_KeyPress);
            this.txtCullPct.Leave += new System.EventHandler(this.txtCullPct_Leave);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.ForeColor = System.Drawing.Color.Black;
            this.label14.Location = new System.Drawing.Point(348, 434);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(17, 17);
            this.label14.TabIndex = 36;
            this.label14.Text = "=";
            // 
            // label15
            // 
            this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.ForeColor = System.Drawing.Color.Black;
            this.label15.Location = new System.Drawing.Point(424, 421);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(280, 58);
            this.label15.TabIndex = 34;
            this.label15.Text = "Percent of sapling biomass assumed of merchantable size";
            // 
            // txtSaplingMerchPct
            // 
            this.txtSaplingMerchPct.ForeColor = System.Drawing.Color.Black;
            this.txtSaplingMerchPct.Location = new System.Drawing.Point(375, 431);
            this.txtSaplingMerchPct.Name = "txtSaplingMerchPct";
            this.txtSaplingMerchPct.Size = new System.Drawing.Size(36, 22);
            this.txtSaplingMerchPct.TabIndex = 35;
            this.txtSaplingMerchPct.Text = "80.0";
            this.txtSaplingMerchPct.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSaplingMerchPct_KeyPress);
            this.txtSaplingMerchPct.Leave += new System.EventHandler(this.txtSaplingMerchPct_Leave);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.ForeColor = System.Drawing.Color.Black;
            this.label12.Location = new System.Drawing.Point(9, 438);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(17, 17);
            this.label12.TabIndex = 33;
            this.label12.Text = "=";
            // 
            // label13
            // 
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.ForeColor = System.Drawing.Color.Black;
            this.label13.Location = new System.Drawing.Point(77, 430);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(258, 54);
            this.label13.TabIndex = 31;
            this.label13.Text = "Percent of woodland species biomass assumed of merchantable size";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label10.Location = new System.Drawing.Point(345, 371);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(26, 17);
            this.label10.TabIndex = 18;
            this.label10.Text = ">=";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.ForeColor = System.Drawing.Color.Black;
            this.label11.Location = new System.Drawing.Point(707, 316);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(26, 17);
            this.label11.TabIndex = 19;
            this.label11.Text = ">=";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(423, 360);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(281, 50);
            this.label1.TabIndex = 8;
            this.label1.Text = "Percent slope threshold at which slope is categorized as steep";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.Black;
            this.label7.Location = new System.Drawing.Point(9, 321);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(26, 17);
            this.label7.TabIndex = 17;
            this.label7.Text = ">=";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.Black;
            this.label8.Location = new System.Drawing.Point(4, 375);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(26, 17);
            this.label8.TabIndex = 18;
            this.label8.Text = ">=";
            // 
            // txtSteepSlopeMinDia
            // 
            this.txtSteepSlopeMinDia.ForeColor = System.Drawing.Color.Black;
            this.txtSteepSlopeMinDia.Location = new System.Drawing.Point(736, 314);
            this.txtSteepSlopeMinDia.Name = "txtSteepSlopeMinDia";
            this.txtSteepSlopeMinDia.Size = new System.Drawing.Size(33, 22);
            this.txtSteepSlopeMinDia.TabIndex = 10;
            this.txtSteepSlopeMinDia.Text = "5.0";
            this.txtSteepSlopeMinDia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSteepSlopeMinDia_KeyPress);
            this.txtSteepSlopeMinDia.Leave += new System.EventHandler(this.txtSteepSlopeMinDia_Leave);
            // 
            // txtMinDiaForChips
            // 
            this.txtMinDiaForChips.ForeColor = System.Drawing.Color.Black;
            this.txtMinDiaForChips.Location = new System.Drawing.Point(36, 317);
            this.txtMinDiaForChips.Name = "txtMinDiaForChips";
            this.txtMinDiaForChips.Size = new System.Drawing.Size(36, 22);
            this.txtMinDiaForChips.TabIndex = 12;
            this.txtMinDiaForChips.Text = "3.0";
            this.txtMinDiaForChips.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMinDiaForChips_KeyPress);
            this.txtMinDiaForChips.Leave += new System.EventHandler(this.txtMinDiaForChips_Leave);
            // 
            // cmbSteepSlopePercent
            // 
            this.cmbSteepSlopePercent.ForeColor = System.Drawing.Color.Black;
            this.cmbSteepSlopePercent.ItemHeight = 16;
            this.cmbSteepSlopePercent.Location = new System.Drawing.Point(375, 368);
            this.cmbSteepSlopePercent.Name = "cmbSteepSlopePercent";
            this.cmbSteepSlopePercent.Size = new System.Drawing.Size(45, 24);
            this.cmbSteepSlopePercent.TabIndex = 0;
            this.cmbSteepSlopePercent.Text = "40";
            this.cmbSteepSlopePercent.SelectedValueChanged += new System.EventHandler(this.cmbSteepSlopePercent_SelectedValueChanged);
            this.cmbSteepSlopePercent.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbSteepSlopePercent_KeyPress);
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(773, 309);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(209, 65);
            this.label3.TabIndex = 9;
            this.label3.Text = "Minimum diameter of trees that will be utilized for any purpose on steep slopes";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.Color.Black;
            this.label9.Location = new System.Drawing.Point(345, 319);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(26, 17);
            this.label9.TabIndex = 19;
            this.label9.Text = ">=";
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(77, 307);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(249, 55);
            this.label4.TabIndex = 11;
            this.label4.Text = "Minimum diameter for chip trees (Trees to be chipped for energy wood)";
            this.label4.Click += new System.EventHandler(this.label4_Click);
            // 
            // txtMinDiaSmallLogs
            // 
            this.txtMinDiaSmallLogs.ForeColor = System.Drawing.Color.Black;
            this.txtMinDiaSmallLogs.Location = new System.Drawing.Point(35, 371);
            this.txtMinDiaSmallLogs.Name = "txtMinDiaSmallLogs";
            this.txtMinDiaSmallLogs.Size = new System.Drawing.Size(36, 22);
            this.txtMinDiaSmallLogs.TabIndex = 15;
            this.txtMinDiaSmallLogs.Text = "7.0";
            this.txtMinDiaSmallLogs.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMinDiaSmallLogs_KeyPress);
            this.txtMinDiaSmallLogs.Leave += new System.EventHandler(this.txtMinDiaSmallLogs_Leave);
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.Black;
            this.label5.Location = new System.Drawing.Point(77, 364);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(258, 60);
            this.label5.TabIndex = 13;
            this.label5.Text = "Minimum diameter for small log trees (trees small enough to fell and process into" +
    " logs by machine)";
            // 
            // grpboxSteepSlopeHarvestMethod
            // 
            this.grpboxSteepSlopeHarvestMethod.Controls.Add(this.lblSteepSlopeDesc);
            this.grpboxSteepSlopeHarvestMethod.Controls.Add(this.txtSteepSlopeDesc);
            this.grpboxSteepSlopeHarvestMethod.Controls.Add(this.cmbSteepSlopeMethod);
            this.grpboxSteepSlopeHarvestMethod.Controls.Add(this.lblSteepSlopeMethod);
            this.grpboxSteepSlopeHarvestMethod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxSteepSlopeHarvestMethod.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.grpboxSteepSlopeHarvestMethod.Location = new System.Drawing.Point(453, 61);
            this.grpboxSteepSlopeHarvestMethod.Name = "grpboxSteepSlopeHarvestMethod";
            this.grpboxSteepSlopeHarvestMethod.Size = new System.Drawing.Size(412, 245);
            this.grpboxSteepSlopeHarvestMethod.TabIndex = 29;
            this.grpboxSteepSlopeHarvestMethod.TabStop = false;
            this.grpboxSteepSlopeHarvestMethod.Text = "Steep Slopes";
            // 
            // lblSteepSlopeDesc
            // 
            this.lblSteepSlopeDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSteepSlopeDesc.Location = new System.Drawing.Point(18, 56);
            this.lblSteepSlopeDesc.Name = "lblSteepSlopeDesc";
            this.lblSteepSlopeDesc.Size = new System.Drawing.Size(182, 16);
            this.lblSteepSlopeDesc.TabIndex = 7;
            this.lblSteepSlopeDesc.Text = "Description";
            // 
            // txtSteepSlopeDesc
            // 
            this.txtSteepSlopeDesc.Enabled = false;
            this.txtSteepSlopeDesc.ForeColor = System.Drawing.Color.Black;
            this.txtSteepSlopeDesc.Location = new System.Drawing.Point(18, 72);
            this.txtSteepSlopeDesc.Multiline = true;
            this.txtSteepSlopeDesc.Name = "txtSteepSlopeDesc";
            this.txtSteepSlopeDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtSteepSlopeDesc.Size = new System.Drawing.Size(288, 166);
            this.txtSteepSlopeDesc.TabIndex = 6;
            this.txtSteepSlopeDesc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSteepSlopeDesc_KeyPress);
            // 
            // cmbSteepSlopeMethod
            // 
            this.cmbSteepSlopeMethod.Enabled = false;
            this.cmbSteepSlopeMethod.ForeColor = System.Drawing.Color.Black;
            this.cmbSteepSlopeMethod.ItemHeight = 17;
            this.cmbSteepSlopeMethod.Location = new System.Drawing.Point(18, 32);
            this.cmbSteepSlopeMethod.Name = "cmbSteepSlopeMethod";
            this.cmbSteepSlopeMethod.Size = new System.Drawing.Size(294, 25);
            this.cmbSteepSlopeMethod.TabIndex = 5;
            this.cmbSteepSlopeMethod.SelectedIndexChanged += new System.EventHandler(this.cmbSteepSlopeMethod_SelectedValueChanged);
            this.cmbSteepSlopeMethod.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbSteepSlopeMethod_KeyPress);
            // 
            // lblSteepSlopeMethod
            // 
            this.lblSteepSlopeMethod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSteepSlopeMethod.Location = new System.Drawing.Point(18, 16);
            this.lblSteepSlopeMethod.Name = "lblSteepSlopeMethod";
            this.lblSteepSlopeMethod.Size = new System.Drawing.Size(130, 16);
            this.lblSteepSlopeMethod.TabIndex = 4;
            this.lblSteepSlopeMethod.Text = "Harvest Method";
            // 
            // txtMinDiaLargeLogs
            // 
            this.txtMinDiaLargeLogs.ForeColor = System.Drawing.Color.Black;
            this.txtMinDiaLargeLogs.Location = new System.Drawing.Point(375, 316);
            this.txtMinDiaLargeLogs.Name = "txtMinDiaLargeLogs";
            this.txtMinDiaLargeLogs.Size = new System.Drawing.Size(37, 22);
            this.txtMinDiaLargeLogs.TabIndex = 16;
            this.txtMinDiaLargeLogs.Text = "20.0";
            this.txtMinDiaLargeLogs.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMinDiaLargeLogs_KeyPress);
            this.txtMinDiaLargeLogs.Leave += new System.EventHandler(this.txtMinDiaLargeLogs_Leave);
            // 
            // grpboxHarvestMethod
            // 
            this.grpboxHarvestMethod.Controls.Add(this.label2);
            this.grpboxHarvestMethod.Controls.Add(this.txtDesc);
            this.grpboxHarvestMethod.Controls.Add(this.cmbMethod);
            this.grpboxHarvestMethod.Controls.Add(this.lblMethod);
            this.grpboxHarvestMethod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxHarvestMethod.ForeColor = System.Drawing.Color.ForestGreen;
            this.grpboxHarvestMethod.Location = new System.Drawing.Point(16, 61);
            this.grpboxHarvestMethod.Name = "grpboxHarvestMethod";
            this.grpboxHarvestMethod.Size = new System.Drawing.Size(368, 245);
            this.grpboxHarvestMethod.TabIndex = 28;
            this.grpboxHarvestMethod.TabStop = false;
            this.grpboxHarvestMethod.Text = "Low Slopes";
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(16, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(184, 16);
            this.label2.TabIndex = 3;
            this.label2.Text = "Description";
            // 
            // txtDesc
            // 
            this.txtDesc.Enabled = false;
            this.txtDesc.Location = new System.Drawing.Point(16, 72);
            this.txtDesc.Multiline = true;
            this.txtDesc.Name = "txtDesc";
            this.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtDesc.Size = new System.Drawing.Size(288, 166);
            this.txtDesc.TabIndex = 2;
            this.txtDesc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRxDesc_KeyPress);
            // 
            // cmbMethod
            // 
            this.cmbMethod.Enabled = false;
            this.cmbMethod.Location = new System.Drawing.Point(16, 32);
            this.cmbMethod.Name = "cmbMethod";
            this.cmbMethod.Size = new System.Drawing.Size(288, 25);
            this.cmbMethod.TabIndex = 1;
            this.cmbMethod.SelectedIndexChanged += new System.EventHandler(this.cmbMethod_SelectedValueChanged);
            this.cmbMethod.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbMethod_KeyPress);
            // 
            // lblMethod
            // 
            this.lblMethod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMethod.Location = new System.Drawing.Point(16, 16);
            this.lblMethod.Name = "lblMethod";
            this.lblMethod.Size = new System.Drawing.Size(130, 16);
            this.lblMethod.TabIndex = 0;
            this.lblMethod.Text = "Harvest Method";
            // 
            // txtWoodlandMerchPct
            // 
            this.txtWoodlandMerchPct.ForeColor = System.Drawing.Color.Black;
            this.txtWoodlandMerchPct.Location = new System.Drawing.Point(36, 435);
            this.txtWoodlandMerchPct.Name = "txtWoodlandMerchPct";
            this.txtWoodlandMerchPct.Size = new System.Drawing.Size(36, 22);
            this.txtWoodlandMerchPct.TabIndex = 32;
            this.txtWoodlandMerchPct.Text = "60.0";
            this.txtWoodlandMerchPct.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtWoodlandMerchPct_KeyPress);
            this.txtWoodlandMerchPct.Leave += new System.EventHandler(this.txtWoodlandMerchPct_Leave);
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.Black;
            this.label6.Location = new System.Drawing.Point(424, 312);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(289, 35);
            this.label6.TabIndex = 14;
            this.label6.Text = "Minimum diameter for large log trees (Trees that require chainsaw felling)";
            // 
            // uc_processor_scenario_harvest_method
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_scenario_harvest_method";
            this.Size = new System.Drawing.Size(1000, 569);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.grpboxSteepSlopeHarvestMethod.ResumeLayout(false);
            this.grpboxSteepSlopeHarvestMethod.PerformLayout();
            this.grpboxHarvestMethod.ResumeLayout(false);
            this.grpboxHarvestMethod.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}
		public void loadvalues()
		{
            this.ScenarioId = this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();

            this.txtDesc.Text="";
			this.txtSteepSlopeDesc.Text="";
        
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(this.ReferenceProcessorScenarioForm.LoadedQueries.m_strTempDbFile, "", ""));
            if (m_oAdo.m_intError == 0)
            {
                //@ToDo: harvest_methods still in ref_master.mdb
                this.m_oRxTools.LoadRxHarvestMethods(m_oAdo, m_oAdo.m_OleDbConnection, this.ReferenceProcessorScenarioForm.LoadedQueries, cmbMethod, cmbSteepSlopeMethod);
            }

			this.cmbSteepSlopePercent.Items.Clear();
			for (int x=90;x>=10;x=x-5)
			{
				this.cmbSteepSlopePercent.Items.Add(x.ToString().Trim());
			}
			this.cmbSteepSlopePercent.Text = "40";

            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadHarvestMethod
                    (frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\db\\scenario_processor_rule_definitions.mdb",
                    ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);
            }
            else
            {
                ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadHarvestMethodSqlite
                    (frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile,
                    ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);
            }


            FIA_Biosum_Manager.ProcessorScenarioItem oItem = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem;
            if (ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.m_intError == 0)
            {
                if (oItem.m_oHarvestMethod.SelectedHarvestMethod.Equals(HarvestMethodSelection.SELECTED))
                {
                    this.rdoTreatment.Checked = false;
                    this.rdoProcessorSpecified.Checked = true;
                    //this.rdoHideLowestCost.Checked = false;
                }
                else
                {
                    this.rdoTreatment.Checked = true;
                    this.rdoProcessorSpecified.Checked = false;
                    //this.rdoHideLowestCost.Checked = false;
                }

                cmbMethod.Text = oItem.m_oHarvestMethod.HarvestMethodLowSlope;
                cmbSteepSlopeMethod.Text = oItem.m_oHarvestMethod.HarvestMethodSteepSlope;
                // Set default harvest methods if otherwise empty; We need them populated if analyst
                // selects processor-specified or LOWEST-COST treatment
                if (this.rdoProcessorSpecified.Checked == true)
                {
                    if (String.IsNullOrEmpty(cmbMethod.Text))
                        cmbMethod.Text = "Ground-Based Mech WT";
                    if (String.IsNullOrEmpty(cmbSteepSlopeMethod.Text))
                        cmbSteepSlopeMethod.Text = "Cable Manual WT";
                }
                txtMinDiaForChips.Text = oItem.m_oHarvestMethod.MinDiaForChips;
                txtMinDiaSmallLogs.Text = oItem.m_oHarvestMethod.MinDiaForSmallLogs;
                txtMinDiaLargeLogs.Text = oItem.m_oHarvestMethod.MinDiaForLargeLogs;
                cmbSteepSlopePercent.Text = oItem.m_oHarvestMethod.SteepSlopePercent;
                txtSteepSlopeMinDia.Text = oItem.m_oHarvestMethod.MinDiaForAllTreesSteepSlope;
                uc_processor_scenario_run.ScenarioHarvestMethodVariables.ProcessLowSlope =
                     oItem.m_oHarvestMethod.ProcessLowSlope;
                uc_processor_scenario_run.ScenarioHarvestMethodVariables.ProcessSteepSlope =
                    oItem.m_oHarvestMethod.ProcessSteepSlope;
                txtWoodlandMerchPct.Text = oItem.m_oHarvestMethod.WoodlandMerchAsPctOfTotalVol;
                txtSaplingMerchPct.Text = oItem.m_oHarvestMethod.SaplingMerchAsPctOfTotalVol;
                txtCullPct.Text = oItem.m_oHarvestMethod.CullPctThreshold;
            }

			
		}

        public void loadvalues_FromProperties()
        {

            this.txtDesc.Text = "";
            this.txtSteepSlopeDesc.Text = "";
            this.cmbSteepSlopePercent.Items.Clear();
            for (int x = 90; x >= 10; x = x - 5)
            {
                this.cmbSteepSlopePercent.Items.Add(x.ToString().Trim());
            }
            this.cmbSteepSlopePercent.Text = "40";

            if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oHarvestMethod != null)
            {
                FIA_Biosum_Manager.ProcessorScenarioItem oItem = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem;
                if (ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.m_intError == 0)
                {
                    if (oItem.m_oHarvestMethod.SelectedHarvestMethod.Equals(HarvestMethodSelection.SELECTED))
                    {
                        this.rdoTreatment.Checked = false;
                        this.rdoProcessorSpecified.Checked = true;
                        //this.rdoHideLowestCost.Checked = false;
                    }
                    else
                    {
                        this.rdoTreatment.Checked = true;
                        this.rdoProcessorSpecified.Checked = false;
                        //this.rdoHideLowestCost.Checked = false;
                    }
                    this.cmbMethod.Text = oItem.m_oHarvestMethod.HarvestMethodLowSlope;
                    cmbSteepSlopeMethod.Text = oItem.m_oHarvestMethod.HarvestMethodSteepSlope;
                    txtMinDiaForChips.Text = oItem.m_oHarvestMethod.MinDiaForChips;
                    txtMinDiaSmallLogs.Text = oItem.m_oHarvestMethod.MinDiaForSmallLogs;
                    txtMinDiaLargeLogs.Text = oItem.m_oHarvestMethod.MinDiaForLargeLogs;
                    cmbSteepSlopePercent.Text = oItem.m_oHarvestMethod.SteepSlopePercent;
                    txtSteepSlopeMinDia.Text = oItem.m_oHarvestMethod.MinDiaForAllTreesSteepSlope;
                    uc_processor_scenario_run.ScenarioHarvestMethodVariables.ProcessLowSlope =
                         oItem.m_oHarvestMethod.ProcessLowSlope;
                    uc_processor_scenario_run.ScenarioHarvestMethodVariables.ProcessSteepSlope =
                        oItem.m_oHarvestMethod.ProcessSteepSlope;
                    txtWoodlandMerchPct.Text = oItem.m_oHarvestMethod.WoodlandMerchAsPctOfTotalVol;
                    txtCullPct.Text = oItem.m_oHarvestMethod.CullPctThreshold;
                    txtSaplingMerchPct.Text = oItem.m_oHarvestMethod.SaplingMerchAsPctOfTotalVol;
                }
            }

        }
		public void savevalues()
		{
            //
            //OPEN CONNECTION TO DB FILE CONTAINING PROCESSOR SCENARIO TABLES
            //
            //scenario mdb connection
            ado_data_access oAdo = null;
            SQLite.ADO.DataMgr oDataMgr = null;
            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite == true)
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
                oAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName + " " +
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
                oDataMgr.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName + " " +
                                  "WHERE TRIM(UPPER(scenario_id)) = '" + ScenarioId.Trim().ToUpper() + "'";
                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
            }

            string strFields = "scenario_id,HarvestMethodSelection,HarvestMethodLowSlope," +
                 "HarvestMethodSteepSlope," +
                 "min_chip_dbh,min_sm_log_dbh," +
                 "min_lg_log_dbh,SteepSlope,min_dbh_steep_slope," +
                 "ProcessLowSlopeYN,ProcessSteepSlopeYN,WoodlandMerchAsPercentOfTotalVol," +
                 "SaplingMerchAsPercentOfTotalVol,CullPctThreshold";
            string strValues = "";

            //
            //SCENARIOID
            //
            strValues ="'" + ScenarioId + "',";
			//
			//HARVEST METHOD SELECTION
			//
            if (this.rdoProcessorSpecified.Checked)
            {
                strValues = strValues + "'" + HarvestMethodSelection.SELECTED.Value + "',";
            }
            else
            {
                strValues = strValues + "'" + HarvestMethodSelection.RX.Value + "',";
            }
			//
			//HARVEST METHOD
			//
			strValues=strValues + "'" + this.cmbMethod.Text.Trim() + "',"; 
			//
			//HARVEST METHOD STEEP SLOPE
			//
			strValues=strValues + "'" + this.cmbSteepSlopeMethod.Text.Trim() + "',"; 
			//
			//MINIMUM CHIP DBH
			//
			if (this.txtMinDiaForChips.Text.Trim().Length > 0)
			{
				strValues=strValues + this.txtMinDiaForChips.Text.Trim() + ",";
			}
			else
			{
				strValues=strValues + "null,";
			}
			//
			//MINIMUM DBH FOR SMALL LOGS
			//
			if (this.txtMinDiaSmallLogs.Text.Trim().Length > 0)
			{
				strValues=strValues + this.txtMinDiaSmallLogs.Text.Trim() + ",";
			}
			else
			{
				strValues=strValues + "null,";
			}
			//
			//MINIMUM DBH FOR LARGE LOGS
			//
			if (this.txtMinDiaLargeLogs.Text.Trim().Length > 0)
			{
				strValues=strValues + this.txtMinDiaLargeLogs.Text.Trim() + ",";
			}
			else
			{
				strValues=strValues + "null,";
			}
			//
			//STEEP SLOPE PERCENT
			//
			if (this.cmbSteepSlopePercent.Text.Trim().Length > 0)
			{
				strValues=strValues + this.cmbSteepSlopePercent.Text.Trim() + ",";
			}
			else
			{
				strValues=strValues + "null,";
			}
			//
			//MIN DBH ON STEEP SLOPE
			//
			if (txtSteepSlopeMinDia.Text.Trim().Length > 0)
			{
				strValues=strValues + txtSteepSlopeMinDia.Text.Trim() + ",";
			}
			else
			{
				strValues=strValues + "null,";
			}
            //
            //PROCESS LOW SLOPE DATA DURING RUN
            //
            if (uc_processor_scenario_run.ScenarioHarvestMethodVariables.ProcessLowSlope == true)
            {
                strValues = strValues + "'Y',";
            }
            else
            {
                strValues = strValues + "'N',";
            }
            //
            //PROCESS STEEP SLOPE DATA DURING RUN
            //
            if (uc_processor_scenario_run.ScenarioHarvestMethodVariables.ProcessSteepSlope == true)
            {
                strValues = strValues + "'Y',";
            }
            else
            {
                strValues = strValues + "'N',";
            }
            //
            //PROCESS WOODLAND MERCH PERCENT
            if (txtWoodlandMerchPct.Text.Trim().Length > 0)
            {
                strValues = strValues + txtWoodlandMerchPct.Text.Trim() + ",";
            }
            else
            {
                strValues = strValues + "null,";
            }
            //
            //PROCESS SAPLING MERCH PERCENT
            if (txtSaplingMerchPct.Text.Trim().Length > 0)
            {
                strValues = strValues + txtSaplingMerchPct.Text.Trim() + ",";
            }
            else
            {
                strValues = strValues + "null,";
            }
            //
            //PROCESS CULL PERCENT
            if (txtCullPct.Text.Trim().Length > 0)
            {
                strValues = strValues + txtCullPct.Text.Trim();
            }
            else
            {
                strValues = strValues + "null";
            }

            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite == true)
            {
                oAdo.m_strSQL = Queries.GetInsertSQL(strFields, strValues, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName);
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                m_intError = oAdo.m_intError;

                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }
            else
            {
                oDataMgr.m_strSQL = Queries.GetInsertSQL(strFields, strValues, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName);
                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                m_intError = oDataMgr.m_intError;

                oDataMgr.CloseConnection(oDataMgr.m_Connection);
                oDataMgr = null;
            }			
		}
		private void cmbMethod_SelectedValueChanged(object sender, System.EventArgs e)
		{
            if (ReferenceProcessorScenarioForm.m_bRulesFirstTime == false) ReferenceProcessorScenarioForm.m_bSave = true;
			getDesc();
			
		}
		private void getDesc()
		{
			if (m_oAdo.m_OleDbDataReader.IsClosed==false) return;

            m_oAdo.m_strSQL = Queries.GenericSelectSQL(this.ReferenceProcessorScenarioForm.LoadedQueries.m_oReference.m_strRefHarvestMethodTable, "description", "TRIM(method)='" + cmbMethod.Text.Trim() + "' AND steep_yn = 'N'");
			this.txtDesc.Text = m_oAdo.getSingleStringValueFromSQLQuery(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL,"temp");
		}

		private void cmbSteepSlopeMethod_SelectedValueChanged(object sender, System.EventArgs e)
		{
            if (ReferenceProcessorScenarioForm.m_bRulesFirstTime == false) ReferenceProcessorScenarioForm.m_bSave = true;
			getDescSteepSlope();
		}
		private void getDescSteepSlope()
		{
			if (m_oAdo.m_OleDbDataReader.IsClosed==false) return;

            m_oAdo.m_strSQL = Queries.GenericSelectSQL(this.ReferenceProcessorScenarioForm.LoadedQueries.m_oReference.m_strRefHarvestMethodTable, "description", "TRIM(method)='" + this.cmbSteepSlopeMethod.Text.Trim() + "' AND steep_yn = 'Y'");
			this.txtSteepSlopeDesc.Text = m_oAdo.getSingleStringValueFromSQLQuery(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL,"temp");
		}

		private void txtRxDesc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			e.Handled=false;	
		}

		private void txtDesc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			e.Handled=true;	
		}

		

		private void txtRxDesc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbMethod_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbSteepSlopeMethod_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void txtSteepSlopeDesc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void txtMinDiaForChips_Leave(object sender, System.EventArgs e)
		{
			m_oValidate.RoundDecimalLength=2;
			m_oValidate.MaxValue=300;
			m_oValidate.MinValue=0;
			m_oValidate.TestForMaxMin=true;
			m_oValidate.NullsAllowed=false;
			m_oValidate.ValidateDecimal(this.txtMinDiaForChips.Text.Trim());
			if (m_oValidate.m_intError==0)
			{
				this.txtMinDiaForChips.Text = m_oValidate.ReturnValue;
			}
			else
				this.txtMinDiaForChips.Focus();
		}

		private void txtMinDiaSmallLogs_Leave(object sender, System.EventArgs e)
		{
			m_oValidate.RoundDecimalLength=2;
			m_oValidate.MaxValue=300;
			m_oValidate.MinValue=0;
			m_oValidate.TestForMaxMin=true;
			m_oValidate.NullsAllowed=false;
			m_oValidate.ValidateDecimal(txtMinDiaSmallLogs.Text.Trim());
			if (m_oValidate.m_intError==0)
			{
				txtMinDiaSmallLogs.Text = m_oValidate.ReturnValue;
			}
			else
				txtMinDiaSmallLogs.Focus();
		}

		private void txtMinDiaLargeLogs_Leave(object sender, System.EventArgs e)
		{
			m_oValidate.RoundDecimalLength=2;
			m_oValidate.MaxValue=300;
			m_oValidate.MinValue=10;
			m_oValidate.TestForMaxMin=true;
			m_oValidate.NullsAllowed=false;
			m_oValidate.ValidateDecimal(txtMinDiaLargeLogs.Text.Trim());
			if (m_oValidate.m_intError==0)
			{
				txtMinDiaLargeLogs.Text = m_oValidate.ReturnValue;
			}
			else
				txtMinDiaLargeLogs.Focus();
		}

		private void txtSteepSlopeMinDia_Leave(object sender, System.EventArgs e)
		{
			m_oValidate.RoundDecimalLength=2;
			m_oValidate.MaxValue=300;
			m_oValidate.MinValue=0;
			m_oValidate.TestForMaxMin=true;
			m_oValidate.NullsAllowed=false;
			m_oValidate.ValidateDecimal(txtSteepSlopeMinDia.Text.Trim());
			if (m_oValidate.m_intError==0)
			{
				txtSteepSlopeMinDia.Text = m_oValidate.ReturnValue;
			}
			else
				txtSteepSlopeMinDia.Focus();
		}

        private void txtWoodlandMerchPct_Leave(object sender, System.EventArgs e)
        {
            m_oValidate.RoundDecimalLength = 2;
            m_oValidate.MaxValue = 100;
            m_oValidate.MinValue = 0;
            m_oValidate.TestForMaxMin = true;
            m_oValidate.NullsAllowed = false;
            m_oValidate.ValidateDecimal(txtWoodlandMerchPct.Text.Trim());
            if (m_oValidate.m_intError == 0)
            {
                txtWoodlandMerchPct.Text = m_oValidate.ReturnValue;
            }
            else
                txtWoodlandMerchPct.Focus();
        }

        private void txtSaplingMerchPct_Leave(object sender, System.EventArgs e)
        {
            m_oValidate.RoundDecimalLength = 2;
            m_oValidate.MaxValue = 100;
            m_oValidate.MinValue = 0;
            m_oValidate.TestForMaxMin = true;
            m_oValidate.NullsAllowed = false;
            m_oValidate.ValidateDecimal(txtSaplingMerchPct.Text.Trim());
            if (m_oValidate.m_intError == 0)
            {
                txtSaplingMerchPct.Text = m_oValidate.ReturnValue;
            }
            else
                txtSaplingMerchPct.Focus();
        }

        private void txtCullPct_Leave(object sender, System.EventArgs e)
        {
            m_oValidate.RoundDecimalLength = 2;
            m_oValidate.MaxValue = 100;
            m_oValidate.MinValue = 0;
            m_oValidate.TestForMaxMin = true;
            m_oValidate.NullsAllowed = false;
            m_oValidate.ValidateDecimal(txtCullPct.Text.Trim());
            if (m_oValidate.m_intError == 0)
            {
                txtCullPct.Text = m_oValidate.ReturnValue;
            }
            else
                txtCullPct.Focus();
        }

		private void cmbSteepSlopePercent_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void txtMinDiaForChips_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

        private void txtMinDiaSmallLogs_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

        private void txtMinDiaLargeLogs_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

        private void txtSteepSlopeMinDia_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }
        private void cmbSteepSlopePercent_SelectedValueChanged(object sender, EventArgs e)
        {
            if (ReferenceProcessorScenarioForm.m_bRulesFirstTime==false) ReferenceProcessorScenarioForm.m_bSave = true;
        }
        private void txtWoodlandMerchPct_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }
        private void txtSaplingMerchPct_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }
        private void txtCullPct_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }
        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void rdoTreatment_CheckedChanged(object sender, EventArgs e)
        {
            if (ReferenceProcessorScenarioForm.m_bRulesFirstTime == false) ReferenceProcessorScenarioForm.m_bSave = true;
            if (this.rdoTreatment.Checked)
            {
                this.cmbSteepSlopeMethod.Enabled = false;
                this.cmbMethod.Enabled = false;

                this.txtDesc.Enabled = false;

                this.txtSteepSlopeDesc.Enabled = false;
            }
        }

        private void rdoHideLowestCost_CheckedChanged(object sender, EventArgs e)
        {
            if (ReferenceProcessorScenarioForm.m_bRulesFirstTime == false) ReferenceProcessorScenarioForm.m_bSave = true;
            if (this.rdoHideLowestCost.Checked)
            {
                this.cmbSteepSlopeMethod.Enabled = false;
                this.cmbMethod.Enabled = false;

                this.txtDesc.Enabled = false;

                this.txtSteepSlopeDesc.Enabled = false;
            }
        }

        private void rdoProcessorSpecified_CheckedChanged(object sender, EventArgs e)
        {
            if (ReferenceProcessorScenarioForm.m_bRulesFirstTime == false) ReferenceProcessorScenarioForm.m_bSave = true;
            if (this.rdoProcessorSpecified.Checked)
            {
                cmbMethod.Enabled = true;
                cmbSteepSlopeMethod.Enabled = true;
                this.txtDesc.Enabled = true;
                this.txtSteepSlopeDesc.Enabled = true;
                if (String.IsNullOrEmpty(cmbMethod.Text))
                    cmbMethod.Text = "Ground-Based Mech WT";
                if (String.IsNullOrEmpty(cmbSteepSlopeMethod.Text))
                    cmbSteepSlopeMethod.Text = "Cable Manual WT";
            }
        }

        public void val_data(System.Collections.Generic.IList<string> lstRx)
        {
            this.m_intError = 0;
            if (this.rdoTreatment.Checked == true)
            {
                string strFvsMDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\db\\fvsmaster.mdb";
                m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strFvsMDB, "", ""));
                string strRxClause = "AND RX IN (";
                foreach (string strRx in lstRx)
                {
                    strRxClause = strRxClause + "'" + strRx + "',";
                }
                //Trim the trailing comma
                strRxClause = strRxClause.Substring(0, (strRxClause.Length - 1));
                strRxClause = strRxClause + " )";
                m_oAdo.m_strSQL = "SELECT * from " + Tables.FVS.DefaultRxTableName + " " +
                                  "WHERE (TRIM(HarvestMethodLowSlope) = '' " +
                                  "OR TRIM(HarvestMethodSteepSlope) = '') " + strRxClause;
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                if (m_oAdo.m_intError == 0)
                {
                    if (m_oAdo.m_OleDbDataReader.HasRows)
                    {
                        this.m_intError = -1;
                        MessageBox.Show("Harvest method is selected to be defined by treatment but at " +
                                        "least one treatment is not associated with a harvest method!", "FIA Biosum", 
                                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                        return;
                    }
                }
            }
        }
	
	}

    public class HarvestMethodSelection
    {
        private HarvestMethodSelection(string value) { Value = value; }

        public string Value { get; set; }

        public static HarvestMethodSelection RX { get { return new HarvestMethodSelection("RX"); } }
        public static HarvestMethodSelection LOWEST_COST { get { return new HarvestMethodSelection("LOWEST_COST"); } }
        public static HarvestMethodSelection SELECTED { get { return new HarvestMethodSelection("SELECTED"); } }

        public override bool Equals(object objInput)
        {
            // Is null?
            if (Object.ReferenceEquals(null, objInput))
            {
                return false;
            }

            // Is the same object?
            if (Object.ReferenceEquals(this, objInput))
            {
                return true;
            }

            // Is the same type?
            if (objInput.GetType() != this.GetType())
            {
                return false;
            }

            // Is the value the same?
            HarvestMethodSelection objSelect = objInput as HarvestMethodSelection;
            return objSelect.Value == Value;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                // Choose large primes to avoid hashing collisions
                const int HashingBase = (int)2166136261;
                const int HashingMultiplier = 16777619;

                int hash = HashingBase;
                hash = (hash * HashingMultiplier) ^ (!Object.ReferenceEquals(null, Value) ? Value.GetHashCode() : 0);
                return hash;
            }
        }

    }
}

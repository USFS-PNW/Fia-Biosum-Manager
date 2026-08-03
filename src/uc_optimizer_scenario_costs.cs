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
        private System.Windows.Forms.Label lblAllocation;
        private System.Windows.Forms.Label lblPrevNoFacDG;
        private System.Windows.Forms.Label lblPrevEconDG;
        private System.Windows.Forms.Label lblMerch;
        private System.Windows.Forms.Label lblWood4;
        private System.Windows.Forms.Label lblWood5;
        private System.Windows.Forms.Label lblWood6;
        private System.Windows.Forms.Label lblChips;
        private System.Windows.Forms.CheckBox chkMerchPrevEconDG;
        private System.Windows.Forms.CheckBox chkChipsPrevEconDG;
        private System.Windows.Forms.CheckBox chkWood4PrevEconDG;
        private System.Windows.Forms.CheckBox chkWood5PrevEconDG;
        private System.Windows.Forms.CheckBox chkWood6PrevEconDG;
        private System.Windows.Forms.CheckBox chkMerchPrevNoFacDG;
        private System.Windows.Forms.CheckBox chkWood4PrevNoFacDG;
        private System.Windows.Forms.CheckBox chkWood5PrevNoFacDG;
        private System.Windows.Forms.CheckBox chkWood6PrevNoFacDG;
        private env m_oEnv;
		public System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Panel panel1;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario=null;

        private FIA_Biosum_Manager.ValidateNumericValues m_oValidate = new ValidateNumericValues();
        private string m_strTextHaulCostSave="";
        private string m_strTextRailHaulCostSave = "";
        private string m_strTextRailMerchTransferSave = "";
        private string m_strTextRailChipTransferSave = "";


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
            this.lblAllocation = new System.Windows.Forms.Label();
            this.lblPrevEconDG = new System.Windows.Forms.Label();
            this.lblPrevNoFacDG = new System.Windows.Forms.Label();
            this.lblMerch = new System.Windows.Forms.Label();
            this.lblWood4 = new System.Windows.Forms.Label();
            this.lblWood5 = new System.Windows.Forms.Label();
            this.lblWood6 = new System.Windows.Forms.Label();
            this.lblChips = new System.Windows.Forms.Label();
            this.chkMerchPrevEconDG = new System.Windows.Forms.CheckBox();
            this.chkChipsPrevEconDG = new System.Windows.Forms.CheckBox();
            this.chkWood4PrevEconDG = new System.Windows.Forms.CheckBox();
            this.chkWood5PrevEconDG = new System.Windows.Forms.CheckBox();
            this.chkWood6PrevEconDG = new System.Windows.Forms.CheckBox();
            this.chkMerchPrevNoFacDG = new System.Windows.Forms.CheckBox();
            this.chkWood4PrevNoFacDG = new System.Windows.Forms.CheckBox();
            this.chkWood5PrevNoFacDG = new System.Windows.Forms.CheckBox();
            this.chkWood6PrevNoFacDG = new System.Windows.Forms.CheckBox();
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
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.txtRailChipTransfer);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.txtRailMerchTransfer);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.txtRailHaulCost);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.txtHaulCost);
            this.panel1.Controls.Add(this.grpboxProfitability);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 48);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(666, 373);
            this.panel1.TabIndex = 31;
            this.panel1.Resize += new System.EventHandler(this.panel1_Resize);
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.Black;
            this.label6.Location = new System.Drawing.Point(6, 328);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(483, 24);
            this.label6.TabIndex = 30;
            this.label6.Text = "Truck To Rail Transfer Load Cost (Hog Fuel) $/gt :";
            // 
            // txtRailChipTransfer
            // 
            this.txtRailChipTransfer.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRailChipTransfer.Location = new System.Drawing.Point(546, 328);
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
            this.label4.Location = new System.Drawing.Point(6, 292);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(522, 24);
            this.label4.TabIndex = 28;
            this.label4.Text = "Truck To Rail Transfer Load Cost (Merch) $/gt :";
            // 
            // txtRailMerchTransfer
            // 
            this.txtRailMerchTransfer.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRailMerchTransfer.Location = new System.Drawing.Point(546, 292);
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
            this.label3.Location = new System.Drawing.Point(6, 256);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(512, 24);
            this.label3.TabIndex = 26;
            this.label3.Text = "Rail Haul Cost Per Green Ton Per Mile:";
            // 
            // txtRailHaulCost
            // 
            this.txtRailHaulCost.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRailHaulCost.Location = new System.Drawing.Point(546, 256);
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
            this.label2.Location = new System.Drawing.Point(6, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(522, 24);
            this.label2.TabIndex = 20;
            this.label2.Text = "Trucking Haul Cost per Green Ton Hour (One Direction):";
            // 
            // txtHaulCost
            // 
            this.txtHaulCost.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtHaulCost.Location = new System.Drawing.Point(546, 28);
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
            this.grpboxProfitability.Controls.Add(this.lblAllocation);
            this.grpboxProfitability.Controls.Add(this.lblPrevEconDG);
            this.grpboxProfitability.Controls.Add(this.lblPrevNoFacDG);
            this.grpboxProfitability.Controls.Add(this.lblMerch);
            this.grpboxProfitability.Controls.Add(this.lblWood4);
            this.grpboxProfitability.Controls.Add(this.lblWood5);
            this.grpboxProfitability.Controls.Add(this.lblWood6);
            this.grpboxProfitability.Controls.Add(this.lblChips);
            this.grpboxProfitability.Controls.Add(this.chkMerchPrevEconDG);
            this.grpboxProfitability.Controls.Add(this.chkChipsPrevEconDG);
            this.grpboxProfitability.Controls.Add(this.chkWood4PrevEconDG);
            this.grpboxProfitability.Controls.Add(this.chkWood5PrevEconDG);
            this.grpboxProfitability.Controls.Add(this.chkWood6PrevEconDG);
            this.grpboxProfitability.Controls.Add(this.chkMerchPrevNoFacDG);
            this.grpboxProfitability.Controls.Add(this.chkWood4PrevNoFacDG);
            this.grpboxProfitability.Controls.Add(this.chkWood5PrevNoFacDG);
            this.grpboxProfitability.Controls.Add(this.chkWood6PrevNoFacDG);
            this.grpboxProfitability.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxProfitability.ForeColor = System.Drawing.Color.Black;
            this.grpboxProfitability.Location = new System.Drawing.Point(6, 64);
            this.grpboxProfitability.Name = "grpboxProfitability";
            this.grpboxProfitability.Size = new System.Drawing.Size(620, 180);
            this.grpboxProfitability.TabIndex = 0;
            this.grpboxProfitability.TabStop = false;
            this.grpboxProfitability.Text = "Optional allocation assignment rules:";
            //
            // lblAllocation
            //
            this.lblAllocation.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAllocation.ForeColor = System.Drawing.Color.Black;
            this.lblAllocation.Location = new System.Drawing.Point(2, 24);
            this.lblAllocation.Name = "lblAllocation";
            this.lblAllocation.Size = new System.Drawing.Size(250, 150);
            this.lblAllocation.TabIndex = 27;
            this.lblAllocation.Text = "See help screen for guidance on when checking boxes may be appropriate.";
            //
            // lblPrevEconDG
            //
            this.lblPrevEconDG.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPrevEconDG.TextAlign = ContentAlignment.BottomCenter;
            this.lblPrevEconDG.ForeColor = System.Drawing.Color.Black;
            this.lblPrevEconDG.Location = new System.Drawing.Point(340, 10);
            this.lblPrevEconDG.Name = "lblPrevEconDG";
            this.lblPrevEconDG.Size = new System.Drawing.Size(130, 50);
            this.lblPrevEconDG.TabIndex = 28;
            this.lblPrevEconDG.Text = "Prevent economic downgrade";
            //
            // lblPrevNoFacDG
            //
            this.lblPrevNoFacDG.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPrevNoFacDG.TextAlign = ContentAlignment.BottomCenter;
            this.lblPrevNoFacDG.ForeColor = System.Drawing.Color.Black;
            this.lblPrevNoFacDG.Location = new System.Drawing.Point(490, 10);
            this.lblPrevNoFacDG.Name = "lblPrevNoFacDG";
            this.lblPrevNoFacDG.Size = new System.Drawing.Size(120, 50);
            this.lblPrevNoFacDG.TabIndex = 29;
            this.lblPrevNoFacDG.Text = "Prevent downgrade when no facility exists";
            //
            // lblMerch
            //
            this.lblMerch.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMerch.TextAlign = ContentAlignment.MiddleRight;
            this.lblMerch.ForeColor = System.Drawing.Color.Black;
            this.lblMerch.Location = new System.Drawing.Point(250, 66);
            this.lblMerch.Name = "lblMerch";
            this.lblMerch.Size = new System.Drawing.Size(74, 16);
            this.lblMerch.TabIndex = 30;
            this.lblMerch.Text = "Merch";
            //
            // lblWood4
            //
            this.lblWood4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWood4.TextAlign = ContentAlignment.MiddleRight;
            this.lblWood4.ForeColor = System.Drawing.Color.Black;
            this.lblWood4.Location = new System.Drawing.Point(250, 86);
            this.lblWood4.Name = "lblWood4";
            this.lblWood4.Size = new System.Drawing.Size(74, 16);
            this.lblWood4.TabIndex = 31;
            this.lblWood4.Text = "Wood 4";
            //
            // lblWood5
            //
            this.lblWood5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWood5.TextAlign = ContentAlignment.MiddleRight;
            this.lblWood5.ForeColor = System.Drawing.Color.Black;
            this.lblWood5.Location = new System.Drawing.Point(250, 106);
            this.lblWood5.Name = "lblWood5";
            this.lblWood5.Size = new System.Drawing.Size(74, 16);
            this.lblWood5.TabIndex = 32;
            this.lblWood5.Text = "Wood 5";
            //
            // lblWood6
            //
            this.lblWood6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWood6.TextAlign = ContentAlignment.MiddleRight;
            this.lblWood6.ForeColor = System.Drawing.Color.Black;
            this.lblWood6.Location = new System.Drawing.Point(250, 126);
            this.lblWood6.Name = "lblWood6";
            this.lblWood6.Size = new System.Drawing.Size(74, 16);
            this.lblWood6.TabIndex = 33;
            this.lblWood6.Text = "Wood 6";
            //
            // lblChips
            //
            this.lblChips.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblChips.TextAlign = ContentAlignment.MiddleRight;
            this.lblChips.ForeColor = System.Drawing.Color.Black;
            this.lblChips.Location = new System.Drawing.Point(250, 146);
            this.lblChips.Name = "lblChips";
            this.lblChips.Size = new System.Drawing.Size(74, 16);
            this.lblChips.TabIndex = 31;
            this.lblChips.Text = "Hog Fuel";
            //
            // chkMerchPrevEconDG
            //
            this.chkMerchPrevEconDG.Location = new System.Drawing.Point(399, 67);
            this.chkMerchPrevEconDG.Name = "chkMerchPrevEconDG";
            this.chkMerchPrevEconDG.Size = new System.Drawing.Size(14, 14);
            this.chkMerchPrevEconDG.TabIndex = 5;
            //
            // chkWood4PrevEconDG
            //
            this.chkWood4PrevEconDG.Location = new System.Drawing.Point(399, 87);
            this.chkWood4PrevEconDG.Name = "chkWood4PrevEconDG";
            this.chkWood4PrevEconDG.Size = new System.Drawing.Size(14, 14);
            this.chkWood4PrevEconDG.TabIndex = 6;
            //
            // chkWood5PrevEconDG
            //
            this.chkWood5PrevEconDG.Location = new System.Drawing.Point(399, 107);
            this.chkWood5PrevEconDG.Name = "chkWood5PrevEconDG";
            this.chkWood5PrevEconDG.Size = new System.Drawing.Size(14, 14);
            this.chkWood5PrevEconDG.TabIndex = 7;
            //
            // chkWood6PrevEconDG
            //
            this.chkWood6PrevEconDG.Location = new System.Drawing.Point(399, 127);
            this.chkWood6PrevEconDG.Name = "chkWood6PrevEconDG";
            this.chkWood6PrevEconDG.Size = new System.Drawing.Size(14, 14);
            this.chkWood6PrevEconDG.TabIndex = 8;
            //
            // chkChipsPrevEconDG
            //
            this.chkChipsPrevEconDG.Location = new System.Drawing.Point(399, 147);
            this.chkChipsPrevEconDG.Name = "chkChipsPrevEconDG";
            this.chkChipsPrevEconDG.Size = new System.Drawing.Size(14, 14);
            this.chkChipsPrevEconDG.TabIndex = 9;
            //
            // chkMerchPrevNoFacDG
            //
            this.chkMerchPrevNoFacDG.Location = new System.Drawing.Point(544, 67);
            this.chkMerchPrevNoFacDG.Name = "chkMerchPrevNoFacDG";
            this.chkMerchPrevNoFacDG.Size = new System.Drawing.Size(14, 14);
            this.chkMerchPrevNoFacDG.TabIndex = 10;
            //
            // chkWood4PrevNoFacDG
            //
            this.chkWood4PrevNoFacDG.Location = new System.Drawing.Point(544, 87);
            this.chkWood4PrevNoFacDG.Name = "chkWood4PrevNoFacDG";
            this.chkWood4PrevNoFacDG.Size = new System.Drawing.Size(14, 14);
            this.chkWood4PrevNoFacDG.TabIndex = 11;
            //
            // chkWood5PrevNoFacDG
            //
            this.chkWood5PrevNoFacDG.Location = new System.Drawing.Point(544, 107);
            this.chkWood5PrevNoFacDG.Name = "chkWood5PrevNoFacDG";
            this.chkWood5PrevNoFacDG.Size = new System.Drawing.Size(14, 14);
            this.chkWood5PrevNoFacDG.TabIndex = 12;
            //
            // chkWood6PrevNoFacDG
            //
            this.chkWood6PrevNoFacDG.Location = new System.Drawing.Point(544, 127);
            this.chkWood6PrevNoFacDG.Name = "chkWood6PrevNoFacDG";
            this.chkWood6PrevNoFacDG.Size = new System.Drawing.Size(14, 14);
            this.chkWood6PrevNoFacDG.TabIndex = 13;
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
            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oRevenue.PreventMerchEconomicDowngrade)
            {
                chkMerchPrevEconDG.CheckState = CheckState.Checked;
            }
            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oRevenue.PreventWood4EconomicDowngrade)
            {
                chkWood4PrevEconDG.CheckState = CheckState.Checked;
            }
            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oRevenue.PreventWood5EconomicDowngrade)
            {
                chkWood5PrevEconDG.CheckState = CheckState.Checked;
            }
            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oRevenue.PreventWood6EconomicDowngrade)
            {
                chkWood6PrevEconDG.CheckState = CheckState.Checked;
            }
            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oRevenue.PreventChipsEconomicDowngrade)
            {
                chkChipsPrevEconDG.CheckState = CheckState.Checked;
            }
            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oRevenue.PreventMerchNoFacilityDowngrade)
            {
                chkMerchPrevNoFacDG.CheckState = CheckState.Checked;
            }
            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oRevenue.PreventWood4NoFacilityDowngrade)
            {
                chkWood4PrevNoFacDG.CheckState = CheckState.Checked;
            }
            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oRevenue.PreventWood5NoFacilityDowngrade)
            {
                chkWood5PrevNoFacDG.CheckState = CheckState.Checked;
            }
            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oRevenue.PreventWood6NoFacilityDowngrade)
            {
                chkWood6PrevNoFacDG.CheckState = CheckState.Checked;
            }

        }

        public int savevalues()
		{
			int x = 0;

			string strHaulCost;
			string strRailHaulCost;
			string strRailBioTransferCost;
			string strRailMerchTransferCost;

			strHaulCost = RoadHaulCostDollarsPerGreenTonPerHour.Replace("$", "");
			strHaulCost = strHaulCost.Replace(",", "");
			if (strHaulCost.Trim().Length == 1) strHaulCost = "0.00";

			strRailHaulCost = RailHaulCostDollarsPerGreenTonPerMile.Replace("$", "");
			strRailHaulCost = strRailHaulCost.Replace(",", "");
			if (strRailHaulCost.Trim().Length == 1) strRailHaulCost = "0.00";

			strRailBioTransferCost = RailChipTransferCostDollarsPerGreenTonPerHour.Replace("$", "");
			strRailBioTransferCost = strRailBioTransferCost.Replace(",", "");
			if (strRailBioTransferCost.Trim().Length == 1) strRailBioTransferCost = "0.00";

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

				//delete all records from the scenario costs table
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

                // default value for all checkbox columns are 'N', so they only need to be updated if they are checked
                if (chkMerchPrevEconDG.Checked)
                {
                    oDataMgr.m_strSQL = "UPDATE scenario_costs SET merch_prev_econDG_YN = 'Y'";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (chkWood4PrevEconDG.Checked)
                {
                    oDataMgr.m_strSQL = "UPDATE scenario_costs SET wood4_prev_econDG_YN = 'Y'";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (chkWood5PrevEconDG.Checked)
                {
                    oDataMgr.m_strSQL = "UPDATE scenario_costs SET wood5_prev_econDG_YN = 'Y'";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (chkWood6PrevEconDG.Checked)
                {
                    oDataMgr.m_strSQL = "UPDATE scenario_costs SET wood6_prev_econDG_YN = 'Y'";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (chkChipsPrevEconDG.Checked)
                {
                    oDataMgr.m_strSQL = "UPDATE scenario_costs SET chips_prev_econDG_YN = 'Y'";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (chkMerchPrevNoFacDG.Checked)
                {
                    oDataMgr.m_strSQL = "UPDATE scenario_costs SET merch_prev_nofacDG_YN = 'Y'";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (chkWood4PrevNoFacDG.Checked)
                {
                    oDataMgr.m_strSQL = "UPDATE scenario_costs SET wood4_prev_nofacDG_YN = 'Y'";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (chkWood5PrevNoFacDG.Checked)
                {
                    oDataMgr.m_strSQL = "UPDATE scenario_costs SET wood5_prev_nofacDG_YN = 'Y'";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (chkWood6PrevNoFacDG.Checked)
                {
                    oDataMgr.m_strSQL = "UPDATE scenario_costs SET wood6_prev_nofacDG_YN = 'Y'";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
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

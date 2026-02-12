using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_RxPackage.
	/// </summary>
	public class uc_rx_package_edit : System.Windows.Forms.UserControl
	{
		public int m_intError=0;
		public string m_strError="";
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		public System.Windows.Forms.TextBox txtPackageDesc;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		public FIA_Biosum_Manager.ResizeFormUsingVisibleScrollBars m_oResizeForm = new ResizeFormUsingVisibleScrollBars();
		private FIA_Biosum_Manager.frmRxPackageItem _frmRxPackageItem;
		private System.Windows.Forms.ComboBox cmbRxPackageId;
		private System.Windows.Forms.RadioButton rdo10YearCycle;
		private System.Windows.Forms.RadioButton rdo5YearCycle;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox grpboxFVSCycle;
        public System.Windows.Forms.CheckBox chkFVSCycleSkip;
		public System.Windows.Forms.TextBox txtRxDesc;
		private System.Windows.Forms.ListView lstRx;
		private System.Windows.Forms.ColumnHeader colYr;
		private System.Windows.Forms.ColumnHeader colRx;
		private System.Windows.Forms.ColumnHeader colDesc;
		private int m_intCurrIndex=-1;
		private System.Windows.Forms.Button btnFVSCycleOk;
		private System.Windows.Forms.Button btnFVSCycleCancel;
		private System.Windows.Forms.Button btnFVSCycleSelectRx;
		private System.Windows.Forms.Button btnFVSCycleEdit;
		private System.Windows.Forms.GroupBox grpboxFVSCycleLength;
		private System.Windows.Forms.ComboBox cmbRx;
		private FIA_Biosum_Manager.Queries m_oQueries = new Queries();
		private RxTools m_oRxTools = new RxTools();
		private string _strRxPackageId="";
        public bool m_bSave = false;

		const int COLUMN_CYCLE=0;
        const int COLUMN_RX = 1;
		const int COLUMN_DESC = 2;
		private System.Windows.Forms.Button btnFVSCycleClearAll;
        private System.Windows.Forms.Button btnFVSCycleClear2;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_rx_package_edit()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			m_oResizeForm.ScrollBarParentControl=panel1;
			m_oResizeForm.ResizeWidth=false;
			m_oResizeForm.MaximumHeight = 650;
			

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
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem(new string[] {
            "00",
            "",
            "",
            ""}, -1);
            System.Windows.Forms.ListViewItem listViewItem2 = new System.Windows.Forms.ListViewItem(new string[] {
            "10",
            "",
            "",
            ""}, -1);
            System.Windows.Forms.ListViewItem listViewItem3 = new System.Windows.Forms.ListViewItem(new string[] {
            "20",
            "",
            "",
            ""}, -1);
            System.Windows.Forms.ListViewItem listViewItem4 = new System.Windows.Forms.ListViewItem(new string[] {
            "30",
            "",
            "",
            ""}, -1);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnFVSCycleClear2 = new System.Windows.Forms.Button();
            this.btnFVSCycleClearAll = new System.Windows.Forms.Button();
            this.btnFVSCycleEdit = new System.Windows.Forms.Button();
            this.lstRx = new System.Windows.Forms.ListView();
            this.colYr = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colRx = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colDesc = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.grpboxFVSCycle = new System.Windows.Forms.GroupBox();
            this.cmbRx = new System.Windows.Forms.ComboBox();
            this.btnFVSCycleCancel = new System.Windows.Forms.Button();
            this.btnFVSCycleOk = new System.Windows.Forms.Button();
            this.chkFVSCycleSkip = new System.Windows.Forms.CheckBox();
            this.btnFVSCycleSelectRx = new System.Windows.Forms.Button();
            this.txtRxDesc = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.grpboxFVSCycleLength = new System.Windows.Forms.GroupBox();
            this.rdo5YearCycle = new System.Windows.Forms.RadioButton();
            this.rdo10YearCycle = new System.Windows.Forms.RadioButton();
            this.cmbRxPackageId = new System.Windows.Forms.ComboBox();
            this.txtPackageDesc = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.grpboxFVSCycle.SuspendLayout();
            this.grpboxFVSCycleLength.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(752, 536);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.cmbRxPackageId);
            this.panel1.Controls.Add(this.txtPackageDesc);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 18);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(746, 515);
            this.panel1.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnFVSCycleClear2);
            this.groupBox2.Controls.Add(this.btnFVSCycleClearAll);
            this.groupBox2.Controls.Add(this.btnFVSCycleEdit);
            this.groupBox2.Controls.Add(this.lstRx);
            this.groupBox2.Controls.Add(this.grpboxFVSCycle);
            this.groupBox2.Controls.Add(this.grpboxFVSCycleLength);
            this.groupBox2.Location = new System.Drawing.Point(16, 112);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(720, 256);
            this.groupBox2.TabIndex = 31;
            this.groupBox2.TabStop = false;
            this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
            // 
            // btnFVSCycleClear2
            // 
            this.btnFVSCycleClear2.Location = new System.Drawing.Point(184, 224);
            this.btnFVSCycleClear2.Name = "btnFVSCycleClear2";
            this.btnFVSCycleClear2.Size = new System.Drawing.Size(64, 25);
            this.btnFVSCycleClear2.TabIndex = 33;
            this.btnFVSCycleClear2.Text = "Clear";
            this.btnFVSCycleClear2.Click += new System.EventHandler(this.btnFVSCycleClear2_Click);
            // 
            // btnFVSCycleClearAll
            // 
            this.btnFVSCycleClearAll.Location = new System.Drawing.Point(248, 224);
            this.btnFVSCycleClearAll.Name = "btnFVSCycleClearAll";
            this.btnFVSCycleClearAll.Size = new System.Drawing.Size(64, 25);
            this.btnFVSCycleClearAll.TabIndex = 32;
            this.btnFVSCycleClearAll.Text = "Clear All";
            this.btnFVSCycleClearAll.Click += new System.EventHandler(this.btnFVSCycleClearAll_Click);
            // 
            // btnFVSCycleEdit
            // 
            this.btnFVSCycleEdit.Enabled = false;
            this.btnFVSCycleEdit.Location = new System.Drawing.Point(16, 224);
            this.btnFVSCycleEdit.Name = "btnFVSCycleEdit";
            this.btnFVSCycleEdit.Size = new System.Drawing.Size(64, 25);
            this.btnFVSCycleEdit.TabIndex = 31;
            this.btnFVSCycleEdit.Text = "Edit";
            this.btnFVSCycleEdit.Click += new System.EventHandler(this.btnFVSCycleEdit_Click);
            // 
            // lstRx
            // 
            this.lstRx.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colYr,
            this.colRx,
            this.colDesc});
            this.lstRx.FullRowSelect = true;
            this.lstRx.GridLines = true;
            this.lstRx.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstRx.HideSelection = false;
            this.lstRx.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1,
            listViewItem2,
            listViewItem3,
            listViewItem4});
            this.lstRx.Location = new System.Drawing.Point(16, 64);
            this.lstRx.MultiSelect = false;
            this.lstRx.Name = "lstRx";
            this.lstRx.Size = new System.Drawing.Size(312, 152);
            this.lstRx.TabIndex = 21;
            this.lstRx.UseCompatibleStateImageBehavior = false;
            this.lstRx.View = System.Windows.Forms.View.Details;
            this.lstRx.SelectedIndexChanged += new System.EventHandler(this.lstRx_SelectedIndexChanged);
            // 
            // colYr
            // 
            this.colYr.Text = "Year";
            this.colYr.Width = 40;
            // 
            // colRx
            // 
            this.colRx.Text = "Treatment";
            this.colRx.Width = 69;
            // 
            // colDesc
            // 
            this.colDesc.Text = "Description";
            this.colDesc.Width = 200;
            // 
            // grpboxFVSCycle
            // 
            this.grpboxFVSCycle.Controls.Add(this.cmbRx);
            this.grpboxFVSCycle.Controls.Add(this.btnFVSCycleCancel);
            this.grpboxFVSCycle.Controls.Add(this.btnFVSCycleOk);
            this.grpboxFVSCycle.Controls.Add(this.chkFVSCycleSkip);
            this.grpboxFVSCycle.Controls.Add(this.btnFVSCycleSelectRx);
            this.grpboxFVSCycle.Controls.Add(this.txtRxDesc);
            this.grpboxFVSCycle.Controls.Add(this.label4);
            this.grpboxFVSCycle.Controls.Add(this.label3);
            this.grpboxFVSCycle.Enabled = false;
            this.grpboxFVSCycle.Location = new System.Drawing.Point(336, 16);
            this.grpboxFVSCycle.Name = "grpboxFVSCycle";
            this.grpboxFVSCycle.Size = new System.Drawing.Size(368, 200);
            this.grpboxFVSCycle.TabIndex = 20;
            this.grpboxFVSCycle.TabStop = false;
            this.grpboxFVSCycle.Text = "Simulation Year";
            // 
            // cmbRx
            // 
            this.cmbRx.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbRx.Location = new System.Drawing.Point(56, 48);
            this.cmbRx.Name = "cmbRx";
            this.cmbRx.Size = new System.Drawing.Size(64, 24);
            this.cmbRx.TabIndex = 21;
            this.cmbRx.SelectedIndexChanged += new System.EventHandler(this.cmbRx_SelectedIndexChanged);
            this.cmbRx.TextChanged += new System.EventHandler(this.cmbRx_TextChanged);
            // 
            // btnFVSCycleCancel
            // 
            this.btnFVSCycleCancel.Location = new System.Drawing.Point(72, 168);
            this.btnFVSCycleCancel.Name = "btnFVSCycleCancel";
            this.btnFVSCycleCancel.Size = new System.Drawing.Size(64, 25);
            this.btnFVSCycleCancel.TabIndex = 20;
            this.btnFVSCycleCancel.Text = "Cancel";
            this.btnFVSCycleCancel.Click += new System.EventHandler(this.btnFVSCycleCancel_Click);
            // 
            // btnFVSCycleOk
            // 
            this.btnFVSCycleOk.Location = new System.Drawing.Point(8, 168);
            this.btnFVSCycleOk.Name = "btnFVSCycleOk";
            this.btnFVSCycleOk.Size = new System.Drawing.Size(64, 25);
            this.btnFVSCycleOk.TabIndex = 19;
            this.btnFVSCycleOk.Text = "OK";
            this.btnFVSCycleOk.Click += new System.EventHandler(this.btnFVSCycleOk_Click);
            // 
            // chkFVSCycleSkip
            // 
            this.chkFVSCycleSkip.Checked = true;
            this.chkFVSCycleSkip.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkFVSCycleSkip.Location = new System.Drawing.Point(8, 16);
            this.chkFVSCycleSkip.Name = "chkFVSCycleSkip";
            this.chkFVSCycleSkip.Size = new System.Drawing.Size(104, 16);
            this.chkFVSCycleSkip.TabIndex = 18;
            this.chkFVSCycleSkip.Text = "Skip Treatment";
            this.chkFVSCycleSkip.CheckedChanged += new System.EventHandler(this.chkFVSCycleSkip_CheckedChanged);
            // 
            // btnFVSCycleSelectRx
            // 
            this.btnFVSCycleSelectRx.Location = new System.Drawing.Point(16, 80);
            this.btnFVSCycleSelectRx.Name = "btnFVSCycleSelectRx";
            this.btnFVSCycleSelectRx.Size = new System.Drawing.Size(104, 25);
            this.btnFVSCycleSelectRx.TabIndex = 10;
            this.btnFVSCycleSelectRx.Text = "Select Treatment";
            this.btnFVSCycleSelectRx.Click += new System.EventHandler(this.btnFVSCycleSelectRx_Click);
            // 
            // txtRxDesc
            // 
            this.txtRxDesc.Enabled = false;
            this.txtRxDesc.Location = new System.Drawing.Point(201, 48);
            this.txtRxDesc.Multiline = true;
            this.txtRxDesc.Name = "txtRxDesc";
            this.txtRxDesc.Size = new System.Drawing.Size(152, 136);
            this.txtRxDesc.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(137, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(72, 16);
            this.label4.TabIndex = 8;
            this.label4.Text = "Description:";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(8, 48);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(40, 16);
            this.label3.TabIndex = 6;
            this.label3.Text = "Rx ID:";
            // 
            // grpboxFVSCycleLength
            // 
            this.grpboxFVSCycleLength.Controls.Add(this.rdo5YearCycle);
            this.grpboxFVSCycleLength.Controls.Add(this.rdo10YearCycle);
            this.grpboxFVSCycleLength.Location = new System.Drawing.Point(16, 16);
            this.grpboxFVSCycleLength.Name = "grpboxFVSCycleLength";
            this.grpboxFVSCycleLength.Size = new System.Drawing.Size(312, 40);
            this.grpboxFVSCycleLength.TabIndex = 30;
            this.grpboxFVSCycleLength.TabStop = false;
            this.grpboxFVSCycleLength.Text = "FVS Cycle Length";
            // 
            // rdo5YearCycle
            // 
            this.rdo5YearCycle.Location = new System.Drawing.Point(16, 18);
            this.rdo5YearCycle.Name = "rdo5YearCycle";
            this.rdo5YearCycle.Size = new System.Drawing.Size(96, 16);
            this.rdo5YearCycle.TabIndex = 29;
            this.rdo5YearCycle.Text = "5 Year Cycles";
            this.rdo5YearCycle.CheckedChanged += new System.EventHandler(this.rdo5YearCycle_CheckedChanged);
            // 
            // rdo10YearCycle
            // 
            this.rdo10YearCycle.Checked = true;
            this.rdo10YearCycle.Location = new System.Drawing.Point(136, 18);
            this.rdo10YearCycle.Name = "rdo10YearCycle";
            this.rdo10YearCycle.Size = new System.Drawing.Size(104, 16);
            this.rdo10YearCycle.TabIndex = 28;
            this.rdo10YearCycle.TabStop = true;
            this.rdo10YearCycle.Text = "10 Year Cycles";
            this.rdo10YearCycle.CheckedChanged += new System.EventHandler(this.rdo10YearCycle_CheckedChanged);
            // 
            // cmbRxPackageId
            // 
            this.cmbRxPackageId.Location = new System.Drawing.Point(88, 24);
            this.cmbRxPackageId.Name = "cmbRxPackageId";
            this.cmbRxPackageId.Size = new System.Drawing.Size(96, 24);
            this.cmbRxPackageId.TabIndex = 27;
            this.cmbRxPackageId.SelectedIndexChanged += new System.EventHandler(this.cmbRxPackageId_SelectedIndexChanged);
            // 
            // txtPackageDesc
            // 
            this.txtPackageDesc.Location = new System.Drawing.Point(284, 16);
            this.txtPackageDesc.Multiline = true;
            this.txtPackageDesc.Name = "txtPackageDesc";
            this.txtPackageDesc.Size = new System.Drawing.Size(440, 88);
            this.txtPackageDesc.TabIndex = 19;
            this.txtPackageDesc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPackageDesc_KeyPress);
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(212, 11);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(72, 16);
            this.label2.TabIndex = 18;
            this.label2.Text = "Description:";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(16, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 16);
            this.label1.TabIndex = 16;
            this.label1.Text = "Package ID:";
            // 
            // uc_rx_package_edit
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_rx_package_edit";
            this.Size = new System.Drawing.Size(752, 536);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.grpboxFVSCycle.ResumeLayout(false);
            this.grpboxFVSCycle.PerformLayout();
            this.grpboxFVSCycleLength.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void label4_Click(object sender, System.EventArgs e)
		{
		
		}

		private void textBox2_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{

			this.ParentForm.DialogResult=DialogResult.OK;
			this.ParentForm.Close();
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			
		}

		private void checkBox1_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnClear_Click(object sender, System.EventArgs e)
		{
			this.txtRxDesc.Text = "";
			this.txtRxDesc.Text="";
		}

		private void button8_Click(object sender, System.EventArgs e)
		{
			MessageBox.Show(this.Height.ToString());
		}
		public void loadvalues()
		{

			this.m_oQueries.m_oFvs.LoadDatasource=true;
			this.m_oQueries.LoadDatasources(true);
			//
			//populate category list box
			//
			this.LoadRxComboBox();
			if (this.ReferenceFormRxPackageItem.m_strAction=="new")
			{
				LoadAvailablePackageIdComboBox();
			}
			else
			{
				this.cmbRxPackageId.Text = this.ReferenceFormRxPackageItem.m_oRxPackageItem.RxPackageId;
				this.txtPackageDesc.Text = this.ReferenceFormRxPackageItem.m_oRxPackageItem.Description;
				if (ReferenceFormRxPackageItem.m_oRxPackageItem.RxCycleLength==10)
				{
					this.rdo10YearCycle.Checked=true;
				}
				else
				{
					this.rdo5YearCycle.Checked=true;
				}
				
				this.cmbRxPackageId.Enabled=false;
				LoadSimYearGrid();
			}
		}
		public void savevalues()
		{
			m_intError=0;
			m_strError="";
			if (this.cmbRxPackageId.Text.Trim().Length == 0)
			{
				this.m_intError=-1;
				m_strError="Select a package id";
			}
			if (this.m_intError != 0)
			{
				MessageBox.Show(m_strError,"FIA Biosum");
			}
			this.ReferenceFormRxPackageItem.m_oRxPackageItem.Description = this.txtPackageDesc.Text;
			this.ReferenceFormRxPackageItem.m_oRxPackageItem.RxPackageId = this.cmbRxPackageId.Text;
			for (int x=0;x<=this.lstRx.Items.Count-1;x++)
			{
				switch (x)
				{
					case 0:
						this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear1Rx = this.lstRx.Items[x].SubItems[COLUMN_RX].Text;
						this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear1Fvs = "1";
						break;
					case 1:
						this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear2Rx = this.lstRx.Items[x].SubItems[COLUMN_RX].Text;
						this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear2Fvs = "2";
						break;
					case 2:
						this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear3Rx = this.lstRx.Items[x].SubItems[COLUMN_RX].Text;
						this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear3Fvs = "3";
						break;
					case 3:
						this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear4Rx = this.lstRx.Items[x].SubItems[COLUMN_RX].Text;
						this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear4Fvs = "4";
						break;
				}
				
			}
			this.ReferenceFormRxPackageItem.m_intError=m_intError;
		}
		private void LoadAvailablePackageIdComboBox()
		{
			this.cmbRxPackageId.Items.Clear();

			int x=0;
			int y=0;
			int intMin = 1;
			int intMax = 999;

			string[] strUsedRxPackageIdArray = frmMain.g_oUtils.ConvertListToArray(this.ReferenceFormRxPackageItem.UsedRxPackageList,",");

			for (x=intMin;x<=intMax;x++)
			{
				if (this.ReferenceFormRxPackageItem.UsedRxPackageList.Trim().Length > 0)
				{
					for (y=0;y<=strUsedRxPackageIdArray.Length-1;y++)
					{
						if (Convert.ToInt32(strUsedRxPackageIdArray[y])==x)break;
						
					
					}
					if (y > strUsedRxPackageIdArray.Length-1)
					{
						this.cmbRxPackageId.Items.Add(Convert.ToString(x).PadLeft(3,'0'));
					}
				}
				else
				{
					this.cmbRxPackageId.Items.Add(Convert.ToString(x).PadLeft(3,'0'));
				}
				
			}
		}
		

		private void LoadRxComboBox()
		{
			this.cmbRx.Items.Clear();
			
			if (ReferenceFormRxPackageItem.ReferenceRxItemCollection != null)
			{
				for (int x=0;x<=this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count-1;x++)
				{
					this.cmbRx.Items.Add(this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId);
				}
			}
		}

		private void LoadSimYearGrid()
		{
			int x;
			//
			//1ST SIM YEAR
			//
            if (this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear1Rx != "000")
            {
                this.lstRx.Items[0].SubItems[COLUMN_RX].Text = this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear1Rx;
                if (this.lstRx.Items[0].SubItems[COLUMN_RX].Text.Trim().Length > 0)
                {
                    if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection != null)
                    {
                        for (x = 0; x <= this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count - 1; x++)
                        {
                            if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim() ==
                                lstRx.Items[0].SubItems[COLUMN_RX].Text.Trim())
                            {
                                this.lstRx.Items[0].SubItems[COLUMN_DESC].Text = ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).Description;
                            }
                        }
                    }
                }
            }
			//
			//2ND SIM YEAR
			//
            if (this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear2Rx != "000")
            {
                this.lstRx.Items[1].SubItems[COLUMN_RX].Text = this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear2Rx;
                if (this.lstRx.Items[1].SubItems[COLUMN_RX].Text.Trim().Length > 0)
                {
                    if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection != null)
                    {
                        for (x = 0; x <= this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count - 1; x++)
                        {
                            if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim() ==
                                lstRx.Items[1].SubItems[COLUMN_RX].Text.Trim())
                            {
                                this.lstRx.Items[1].SubItems[COLUMN_DESC].Text = ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).Description;
                            }
                        }
                    }
                }
            }
			//
			//3rd SIM YEAR
			//
            if (this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear3Rx != "000")
            {
                this.lstRx.Items[2].SubItems[COLUMN_RX].Text = this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear3Rx;
                if (this.lstRx.Items[2].SubItems[COLUMN_RX].Text.Trim().Length > 0)
                {
                    if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection != null)
                    {
                        for (x = 0; x <= this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count - 1; x++)
                        {
                            if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim() ==
                                lstRx.Items[2].SubItems[COLUMN_RX].Text.Trim())
                            {
                                this.lstRx.Items[2].SubItems[COLUMN_DESC].Text = ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).Description;
                            }
                        }
                    }
                }
            }
			//
			//4th SIM YEAR
			//
            if (this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear4Rx != "000")
            {
                this.lstRx.Items[3].SubItems[COLUMN_RX].Text = this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear4Rx;
                if (this.lstRx.Items[3].SubItems[COLUMN_RX].Text.Trim().Length > 0)
                {
                    if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection != null)
                    {
                        for (x = 0; x <= this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count - 1; x++)
                        {
                            if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim() ==
                                lstRx.Items[3].SubItems[COLUMN_RX].Text.Trim())
                            {
                                this.lstRx.Items[3].SubItems[COLUMN_DESC].Text = ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).Description;
                            }
                        }
                    }
                }
            }

		}
		private void UpdateFVSCommandsListBox()
		{
			int x;
			this.lstRx.Items[0].SubItems[COLUMN_RX].Text = this.ReferenceFormRxPackageItem.m_oRxPackageItem.SimulationYear1Rx;
			if (this.lstRx.Items[0].SubItems[COLUMN_RX].Text.Trim().Length > 0)
			{
				if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection != null)
				{
					for (x=0;x<=this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count-1;x++)
					{
						if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim()==
							lstRx.Items[0].SubItems[COLUMN_RX].Text.Trim())
						{
							this.lstRx.Items[0].SubItems[COLUMN_DESC].Text = ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).Description;
						}
					}
				}
			}
		}
		private void chkFVSCycleSkip_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.chkFVSCycleSkip.Checked)
			{
				
				this.btnFVSCycleSelectRx.Enabled=false;
			    this.cmbRx.Text="";
				this.cmbRx.Enabled=false;
				this.txtRxDesc.Text="";
			}
			else
			{
				cmbRx.Enabled=true;
				this.btnFVSCycleSelectRx.Enabled=true;
			}
			

		}

		private void lstRx_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			UpdateSimulation();
		}
		private void UpdateSimulation()
		{
			if (this.lstRx.SelectedItems.Count==0) return;
			this.grpboxFVSCycle.Text = "Simulation Year " + this.lstRx.SelectedItems[0].SubItems[0].Text.Trim();
			this.txtRxDesc.Text = "";
			this.cmbRx.Text = "";
			if (this.lstRx.SelectedItems[0].SubItems[1].Text.Trim().Length==0)
			{
				this.chkFVSCycleSkip.Checked=true;
				//GP this.chkGP.Checked=false;
			}
			else
			{
				
				//GP if (this.lstRx.SelectedItems[0].SubItems[1].Text.Trim()=="GP")
				//GP {
					
				//GP 	this.chkGP.Checked=true;
				//GP 	this.cmbRx.Enabled=false;
				//GP }
				//GP else
				//GP {
					this.chkFVSCycleSkip.Checked=false;
				//GP }
				
				
				this.txtRxDesc.Text = this.lstRx.SelectedItems[0].SubItems[2].Text.Trim();
				if (lstRx.SelectedItems[0].SubItems[1].Text.Trim().Length > 0)
						this.cmbRx.Text = this.lstRx.SelectedItems[0].SubItems[1].Text.Trim();
			}

			this.m_intCurrIndex=this.lstRx.SelectedItems[0].Index;

			if (this.lstRx.SelectedItems.Count==0) return;
			
			this.btnFVSCycleEdit.Enabled=true;
            this.btnFVSCycleClearAll.Enabled = true;
			

		}





        private void btnFVSCycleOk_Click(object sender, System.EventArgs e)
		{
			val_rx();
			if (this.m_intError==0)
			{
				bool bCheck=false;
				string strCurRx = this.lstRx.Items[this.m_intCurrIndex].SubItems[COLUMN_RX].Text.Trim();
				string strCurFvsCycle="";
				//GP if (strCurRx.Trim().ToUpper()=="GP")
				//GP {
				//GP	strCurRx="";
				//GP }
				//GP else
				//GP {
					bCheck=true;
					strCurFvsCycle = Convert.ToString(this.lstRx.Items[this.m_intCurrIndex].Index + 1);
				//GP }

				this.grpboxFVSCycle.Enabled=false;
				this.lstRx.Enabled=true;



			    
				this.lstRx.Items[this.m_intCurrIndex].SubItems[1].Text="";
				this.lstRx.Items[this.m_intCurrIndex].SubItems[2].Text="";
				

				if ((this.chkFVSCycleSkip.Checked || this.chkFVSCycleSkip.Checked) && bCheck)
				{
                    this.ReferenceFormRxPackageItem.uc_rx_package_harvest_cost_column_list1.RemoveRxItemsFromList(strCurRx, strCurFvsCycle);
                    
				}
				else
				{
				    if (bCheck)
					{
						if (strCurRx.Trim().Length > 0 && strCurRx.Trim() != this.cmbRx.Text.Trim())
						{
                            this.ReferenceFormRxPackageItem.uc_rx_package_harvest_cost_column_list1.RemoveRxItemsFromList(strCurRx, strCurFvsCycle);
                            this.ReferenceFormRxPackageItem.uc_rx_package_harvest_cost_column_list1.AddRxItemsToList(cmbRx.Text, strCurFvsCycle);
						}
						else
						{
                            this.ReferenceFormRxPackageItem.uc_rx_package_harvest_cost_column_list1.AddRxItemsToList(cmbRx.Text, strCurFvsCycle);
						}
						
					}
				}
			
				if (this.chkFVSCycleSkip.Checked)
				{

					//GP if (this.chkGP.Checked)
					//GP {
					//GP	this.lstRx.Items[this.m_intCurrIndex].SubItems[1].Text="GP";
						this.lstRx.Items[this.m_intCurrIndex].SubItems[2].Text= "";
					//GP}
				}
				else
				{
					//GPif (this.chkGP.Checked)
					//GP{
					//GP	this.lstRx.Items[this.m_intCurrIndex].SubItems[1].Text="GP";
					//GP	this.lstRx.Items[this.m_intCurrIndex].SubItems[2].Text= "";
					//GP}
					//GPelse
					//GP{
						this.lstRx.Items[this.m_intCurrIndex].SubItems[1].Text=cmbRx.Text;
						this.lstRx.Items[this.m_intCurrIndex].SubItems[2].Text= this.txtRxDesc.Text;
					//}
				}
				this.btnFVSCycleEdit.Enabled=true;

				this.btnFVSCycleClear2.Enabled=true;
				this.btnFVSCycleClearAll.Enabled=true;
                this.m_bSave = true;
			}

			
			

		}
		private void val_rx()
		{
			m_intError=0;
			m_strError="";
			
			if ((this.cmbRx.Text.Trim().Length == 0 && this.chkFVSCycleSkip.Checked==false))
			{
				MessageBox.Show("Select Treatment","FIA Biosum");
				m_intError=-1;
				return;
		
			}
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnFVSCycleEdit_Click(object sender, System.EventArgs e)
		{
			this.lstRx.Enabled=false;
			this.btnFVSCycleEdit.Enabled=false;
			//GP this.btnFVSCycleGP.Enabled=false;
			this.btnFVSCycleClear2.Enabled=false;
			this.btnFVSCycleClearAll.Enabled=false;
		    this.grpboxFVSCycle.Enabled=true;
			if (this.chkFVSCycleSkip.Checked)
			{
				this.cmbRx.Enabled=false;
				this.btnFVSCycleSelectRx.Enabled=false;
			}
			else
			{
				this.btnFVSCycleSelectRx.Enabled=true;
			}
		}

		private void btnFVSCycleCancel_Click(object sender, System.EventArgs e)
		{
			this.grpboxFVSCycle.Enabled=false;
			this.lstRx.Enabled=true;
			UpdateSimulation();


		}

		

		private void cmbRx_TextChanged(object sender, System.EventArgs e)
		{
			if (this.cmbRx.Text.Trim().Length > 0)
			{
				//GP if (this.cmbRx.Text.Trim()=="GP")
				//GP {
				//GP this.chkGP.Checked=true;
				//GP }
				//GP else
				//GP{
					//GP this.chkGP.Checked=false;
					if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection != null)
					{
						for (int x=0;x<=this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count-1;x++)
						{
							if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim()==
								cmbRx.Text.Trim())
							{
								this.txtRxDesc.Text = ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).Description;
							}
						}
					}
				//GP }
			}
		}

		private void btnFVSCycleSelectRx_Click(object sender, System.EventArgs e)
		{
			if (ReferenceFormRxPackageItem.ReferenceRxItemCollection != null)
			{
				FIA_Biosum_Manager.frmDialog frmtemp = new frmDialog();
				frmtemp.uc_previous_expressions1.LoadRxItemCollection(ReferenceFormRxPackageItem.ReferenceRxItemCollection);
				frmtemp.ShowDialog();
				if (frmtemp.DialogResult == DialogResult.OK)
				{
					this.cmbRx.Text = frmtemp.uc_previous_expressions1.listView1.SelectedItems[0].SubItems[1].Text.Trim();
				}
			}
		}

		private void rdo5YearCycle_CheckedChanged(object sender, System.EventArgs e)
		{
			if (rdo5YearCycle.Checked)
			{
				ReferenceFormRxPackageItem.m_oRxPackageItem.RxCycleLength=5;
				this.lstRx.Items[1].Text = "05";
				this.lstRx.Items[2].Text = "10";
				this.lstRx.Items[3].Text = "15";
                this.m_bSave = true;
			}
		}

		private void rdo10YearCycle_CheckedChanged(object sender, System.EventArgs e)
		{
			if (rdo10YearCycle.Checked)
			{
				ReferenceFormRxPackageItem.m_oRxPackageItem.RxCycleLength=10;
				this.lstRx.Items[1].Text = "10";
				this.lstRx.Items[2].Text = "20";
				this.lstRx.Items[3].Text = "30";
                this.m_bSave = true;
			}
		}

		private void cmbRxPackageId_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		}

		

		private void btnFVSCycleClearAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=this.lstRx.Items.Count-1;x++)
			{
				if (this.lstRx.Items[x].SubItems[COLUMN_RX].Text.Trim().Length > 0)
				{
                    this.ReferenceFormRxPackageItem.uc_rx_package_harvest_cost_column_list1.RemoveRxItemsFromList(lstRx.Items[x].SubItems[COLUMN_RX].Text.Trim(),Convert.ToString(lstRx.Items[x].Index + 1));
				}
				this.lstRx.Items[x].SubItems[COLUMN_RX].Text="";
				this.lstRx.Items[x].SubItems[COLUMN_DESC].Text="";
				this.cmbRx.Text="";
				//GP this.chkGP.Checked=false;
				this.chkFVSCycleSkip.Checked=true;
				this.txtRxDesc.Text="";
			}
		}

		private void btnFVSCycleClear2_Click(object sender, System.EventArgs e)
		{
			if (this.lstRx.SelectedItems.Count==0) return;

			if (this.lstRx.SelectedItems[0].SubItems[COLUMN_RX].Text.Trim().Length > 0)
			{
                this.ReferenceFormRxPackageItem.uc_rx_package_harvest_cost_column_list1.RemoveRxItemsFromList(lstRx.SelectedItems[0].SubItems[COLUMN_RX].Text.Trim(), Convert.ToString(lstRx.SelectedItems[0].Index + 1));
			}
			this.lstRx.SelectedItems[0].SubItems[COLUMN_RX].Text="";
			this.lstRx.SelectedItems[0].SubItems[COLUMN_DESC].Text="";
			this.cmbRx.Text="";
			//GP this.chkGP.Checked=false;
			this.chkFVSCycleSkip.Checked=true;
			this.txtRxDesc.Text="";
		}

        private void txtPackageDesc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            m_bSave = true;
        }

		public string RxPackageId
		{
			set {_strRxPackageId=this.cmbRxPackageId.Text;}
			get {return this.cmbRxPackageId.Text;}
		}
		public FIA_Biosum_Manager.frmRxPackageItem ReferenceFormRxPackageItem
		{
			get {return this._frmRxPackageItem;}
			set {this._frmRxPackageItem=value;}

		}

        private void cmbRx_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbRx.SelectedIndex > -1 && ! String.IsNullOrEmpty(this.cmbRx.Text))
            {
                txtRxDesc.Clear();
                string strRx = this.cmbRx.Text.Trim();
                foreach (RxItem oItem in this.ReferenceFormRxPackageItem.ReferenceRxItemCollection)
                {
                    if (strRx.Equals(oItem.RxId))
                    {
                        txtRxDesc.Text = oItem.Description;
                    }
                }
            }
        }
	}
}

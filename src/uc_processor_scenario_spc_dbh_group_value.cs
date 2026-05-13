using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_scenario_spc_dbh_group_values.
	/// </summary>
	public class uc_processor_scenario_spc_dbh_group_value : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.TextBox txtSpeciesGroup;
		private System.Windows.Forms.TextBox txtDbhGroup;
		private System.Windows.Forms.TextBox txtChipPct;
		private System.Windows.Forms.TextBox txtMerchValue;
		private System.Windows.Forms.TextBox txtMerchPct;
		private System.Windows.Forms.TextBox txtWood4Value;
		private System.Windows.Forms.TextBox txtWood4Pct;
		private System.Windows.Forms.TextBox txtWood5Value;
		private System.Windows.Forms.TextBox txtWood5Pct;
		private System.Windows.Forms.TextBox txtWood6Value;
		private System.Windows.Forms.TextBox txtWood6Pct;
		private FIA_Biosum_Manager.ValidateNumericValues m_oValidate = new ValidateNumericValues();
		private string m_strCubicFootDollarValueSave="";
        public bool m_bSave = false;
        private CheckBox chkChips;
        private frmProcessorScenario _oFrmProcessorScenario = null;
		
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_processor_scenario_spc_dbh_group_value()
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
		public string CubicFootDollarValue
		{
			get {return this.txtMerchValue.Text.Trim();}
			set {this.txtMerchValue.Text = value;m_strCubicFootDollarValueSave=value;}
		}
		public string DbhGroup
		{
			get {return this.txtDbhGroup.Text.Trim();}
			set {this.txtDbhGroup.Text = value;}
		}
		public string SpeciesGroup
		{
			get {return this.txtSpeciesGroup.Text.Trim();}
			set {this.txtSpeciesGroup.Text = value;}
		}
        public string GetWoodBin()
        {
            if (chkChips.Checked) return "C";
            else return "M";
        }
        public bool EnergyWood
        {
            get { return chkChips.Checked; }
            set { chkChips.Checked = value; }
        }
        public void SetWoodBin(string p_strWoodBin)
        {
            if (p_strWoodBin.Trim() == "C")
            {
                chkChips.Checked = true;
            }
            else
                chkChips.Checked = false;
        }
        public frmProcessorScenario ReferenceProcessorScenarioForm
        {
            get { return _oFrmProcessorScenario; }
            set { _oFrmProcessorScenario = value; }
        }

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.txtSpeciesGroup = new System.Windows.Forms.TextBox();
            this.txtDbhGroup = new System.Windows.Forms.TextBox();
            this.txtMerchValue = new System.Windows.Forms.TextBox();
            this.chkChips = new System.Windows.Forms.CheckBox();
			this.txtChipPct = new System.Windows.Forms.TextBox();
			this.txtMerchPct = new System.Windows.Forms.TextBox();
			this.txtWood4Value = new System.Windows.Forms.TextBox();
			this.txtWood4Pct = new System.Windows.Forms.TextBox();
			this.txtWood5Value = new System.Windows.Forms.TextBox();
			this.txtWood5Pct = new System.Windows.Forms.TextBox();
			this.txtWood6Value = new System.Windows.Forms.TextBox();
			this.txtWood6Pct = new System.Windows.Forms.TextBox();
			this.SuspendLayout();
            // 
            // txtSpeciesGroup
            // 
            this.txtSpeciesGroup.Location = new System.Drawing.Point(8, 6);
            this.txtSpeciesGroup.Name = "txtSpeciesGroup";
            this.txtSpeciesGroup.Size = new System.Drawing.Size(150, 20);
            this.txtSpeciesGroup.TabIndex = 0;
            this.txtSpeciesGroup.Enter += new System.EventHandler(this.txtSpeciesGroup_Enter);
            this.txtSpeciesGroup.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSpeciesGroup_KeyPress);
            // 
            // txtDbhGroup
            // 
            this.txtDbhGroup.Location = new System.Drawing.Point(166, 6);
            this.txtDbhGroup.Name = "txtDbhGroup";
            this.txtDbhGroup.Size = new System.Drawing.Size(66, 20);
            this.txtDbhGroup.TabIndex = 1;
            this.txtDbhGroup.Enter += new System.EventHandler(this.txtDbhGroup_Enter);
            this.txtDbhGroup.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDbhGroup_KeyPress);
			// 
			// chkChips
			// 
			this.chkChips.AutoSize = true;
			this.chkChips.Location = new System.Drawing.Point(265, 9);
			this.chkChips.Name = "chkChips";
			this.chkChips.Size = new System.Drawing.Size(15, 14);
			this.chkChips.TabIndex = 2;
			this.chkChips.UseVisualStyleBackColor = true;
			this.chkChips.CheckedChanged += new System.EventHandler(this.chkChips_CheckedChanged);
			//
			// txtChipPct
			//
			this.txtChipPct.Location = new System.Drawing.Point(323, 6);
			this.txtChipPct.Name = "txtChipPct";
			this.txtChipPct.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtChipPct.ShortcutsEnabled = false;
			this.txtChipPct.Size = new System.Drawing.Size(40, 20);
			this.txtChipPct.TabIndex = 3;
			this.txtChipPct.Text = "0";
			this.txtChipPct.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			// 
			// txtMerchValue
			// 
			this.txtMerchValue.Location = new System.Drawing.Point(381, 6);
            this.txtMerchValue.Name = "txtMerchValue";
            this.txtMerchValue.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtMerchValue.ShortcutsEnabled = false;
            this.txtMerchValue.Size = new System.Drawing.Size(60, 20);
            this.txtMerchValue.TabIndex = 4;
            this.txtMerchValue.Text = "$0.00";
            this.txtMerchValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtMerchValue.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMerchValue_KeyPress);
            this.txtMerchValue.Leave += new System.EventHandler(this.txtValue_Leave);
			//
			// txtMerchPct
			//
			this.txtMerchPct.Location = new System.Drawing.Point(459, 6);
			this.txtMerchPct.Name = "txtMerchPct";
			this.txtMerchPct.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtMerchPct.ShortcutsEnabled = false;
			this.txtMerchPct.Size = new System.Drawing.Size(40, 20);
			this.txtMerchPct.TabIndex = 5;
			this.txtMerchPct.Text = "100";
			this.txtMerchPct.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			//
			// txtWood4Value
			//
			this.txtWood4Value.Location = new System.Drawing.Point(517, 6);
			this.txtWood4Value.Name = "txtWood4Value";
			this.txtWood4Value.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtWood4Value.ShortcutsEnabled = false;
			this.txtWood4Value.Size = new System.Drawing.Size(60, 20);
			this.txtWood4Value.TabIndex = 6;
			this.txtWood4Value.Text = "$0.00";
			this.txtWood4Value.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			//
			// txtWood4Pct
			//
			this.txtWood4Pct.Location = new System.Drawing.Point(595, 6);
			this.txtWood4Pct.Name = "txtWood4Pct";
			this.txtWood4Pct.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtWood4Pct.ShortcutsEnabled = false;
			this.txtWood4Pct.Size = new System.Drawing.Size(40, 20);
			this.txtWood4Pct.TabIndex = 7;
			this.txtWood4Pct.Text = "0";
			this.txtWood4Pct.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			//
			// txtWood5Value
			//
			this.txtWood5Value.Location = new System.Drawing.Point(653, 6);
			this.txtWood5Value.Name = "txtWood5Value";
			this.txtWood5Value.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtWood5Value.ShortcutsEnabled = false;
			this.txtWood5Value.Size = new System.Drawing.Size(60, 20);
			this.txtWood5Value.TabIndex = 8;
			this.txtWood5Value.Text = "$0.00";
			this.txtWood5Value.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			//
			// txtWood5Pct
			//
			this.txtWood5Pct.Location = new System.Drawing.Point(731, 6);
			this.txtWood5Pct.Name = "txtWood5Pct";
			this.txtWood5Pct.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtWood5Pct.ShortcutsEnabled = false;
			this.txtWood5Pct.Size = new System.Drawing.Size(40, 20);
			this.txtWood5Pct.TabIndex = 9;
			this.txtWood5Pct.Text = "0";
			this.txtWood5Pct.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			//
			// txtWood6Value
			//
			this.txtWood6Value.Location = new System.Drawing.Point(789, 6);
			this.txtWood6Value.Name = "txtWood6Value";
			this.txtWood6Value.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtWood6Value.ShortcutsEnabled = false;
			this.txtWood6Value.Size = new System.Drawing.Size(60, 20);
			this.txtWood6Value.TabIndex = 10;
			this.txtWood6Value.Text = "$0.00";
			this.txtWood6Value.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			//
			// txtWood6Pct
			//
			this.txtWood6Pct.Location = new System.Drawing.Point(867, 6);
			this.txtWood6Pct.Name = "txtWood6Pct";
			this.txtWood6Pct.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.txtWood6Pct.ShortcutsEnabled = false;
			this.txtWood6Pct.Size = new System.Drawing.Size(40, 20);
			this.txtWood6Pct.TabIndex = 11;
			this.txtWood6Pct.Text = "0";
			this.txtWood6Pct.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			// 
			// uc_processor_scenario_spc_dbh_group_value
			// 
			this.Controls.Add(this.txtSpeciesGroup);
			this.Controls.Add(this.txtDbhGroup);
			this.Controls.Add(this.chkChips);
			this.Controls.Add(this.txtChipPct);
			this.Controls.Add(this.txtMerchValue);
			this.Controls.Add(this.txtMerchPct);
			this.Controls.Add(this.txtWood4Value);
			this.Controls.Add(this.txtWood4Pct);
			this.Controls.Add(this.txtWood5Value);
			this.Controls.Add(this.txtWood5Pct);
			this.Controls.Add(this.txtWood6Value);
			this.Controls.Add(this.txtWood6Pct);
            this.Name = "uc_processor_scenario_spc_dbh_group_value";
            this.Size = new System.Drawing.Size(915, 32);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		private void txtSpeciesGroup_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void txtDbhGroup_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void txtValue_Leave(object sender, System.EventArgs e)
		{
			m_oValidate.RoundDecimalLength=2;
			m_oValidate.Money=true;
			m_oValidate.NullsAllowed=false;
			m_oValidate.TestForMaxMin=false;
			m_oValidate.TestForMin=true;
			m_oValidate.MinValue=0;
			m_oValidate.ValidateDecimal(txtMerchValue.Text);
            if (m_oValidate.m_intError == 0)
            {
                
                                
                txtMerchValue.Text = m_oValidate.ReturnValue;

            }
            else
            {
                this.txtMerchValue.Text = this.m_strCubicFootDollarValueSave;
                this.txtMerchValue.Focus();

            }


		}
		public void SaveValues()
		{
			this.m_strCubicFootDollarValueSave=this.txtMerchValue.Text;
		}

		private void txtSpeciesGroup_Enter(object sender, System.EventArgs e)
		{
			txtMerchValue.Focus();
		}

		private void txtDbhGroup_Enter(object sender, System.EventArgs e)
		{
			txtMerchValue.Focus();
		}

        private void chkChips_CheckedChanged(object sender, EventArgs e)
        {
            if (this.ReferenceProcessorScenarioForm != null) ReferenceProcessorScenarioForm.m_bSave = true;
			if (chkChips.Checked)
			{
				this.txtChipPct.Text = "100";
				this.txtChipPct.Enabled = false;
				this.txtMerchValue.Enabled = false;
				this.txtMerchPct.Text = "0";
				this.txtMerchPct.Enabled = false;
				this.txtWood4Value.Enabled = false;
				this.txtWood4Pct.Text = "0";
				this.txtWood4Pct.Enabled = false;
				this.txtWood5Value.Enabled = false;
				this.txtWood5Pct.Text = "0";
				this.txtWood5Pct.Enabled = false;
				this.txtWood6Value.Enabled = false;
				this.txtWood6Pct.Text = "0";
				this.txtWood6Pct.Enabled = false;
			}
			else
			{
				this.txtChipPct.Enabled = true;
				this.txtMerchValue.Enabled = true;
				this.txtMerchPct.Enabled = true;
				this.txtWood4Value.Enabled = true;
				this.txtWood4Pct.Enabled = true;
				this.txtWood5Value.Enabled = true;
				this.txtWood5Pct.Enabled = true;
				this.txtWood6Value.Enabled = true;
				this.txtWood6Pct.Enabled = true;
			}
        }

        private void txtMerchValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

		
	}
	public class uc_processor_scenario_spc_dbh_group_value_collection : System.Collections.CollectionBase
	{
		public uc_processor_scenario_spc_dbh_group_value_collection()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public void Add(uc_processor_scenario_spc_dbh_group_value p_uc)
		{
			// vérify if object is not already in
			if (this.List.Contains(p_uc))
				throw new InvalidOperationException();
 
			// adding it
			this.List.Add(p_uc);
 
			// return collection
			//return this;
		}
		public void Remove(int index)
		{
			// Check to see if there is a widget at the supplied index.
			if (index > Count - 1 || index < 0)
				// If no widget exists, a messagebox is shown and the operation 
				// is canColumned.
			{
				System.Windows.Forms.MessageBox.Show("Index not valid!");
			}
			else
			{
				List.RemoveAt(index); 
			}
		}
		public FIA_Biosum_Manager.uc_processor_scenario_spc_dbh_group_value Item(int Index)
		{
			// The appropriate item is retrieved from the List object and
			// explicitly cast to the Widget type, then returned to the 
			// caller.
			return (FIA_Biosum_Manager.uc_processor_scenario_spc_dbh_group_value) List[Index];
		}

	}
}

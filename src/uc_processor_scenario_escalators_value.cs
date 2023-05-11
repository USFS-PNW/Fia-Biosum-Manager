using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_scenario_escalators_values.
	/// </summary>
	public class uc_processor_scenario_escalators_value : System.Windows.Forms.UserControl
	{
		private FIA_Biosum_Manager.ValidateNumericValues m_oValidate = new ValidateNumericValues();
		private System.Windows.Forms.TextBox txtCycle2;
		private System.Windows.Forms.TextBox txtCycle3;
		private System.Windows.Forms.TextBox txtCycle4;
		private string m_strCycle2Save="1.00";
		private string m_strCycle3Save="1.00";
		private string m_strCycle4Save="1.00";
        private frmProcessorScenario _oFrmProcessorScenario = null;

		//protected override bool ProcessDialogKey(Keys keyData); 



	
		
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_processor_scenario_escalators_value()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			m_oValidate.RoundDecimalLength=2;
			m_oValidate.Money=false;
			m_oValidate.NullsAllowed=false;
			m_oValidate.TestForMaxMin=true;
			m_oValidate.MaxValue=1.99;
			m_oValidate.MinValue=-1.99;

			// TODO: Add any initialization after the InitializeComponent call

		}
		public void SaveValues()
		{
			m_strCycle2Save=Cycle2;
			m_strCycle3Save=Cycle3;
			m_strCycle4Save=Cycle4;

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
            get { return _oFrmProcessorScenario; }
            set { _oFrmProcessorScenario = value; }
        }
		public string Cycle2
		{
			get {return this.txtCycle2.Text.Trim();}
			set {this.txtCycle2.Text = value; this.m_strCycle2Save=value;}
		}
		public string Cycle3
		{
			get {return this.txtCycle3.Text.Trim();}
			set {this.txtCycle3.Text = value; m_strCycle3Save=value;}
		}
		public string Cycle4
		{
			get {return this.txtCycle4.Text.Trim();}
			set {this.txtCycle4.Text = value;m_strCycle4Save=value;}
		}
		
		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.txtCycle2 = new System.Windows.Forms.TextBox();
            this.txtCycle3 = new System.Windows.Forms.TextBox();
            this.txtCycle4 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txtCycle2
            // 
            this.txtCycle2.Location = new System.Drawing.Point(8, 6);
            this.txtCycle2.MaxLength = 5;
            this.txtCycle2.Name = "txtCycle2";
            this.txtCycle2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtCycle2.Size = new System.Drawing.Size(112, 22);
            this.txtCycle2.TabIndex = 0;
            this.txtCycle2.Text = "1.00";
            this.txtCycle2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtCycle2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCycle1_KeyPress);
            this.txtCycle2.Leave += new System.EventHandler(this.txtCycle1_Leave);
            // 
            // txtCycle3
            // 
            this.txtCycle3.Location = new System.Drawing.Point(136, 6);
            this.txtCycle3.MaxLength = 5;
            this.txtCycle3.Name = "txtCycle3";
            this.txtCycle3.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtCycle3.Size = new System.Drawing.Size(112, 22);
            this.txtCycle3.TabIndex = 1;
            this.txtCycle3.Text = "1.00";
            this.txtCycle3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtCycle3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCycle2_KeyPress);
            this.txtCycle3.Leave += new System.EventHandler(this.txtCycle2_Leave);
            // 
            // txtCycle4
            // 
            this.txtCycle4.Location = new System.Drawing.Point(264, 6);
            this.txtCycle4.MaxLength = 5;
            this.txtCycle4.Name = "txtCycle4";
            this.txtCycle4.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtCycle4.Size = new System.Drawing.Size(112, 22);
            this.txtCycle4.TabIndex = 2;
            this.txtCycle4.Text = "1.00";
            this.txtCycle4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtCycle4.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCycle3_KeyPress);
            this.txtCycle4.Leave += new System.EventHandler(this.txtCycle3_Leave);
            // 
            // uc_processor_scenario_escalators_value
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.Controls.Add(this.txtCycle4);
            this.Controls.Add(this.txtCycle3);
            this.Controls.Add(this.txtCycle2);
            this.Name = "uc_processor_scenario_escalators_value";
            this.Size = new System.Drawing.Size(384, 32);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		
		

		
		private void ValidateValue(System.Windows.Forms.TextBox p_txtValue)
		{
			m_oValidate.ValidateDecimal(p_txtValue.Text.Trim());
			if (m_oValidate.m_intError==0)
				p_txtValue.Text  = m_oValidate.ReturnValue;
			else
			{
				if (p_txtValue.Name.Trim().ToUpper() =="TXTCYCLE1")
					p_txtValue.Text = m_strCycle2Save;
				else if (p_txtValue.Name.Trim().ToUpper() =="TXTCYCLE2")
					p_txtValue.Text = m_strCycle3Save;
				else if (p_txtValue.Name.Trim().ToUpper() =="TXTCYCLE3")
					p_txtValue.Text = m_strCycle4Save;
				p_txtValue.Focus();
			}
		}

		private void txtCycle3_Leave(object sender, System.EventArgs e)
		{
			ValidateValue(txtCycle4);
			
		}

		private void txtCycle2_Leave(object sender, System.EventArgs e)
		{
			ValidateValue(txtCycle3);
		}

		private void txtCycle1_Leave(object sender, System.EventArgs e)
		{
			
			ValidateValue(txtCycle2);
			
		}

        private void txtCycle1_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

        private void txtCycle2_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

        private void txtCycle3_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }



					
		





		
	}
	
}

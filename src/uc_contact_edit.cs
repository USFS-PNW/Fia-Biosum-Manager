using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_contact_edit.
	/// </summary>
	public class uc_contact_edit : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnOK;
		private int m_intError=0;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.CheckBox chkOptimizer;
		private System.Windows.Forms.CheckBox chkFvs;
		private System.Windows.Forms.CheckBox chkProcessor;
		private System.Windows.Forms.CheckBox chkOpcost;
		private System.Windows.Forms.CheckBox chkGis;
		private System.Windows.Forms.TextBox txtEmail;
		private System.Windows.Forms.TextBox txtAreaCode;
		private System.Windows.Forms.TextBox txtCity;
		private System.Windows.Forms.TextBox txtStreet;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.ComboBox cmbState;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox txtZip1;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox txtZip2;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.TextBox txtPrefix;
		private System.Windows.Forms.TextBox txtPhone;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.TextBox txtExt;
		private System.Windows.Forms.TextBox txtName;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.TextBox txtOrg;


		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_contact_edit(ado_data_access p_ado,string p_strTreeSpcTable)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.txtPhone.Text="";
			this.txtAreaCode.Text="";
			this.txtPrefix.Text="";
			this.txtExt.Text="";
			this.txtZip1.Text="";
			this.txtZip2.Text="";
			this.txtStreet.Text="";
			this.txtOrg.Text="";
			this.populate_state_combo();



			


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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtOrg = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txtExt = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtPhone = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtPrefix = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtZip2 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtZip1 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbState = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkGis = new System.Windows.Forms.CheckBox();
            this.chkOpcost = new System.Windows.Forms.CheckBox();
            this.chkProcessor = new System.Windows.Forms.CheckBox();
            this.chkFvs = new System.Windows.Forms.CheckBox();
            this.chkOptimizer = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtEmail = new System.Windows.Forms.TextBox();
            this.txtAreaCode = new System.Windows.Forms.TextBox();
            this.txtCity = new System.Windows.Forms.TextBox();
            this.txtStreet = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.txtName = new System.Windows.Forms.TextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtOrg);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.txtExt);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.txtPhone);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.txtPrefix);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.txtZip2);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.txtZip1);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.cmbState);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtEmail);
            this.groupBox1.Controls.Add(this.txtAreaCode);
            this.groupBox1.Controls.Add(this.txtCity);
            this.groupBox1.Controls.Add(this.txtStreet);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.txtName);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnOK);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.groupBox1.Size = new System.Drawing.Size(624, 592);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // txtOrg
            // 
            this.txtOrg.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOrg.Location = new System.Drawing.Point(219, 88);
            this.txtOrg.MaxLength = 75;
            this.txtOrg.Name = "txtOrg";
            this.txtOrg.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtOrg.Size = new System.Drawing.Size(352, 23);
            this.txtOrg.TabIndex = 1;
            // 
            // label13
            // 
            this.label13.Location = new System.Drawing.Point(43, 88);
            this.label13.Name = "label13";
            this.label13.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label13.Size = new System.Drawing.Size(152, 16);
            this.label13.TabIndex = 43;
            this.label13.Text = "Organization";
            // 
            // txtExt
            // 
            this.txtExt.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtExt.Location = new System.Drawing.Point(500, 440);
            this.txtExt.MaxLength = 4;
            this.txtExt.Name = "txtExt";
            this.txtExt.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtExt.Size = new System.Drawing.Size(48, 23);
            this.txtExt.TabIndex = 11;
            this.txtExt.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtExt_KeyPress);
            // 
            // label9
            // 
            this.label9.Location = new System.Drawing.Point(452, 444);
            this.label9.Name = "label9";
            this.label9.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.label9.Size = new System.Drawing.Size(40, 16);
            this.label9.TabIndex = 41;
            this.label9.Text = "Ext.";
            // 
            // txtPhone
            // 
            this.txtPhone.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPhone.Location = new System.Drawing.Point(380, 440);
            this.txtPhone.MaxLength = 4;
            this.txtPhone.Name = "txtPhone";
            this.txtPhone.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtPhone.Size = new System.Drawing.Size(56, 23);
            this.txtPhone.TabIndex = 10;
            this.txtPhone.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPhone_KeyPress);
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(356, 444);
            this.label8.Name = "label8";
            this.label8.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.label8.Size = new System.Drawing.Size(16, 16);
            this.label8.TabIndex = 39;
            this.label8.Text = "-";
            // 
            // txtPrefix
            // 
            this.txtPrefix.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPrefix.Location = new System.Drawing.Point(292, 440);
            this.txtPrefix.MaxLength = 3;
            this.txtPrefix.Name = "txtPrefix";
            this.txtPrefix.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtPrefix.Size = new System.Drawing.Size(48, 23);
            this.txtPrefix.TabIndex = 9;
            this.txtPrefix.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPrefix_KeyPress);
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(278, 444);
            this.label7.Name = "label7";
            this.label7.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.label7.Size = new System.Drawing.Size(16, 16);
            this.label7.TabIndex = 37;
            this.label7.Text = "-";
            // 
            // txtZip2
            // 
            this.txtZip2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtZip2.Location = new System.Drawing.Point(324, 408);
            this.txtZip2.MaxLength = 4;
            this.txtZip2.Name = "txtZip2";
            this.txtZip2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtZip2.Size = new System.Drawing.Size(56, 23);
            this.txtZip2.TabIndex = 7;
            this.txtZip2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtZip2_KeyPress);
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(300, 408);
            this.label6.Name = "label6";
            this.label6.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.label6.Size = new System.Drawing.Size(16, 16);
            this.label6.TabIndex = 35;
            this.label6.Text = "-";
            // 
            // txtZip1
            // 
            this.txtZip1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtZip1.Location = new System.Drawing.Point(221, 408);
            this.txtZip1.MaxLength = 5;
            this.txtZip1.Name = "txtZip1";
            this.txtZip1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtZip1.Size = new System.Drawing.Size(64, 23);
            this.txtZip1.TabIndex = 6;
            this.txtZip1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtZip1_KeyPress);
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(52, 408);
            this.label4.Name = "label4";
            this.label4.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label4.Size = new System.Drawing.Size(160, 16);
            this.label4.TabIndex = 33;
            this.label4.Text = "Zip Code";
            // 
            // cmbState
            // 
            this.cmbState.Location = new System.Drawing.Point(220, 376);
            this.cmbState.Name = "cmbState";
            this.cmbState.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmbState.Size = new System.Drawing.Size(232, 24);
            this.cmbState.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(52, 376);
            this.label2.Name = "label2";
            this.label2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label2.Size = new System.Drawing.Size(160, 16);
            this.label2.TabIndex = 31;
            this.label2.Text = "State";
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox2.Controls.Add(this.chkGis);
            this.groupBox2.Controls.Add(this.chkOpcost);
            this.groupBox2.Controls.Add(this.chkProcessor);
            this.groupBox2.Controls.Add(this.chkFvs);
            this.groupBox2.Controls.Add(this.chkOptimizer);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(220, 112);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.groupBox2.Size = new System.Drawing.Size(352, 192);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Select One Or More User Associated Processes";
            // 
            // chkGis
            // 
            this.chkGis.Location = new System.Drawing.Point(16, 152);
            this.chkGis.Name = "chkGis";
            this.chkGis.Size = new System.Drawing.Size(320, 24);
            this.chkGis.TabIndex = 4;
            this.chkGis.Text = "Maps And Travel Times (GIS)";
            // 
            // chkOpcost
            // 
            this.chkOpcost.Location = new System.Drawing.Point(16, 118);
            this.chkOpcost.Name = "chkOpcost";
            this.chkOpcost.Size = new System.Drawing.Size(320, 24);
            this.chkOpcost.TabIndex = 3;
            this.chkOpcost.Text = "Wood Harvesting And Costs (OpCost)";
            // 
            // chkProcessor
            // 
            this.chkProcessor.Location = new System.Drawing.Point(16, 88);
            this.chkProcessor.Name = "chkProcessor";
            this.chkProcessor.Size = new System.Drawing.Size(320, 24);
            this.chkProcessor.TabIndex = 2;
            this.chkProcessor.Text = "Wood Volumes And Values (Processor)";
            // 
            // chkFvs
            // 
            this.chkFvs.Location = new System.Drawing.Point(16, 56);
            this.chkFvs.Name = "chkFvs";
            this.chkFvs.Size = new System.Drawing.Size(320, 24);
            this.chkFvs.TabIndex = 1;
            this.chkFvs.Text = "Forest Vegetation Simulator (FVS)";
            // 
            // chkOptimizer
            // 
            this.chkOptimizer.Location = new System.Drawing.Point(16, 24);
            this.chkOptimizer.Name = "chkOptimizer";
            this.chkOptimizer.Size = new System.Drawing.Size(320, 24);
            this.chkOptimizer.TabIndex = 0;
            this.chkOptimizer.Text = "Treatment Optimizer";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(44, 120);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label1.Size = new System.Drawing.Size(152, 16);
            this.label1.TabIndex = 29;
            this.label1.Text = "Process";
            // 
            // txtEmail
            // 
            this.txtEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtEmail.Location = new System.Drawing.Point(220, 472);
            this.txtEmail.MaxLength = 50;
            this.txtEmail.Name = "txtEmail";
            this.txtEmail.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtEmail.Size = new System.Drawing.Size(296, 23);
            this.txtEmail.TabIndex = 12;
            // 
            // txtAreaCode
            // 
            this.txtAreaCode.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtAreaCode.Location = new System.Drawing.Point(220, 440);
            this.txtAreaCode.MaxLength = 3;
            this.txtAreaCode.Name = "txtAreaCode";
            this.txtAreaCode.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtAreaCode.Size = new System.Drawing.Size(48, 23);
            this.txtAreaCode.TabIndex = 8;
            this.txtAreaCode.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAreaCode_KeyPress);
            // 
            // txtCity
            // 
            this.txtCity.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCity.Location = new System.Drawing.Point(220, 344);
            this.txtCity.MaxLength = 50;
            this.txtCity.Name = "txtCity";
            this.txtCity.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtCity.Size = new System.Drawing.Size(296, 23);
            this.txtCity.TabIndex = 4;
            // 
            // txtStreet
            // 
            this.txtStreet.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtStreet.Location = new System.Drawing.Point(220, 312);
            this.txtStreet.MaxLength = 30;
            this.txtStreet.Name = "txtStreet";
            this.txtStreet.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtStreet.Size = new System.Drawing.Size(296, 23);
            this.txtStreet.TabIndex = 3;
            // 
            // label12
            // 
            this.label12.Location = new System.Drawing.Point(44, 480);
            this.label12.Name = "label12";
            this.label12.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label12.Size = new System.Drawing.Size(160, 16);
            this.label12.TabIndex = 20;
            this.label12.Text = "Email Address";
            // 
            // label11
            // 
            this.label11.Location = new System.Drawing.Point(52, 443);
            this.label11.Name = "label11";
            this.label11.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label11.Size = new System.Drawing.Size(160, 16);
            this.label11.TabIndex = 19;
            this.label11.Text = "Work Phone";
            // 
            // label10
            // 
            this.label10.Location = new System.Drawing.Point(28, 345);
            this.label10.Name = "label10";
            this.label10.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label10.Size = new System.Drawing.Size(184, 16);
            this.label10.TabIndex = 18;
            this.label10.Text = "City";
            // 
            // txtName
            // 
            this.txtName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtName.Location = new System.Drawing.Point(220, 57);
            this.txtName.MaxLength = 50;
            this.txtName.Name = "txtName";
            this.txtName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtName.Size = new System.Drawing.Size(352, 23);
            this.txtName.TabIndex = 0;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(312, 520);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(88, 48);
            this.btnCancel.TabIndex = 14;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(224, 520);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(88, 48);
            this.btnOK.TabIndex = 13;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(52, 315);
            this.label5.Name = "label5";
            this.label5.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label5.Size = new System.Drawing.Size(160, 16);
            this.label5.TabIndex = 7;
            this.label5.Text = "Street Address";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(44, 57);
            this.label3.Name = "label3";
            this.label3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label3.Size = new System.Drawing.Size(152, 16);
            this.label3.TabIndex = 5;
            this.label3.Text = "Name";
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 19);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblTitle.Size = new System.Drawing.Size(618, 24);
            this.lblTitle.TabIndex = 2;
            this.lblTitle.Text = "Contacts Edit";
            // 
            // uc_contact_edit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_contact_edit";
            this.Size = new System.Drawing.Size(624, 592);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void txtSpCd_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (Char.IsDigit((char)e.KeyValue))
			{
			}
			else
			{
				if (e.KeyCode == Keys.Back)
				{
				}
				else
				{
					e.Handled=true;	
				}
			}
		}

		private void txtConvert_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (Char.IsDigit((char)e.KeyValue))
			{
			}
			else
			{
				if (e.KeyCode == Keys.Back)
				{
				}
				else
				{
					e.Handled=true;	
				}
			}
		}

		private void cmbVariant_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			e.Handled=true;
		}

		private void txtSpCd_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
			
		}

		private void txtConvert_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
		}
		private void val_data()
		{
			//return;

			this.m_intError=0;
			

			if (this.txtName.Text.Trim().Length == 0)
			{
				MessageBox.Show("!!Enter A Name!!","FIA Biosum",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				this.txtName.Focus();
				return;
			}
			if (this.txtAreaCode.Text.Trim().Length > 0 ||
				this.txtPrefix.Text.Trim().Length > 0 || 
				this.txtPhone.Text.Trim().Length  > 0 ||
				this.txtExt.Text.Trim().Length  > 0)
			{
				if (this.txtAreaCode.Text.Trim().Length < 3)
				{
					MessageBox.Show("!!Enter Area Code!!","FIA Biosum",
						System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);
					this.m_intError=-1;
					this.txtAreaCode.Focus();
					return;
				}
				if (this.txtPrefix.Text.Trim().Length < 3)
				{
					MessageBox.Show("!!Enter Phone Prefix!!","FIA Biosum",
						System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);
					this.m_intError=-1;
					this.txtPrefix.Focus();
					return;
				}
				if (this.txtPhone.Text.Trim().Length < 4)
				{
					MessageBox.Show("!!Enter Phone Number!!","FIA Biosum",
						System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);
					this.m_intError=-1;
					this.txtPhone.Focus();
					return;
				}
			}

			if (this.txtZip1.Text.Trim().Length > 0)
			{
				if (this.txtZip1.Text.Trim().Length < 5)
				{
					MessageBox.Show("!!Enter Zip Code!!","FIA Biosum",
						System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);
					this.m_intError=-1;
					this.txtZip1.Focus();
					return;

				}
			}
			if (this.txtZip2.Text.Trim().Length > 0)
			{
				if (this.txtZip2.Text.Trim().Length < 4)
				{
					MessageBox.Show("!!Enter Zip Code!!","FIA Biosum",
						System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);
					this.m_intError=-1;
					this.txtZip2.Focus();
					return;

				}
				if (this.txtZip1.Text.Trim().Length < 5)
				{
					MessageBox.Show("!!Enter Zip Code!!","FIA Biosum",
						System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);
					this.m_intError=-1;
					this.txtZip1.Focus();
					return;

				}
			}




		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.Cancel;
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			this.val_data();
			if (this.m_intError==0)
				((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.OK;
		}

		private void txtZip1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
		}

		private void txtZip2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
		}

		private void txtAreaCode_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
		}

		private void txtPrefix_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
		}

		private void txtPhone_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
		}

		private void txtExt_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
		}
		private void populate_state_combo()
		{
			this.cmbState.Items.Clear();
			SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
			string strConn = dataMgr.GetConnectionString(frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
				frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceSqliteFile);
			try
			{
				using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
				{
					conn.Open();
					dataMgr.SqlQueryReader(conn, "select * from states order by stabv");
					if (dataMgr.m_DataReader.HasRows == true)
					{
						while (dataMgr.m_DataReader.Read())
						{
							string strState = dataMgr.m_DataReader["stabv"].ToString().Trim() + " - " +
								dataMgr.m_DataReader["state_name"].ToString();
							this.cmbState.Items.Add(strState);
						}
					}
					dataMgr.m_DataReader.Close();
				}
			}
			catch 
			{
				// munch
			}        
		}
	
		public string strName
		{
			set	{ this.txtName.Text = value; }
			get { return this.txtName.Text.ToString().Trim(); }
		}
		public string strEmail
		{
			set	{ this.txtEmail.Text = value; }
			get { return this.txtEmail.Text.ToString(); }
		}
		public string strStreetAddress
		{
			set	{ this.txtStreet.Text = value; }
			get { return this.txtStreet.Text.ToString().Trim(); }
		}
		
		public string strPhoneNumber
		{
			set	
			{ 
				for (int x=0;x<=value.Trim().Length -1;x++)
				{
					if (x <= 2) this.txtAreaCode.Text += value.Substring(x,1);
					else if (x <= 5) this.txtPrefix.Text += value.Substring(x,1);
					else if (x <= 9) this.txtPhone.Text += value.Substring(x,1);
					else this.txtExt.Text += value.Substring(x,1);
				}
			}
			get 
			{
				return this.txtAreaCode.Text.Trim() + this.txtPrefix.Text.Trim() + this.txtPhone.Text.Trim() + this.txtExt.Text.Trim();
			}
		}
		public string strZipCode
		{
			set	
			{ 
				for (int x=0;x<=value.Trim().Length -1;x++)
				{
					if (x <= 4) this.txtZip1.Text += value.Substring(x,1);
					else this.txtZip2.Text += value.Substring(x,1);
				}
			}
			get 
			{
				return this.txtZip1.Text.Trim() + this.txtZip2.Text.Trim();
			}
		}
		public string strState
		{
			set	{ this.cmbState.Text = value; }
			get 
			{ 
				if (this.cmbState.Text.Trim().Length < 2)
				{
					return "";
				}
				else
				{
					return this.cmbState.Text.Trim().Substring(0,2).ToString();
				}
			}
		}
		public string strCity
		{
			set	{ this.txtCity.Text  = value; }
			get { return this.txtCity.Text.Trim(); }
		}
		public string strOrganization
		{
			set { this.txtOrg.Text = value; }
			get { return this.txtOrg.Text.Trim(); }
		}
		public bool checkFvs
		{
			set	{ this.chkFvs.Checked = value; }
			get { return this.chkFvs.Checked; }
		}
		public bool checkGis
		{
			set	{ this.chkGis.Checked = value; }
			get { return this.chkGis.Checked; }
		}
		public bool checkFrcs
		{
			set	{ this.chkOpcost.Checked = value; }
			get { return this.chkOpcost.Checked; }
		}
		public bool checkProcessor
		{
			set	{ this.chkProcessor.Checked = value; }
			get { return this.chkProcessor.Checked; }
		}
		public bool checkCore
		{
			set	{ this.chkOptimizer.Checked = value; }
			get { return this.chkOptimizer.Checked; }
		}
		



		

	}
}

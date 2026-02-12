using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;


namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_rx_harvest_method.
	/// </summary>
	public class uc_rx_harvest_method : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.GroupBox grpboxHarvestMethod;
		private System.Windows.Forms.GroupBox grpboxSteepSlopeHarvestMethod;
		private System.Windows.Forms.TextBox txtDesc;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.ComboBox cmbMethod;
		private System.Windows.Forms.Label lblMethod;
		private System.Windows.Forms.Label lblDesc;
		private System.Windows.Forms.Label lblSteepSlopeDesc;
		private System.Windows.Forms.TextBox txtSteepSlopeDesc;
		private System.Windows.Forms.ComboBox cmbSteepSlopeMethod;
		private System.Windows.Forms.Label lblSteepSlopeMethod;
		private System.Windows.Forms.TextBox txtRxDesc;
		private System.Windows.Forms.Label lblRxDesc;
		private FIA_Biosum_Manager.frmRxItem _frmRxItem=null;
		private Queries m_oQueries = new Queries();
		private RxTools m_oRxTools = new RxTools();
        private System.Collections.Generic.IDictionary<string, string> m_descrDictionary =
            new System.Collections.Generic.Dictionary<string, string>();
        private System.Collections.Generic.IDictionary<string, string> m_steepDescrDictionary =
            new System.Collections.Generic.Dictionary<string, string>();

        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

		public uc_rx_harvest_method()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			

			// TODO: Add any initialization after the InitializeComponent call

		}
		~uc_rx_harvest_method()
		{
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblRxDesc = new System.Windows.Forms.Label();
            this.txtRxDesc = new System.Windows.Forms.TextBox();
            this.lblDesc = new System.Windows.Forms.Label();
            this.grpboxSteepSlopeHarvestMethod = new System.Windows.Forms.GroupBox();
            this.lblSteepSlopeDesc = new System.Windows.Forms.Label();
            this.txtSteepSlopeDesc = new System.Windows.Forms.TextBox();
            this.cmbSteepSlopeMethod = new System.Windows.Forms.ComboBox();
            this.lblSteepSlopeMethod = new System.Windows.Forms.Label();
            this.grpboxHarvestMethod = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDesc = new System.Windows.Forms.TextBox();
            this.cmbMethod = new System.Windows.Forms.ComboBox();
            this.lblMethod = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.grpboxSteepSlopeHarvestMethod.SuspendLayout();
            this.grpboxHarvestMethod.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(704, 454);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.lblRxDesc);
            this.panel1.Controls.Add(this.txtRxDesc);
            this.panel1.Controls.Add(this.lblDesc);
            this.panel1.Controls.Add(this.grpboxSteepSlopeHarvestMethod);
            this.panel1.Controls.Add(this.grpboxHarvestMethod);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 16);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(698, 435);
            this.panel1.TabIndex = 0;
            // 
            // lblRxDesc
            // 
            this.lblRxDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRxDesc.Location = new System.Drawing.Point(16, 8);
            this.lblRxDesc.Name = "lblRxDesc";
            this.lblRxDesc.Size = new System.Drawing.Size(320, 16);
            this.lblRxDesc.TabIndex = 4;
            this.lblRxDesc.Text = "Treatment Description (Read Only)";
            // 
            // txtRxDesc
            // 
            this.txtRxDesc.Location = new System.Drawing.Point(16, 24);
            this.txtRxDesc.Multiline = true;
            this.txtRxDesc.Name = "txtRxDesc";
            this.txtRxDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtRxDesc.Size = new System.Drawing.Size(656, 72);
            this.txtRxDesc.TabIndex = 3;
            this.txtRxDesc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRxDesc_KeyDown);
            this.txtRxDesc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRxDesc_KeyPress);
            // 
            // lblDesc
            // 
            this.lblDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDesc.Location = new System.Drawing.Point(16, 104);
            this.lblDesc.Name = "lblDesc";
            this.lblDesc.Size = new System.Drawing.Size(499, 36);
            this.lblDesc.TabIndex = 2;
            this.lblDesc.Text = "Default harvest method for the treatment.\r\n(Assumed unless overridden by a differ" +
    "ent method specified in a Processor scenario.)";
            // 
            // grpboxSteepSlopeHarvestMethod
            // 
            this.grpboxSteepSlopeHarvestMethod.Controls.Add(this.lblSteepSlopeDesc);
            this.grpboxSteepSlopeHarvestMethod.Controls.Add(this.txtSteepSlopeDesc);
            this.grpboxSteepSlopeHarvestMethod.Controls.Add(this.cmbSteepSlopeMethod);
            this.grpboxSteepSlopeHarvestMethod.Controls.Add(this.lblSteepSlopeMethod);
            this.grpboxSteepSlopeHarvestMethod.Location = new System.Drawing.Point(344, 143);
            this.grpboxSteepSlopeHarvestMethod.Name = "grpboxSteepSlopeHarvestMethod";
            this.grpboxSteepSlopeHarvestMethod.Size = new System.Drawing.Size(325, 280);
            this.grpboxSteepSlopeHarvestMethod.TabIndex = 1;
            this.grpboxSteepSlopeHarvestMethod.TabStop = false;
            this.grpboxSteepSlopeHarvestMethod.Text = "Steep Slope Harvest Method";
            // 
            // lblSteepSlopeDesc
            // 
            this.lblSteepSlopeDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSteepSlopeDesc.Location = new System.Drawing.Point(18, 56);
            this.lblSteepSlopeDesc.Name = "lblSteepSlopeDesc";
            this.lblSteepSlopeDesc.Size = new System.Drawing.Size(182, 16);
            this.lblSteepSlopeDesc.TabIndex = 7;
            this.lblSteepSlopeDesc.Text = "Description (Read Only)";
            // 
            // txtSteepSlopeDesc
            // 
            this.txtSteepSlopeDesc.Location = new System.Drawing.Point(18, 72);
            this.txtSteepSlopeDesc.Multiline = true;
            this.txtSteepSlopeDesc.Name = "txtSteepSlopeDesc";
            this.txtSteepSlopeDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtSteepSlopeDesc.Size = new System.Drawing.Size(288, 184);
            this.txtSteepSlopeDesc.TabIndex = 6;
            this.txtSteepSlopeDesc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSteepSlopeDesc_KeyDown);
            // 
            // cmbSteepSlopeMethod
            // 
            this.cmbSteepSlopeMethod.Location = new System.Drawing.Point(18, 32);
            this.cmbSteepSlopeMethod.Name = "cmbSteepSlopeMethod";
            this.cmbSteepSlopeMethod.Size = new System.Drawing.Size(294, 21);
            this.cmbSteepSlopeMethod.TabIndex = 5;
            this.cmbSteepSlopeMethod.SelectedValueChanged += new System.EventHandler(this.cmbSteepSlopeMethod_SelectedValueChanged);
            this.cmbSteepSlopeMethod.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbSteepSlopeMethod_KeyPress);
            // 
            // lblSteepSlopeMethod
            // 
            this.lblSteepSlopeMethod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSteepSlopeMethod.Location = new System.Drawing.Point(18, 16);
            this.lblSteepSlopeMethod.Name = "lblSteepSlopeMethod";
            this.lblSteepSlopeMethod.Size = new System.Drawing.Size(72, 16);
            this.lblSteepSlopeMethod.TabIndex = 4;
            this.lblSteepSlopeMethod.Text = "Method";
            // 
            // grpboxHarvestMethod
            // 
            this.grpboxHarvestMethod.Controls.Add(this.label2);
            this.grpboxHarvestMethod.Controls.Add(this.txtDesc);
            this.grpboxHarvestMethod.Controls.Add(this.cmbMethod);
            this.grpboxHarvestMethod.Controls.Add(this.lblMethod);
            this.grpboxHarvestMethod.ForeColor = System.Drawing.Color.Black;
            this.grpboxHarvestMethod.Location = new System.Drawing.Point(8, 143);
            this.grpboxHarvestMethod.Name = "grpboxHarvestMethod";
            this.grpboxHarvestMethod.Size = new System.Drawing.Size(325, 280);
            this.grpboxHarvestMethod.TabIndex = 0;
            this.grpboxHarvestMethod.TabStop = false;
            this.grpboxHarvestMethod.Text = "Low Slope Harvest Method";
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(16, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(184, 16);
            this.label2.TabIndex = 3;
            this.label2.Text = "Description (Read Only)";
            // 
            // txtDesc
            // 
            this.txtDesc.Location = new System.Drawing.Point(16, 72);
            this.txtDesc.Multiline = true;
            this.txtDesc.Name = "txtDesc";
            this.txtDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtDesc.Size = new System.Drawing.Size(288, 184);
            this.txtDesc.TabIndex = 2;
            this.txtDesc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDesc_KeyDown);
            // 
            // cmbMethod
            // 
            this.cmbMethod.Location = new System.Drawing.Point(16, 32);
            this.cmbMethod.Name = "cmbMethod";
            this.cmbMethod.Size = new System.Drawing.Size(288, 21);
            this.cmbMethod.TabIndex = 1;
            this.cmbMethod.SelectedValueChanged += new System.EventHandler(this.cmbMethod_SelectedValueChanged);
            this.cmbMethod.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbMethod_KeyPress);
            // 
            // lblMethod
            // 
            this.lblMethod.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMethod.Location = new System.Drawing.Point(16, 16);
            this.lblMethod.Name = "lblMethod";
            this.lblMethod.Size = new System.Drawing.Size(66, 16);
            this.lblMethod.TabIndex = 0;
            this.lblMethod.Text = "Method";
            // 
            // uc_rx_harvest_method
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_rx_harvest_method";
            this.Size = new System.Drawing.Size(704, 454);
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
		public void loadvalues()
		{
			int x;
			this.txtDesc.Text="";
			this.txtRxDesc.Text="";
			this.txtSteepSlopeDesc.Text="";
			m_oQueries.m_oFvs.LoadDatasource=true;
			m_oQueries.m_oReference.LoadDatasource=true;
			m_oQueries.LoadDatasources(true);
            string strDbFile = m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.HarvestMethods);
            m_descrDictionary.Clear();
            m_steepDescrDictionary.Clear();
            this.m_oRxTools.LoadRxHarvestMethods(strDbFile, m_oQueries, cmbMethod, cmbSteepSlopeMethod, m_descrDictionary,
                m_steepDescrDictionary);
		    for (x=0;x<=this.cmbMethod.Items.Count-1;x++)
			{
                if (ReferenceFormRxItem.m_oRxItem.HarvestMethodLowSlope.Trim().ToUpper() ==
			        cmbMethod.Items[x].ToString().Trim().ToUpper())
				{
					cmbMethod.Text = cmbMethod.Items[x].ToString().Trim();
				}
			}
			for (x=0;x<=this.cmbSteepSlopeMethod.Items.Count-1;x++)
			{
				if (ReferenceFormRxItem.m_oRxItem.HarvestMethodSteepSlope.Trim().ToUpper() ==
					cmbSteepSlopeMethod.Items[x].ToString().Trim().ToUpper())
				{
					cmbSteepSlopeMethod.Text = cmbSteepSlopeMethod.Items[x].ToString().Trim();
				}
			}
            this.txtRxDesc.Text = this.ReferenceFormRxItem.m_oRxItem.Description;			
		}

		private void cmbMethod_SelectedValueChanged(object sender, System.EventArgs e)
		{
			getDesc();
			
		}
		private void getDesc()
		{
            if (m_descrDictionary.Keys.Contains(cmbMethod.Text.Trim()))
            {
                this.txtDesc.Text = m_descrDictionary[cmbMethod.Text.Trim()];
            }            
		}

		private void cmbSteepSlopeMethod_SelectedValueChanged(object sender, System.EventArgs e)
		{
			getDescSteepSlope();
		}
		private void getDescSteepSlope()
		{
            if (m_steepDescrDictionary.Keys.Contains(cmbSteepSlopeMethod.Text.Trim()))
            {
                this.txtSteepSlopeDesc.Text = m_steepDescrDictionary[cmbSteepSlopeMethod.Text.Trim()];
            }
		}

		private void txtRxDesc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			e.Handled=false;	
		}

		private void txtDesc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			e.Handled=true;	
		}

		private void txtSteepSlopeDesc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
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
	
		public string RxDescription
		{
			set {this.txtRxDesc.Text = value;}
		}
		public string MethodLowSlope
		{
			get {return this.cmbMethod.Text;}
		}
		public string MethodSteepSlope
		{
			get {return this.cmbSteepSlopeMethod.Text;}
		}


		public FIA_Biosum_Manager.frmRxItem ReferenceFormRxItem
		{
			get {return this._frmRxItem;}
			set {this._frmRxItem=value;}

		}
	}
}

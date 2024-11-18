using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_Rx.
	/// </summary>
	public class uc_rx_edit : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Panel panel1;
		public System.Windows.Forms.ListBox lstPackageMember;
		public System.Windows.Forms.TextBox txtDesc;
        public FIA_Biosum_Manager.ResizeFormUsingVisibleScrollBars m_oResizeForm = new ResizeFormUsingVisibleScrollBars();
		private System.Windows.Forms.Label lblPackageMember;
		private System.Windows.Forms.Label lblDesc;
		private FIA_Biosum_Manager.frmRxItem _frmRxItem=null;
		private ado_data_access m_oAdo = new ado_data_access();
		private Queries m_oQueries = new Queries();
        private ComboBox cmbRxId;
        private Label lblId;

        //private FIA_Biosum_Manager.RxItem_Collection _oRxItemCollection;




        //private ScrollBars _visibleScrollbars = ScrollBars.None;
        //public event EventHandler VisibleScrollbarsChanged;
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

		public uc_rx_edit()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			m_oResizeForm.ScrollBarParentControl=panel1;
			m_oResizeForm.MaximumWidth=770;
			m_oResizeForm.MaximumHeight=630;

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
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblPackageMember = new System.Windows.Forms.Label();
            this.lblDesc = new System.Windows.Forms.Label();
            this.lstPackageMember = new System.Windows.Forms.ListBox();
            this.txtDesc = new System.Windows.Forms.TextBox();
            this.cmbRxId = new System.Windows.Forms.ComboBox();
            this.lblId = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(755, 560);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.lblId);
            this.panel1.Controls.Add(this.cmbRxId);
            this.panel1.Controls.Add(this.lblPackageMember);
            this.panel1.Controls.Add(this.lblDesc);
            this.panel1.Controls.Add(this.lstPackageMember);
            this.panel1.Controls.Add(this.txtDesc);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 18);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(749, 539);
            this.panel1.TabIndex = 0;
            this.panel1.Resize += new System.EventHandler(this.panel1_Resize);
            // 
            // lblPackageMember
            // 
            this.lblPackageMember.Location = new System.Drawing.Point(493, 57);
            this.lblPackageMember.Name = "lblPackageMember";
            this.lblPackageMember.Size = new System.Drawing.Size(235, 16);
            this.lblPackageMember.TabIndex = 21;
            this.lblPackageMember.Text = "Packages that use this treatment";
            // 
            // lblDesc
            // 
            this.lblDesc.Location = new System.Drawing.Point(16, 57);
            this.lblDesc.Name = "lblDesc";
            this.lblDesc.Size = new System.Drawing.Size(96, 16);
            this.lblDesc.TabIndex = 20;
            this.lblDesc.Text = "Description";
            this.lblDesc.Click += new System.EventHandler(this.label2_Click);
            // 
            // lstPackageMember
            // 
            this.lstPackageMember.ItemHeight = 16;
            this.lstPackageMember.Location = new System.Drawing.Point(496, 77);
            this.lstPackageMember.Name = "lstPackageMember";
            this.lstPackageMember.Size = new System.Drawing.Size(240, 164);
            this.lstPackageMember.TabIndex = 19;
            // 
            // txtDesc
            // 
            this.txtDesc.Location = new System.Drawing.Point(16, 77);
            this.txtDesc.Multiline = true;
            this.txtDesc.Name = "txtDesc";
            this.txtDesc.Size = new System.Drawing.Size(448, 168);
            this.txtDesc.TabIndex = 18;
            // 
            // cmbRxId
            // 
            this.cmbRxId.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbRxId.Location = new System.Drawing.Point(124, 5);
            this.cmbRxId.Name = "cmbRxId";
            this.cmbRxId.Size = new System.Drawing.Size(96, 24);
            this.cmbRxId.TabIndex = 30;
            // 
            // lblId
            // 
            this.lblId.Location = new System.Drawing.Point(16, 8);
            this.lblId.Name = "lblId";
            this.lblId.Size = new System.Drawing.Size(111, 18);
            this.lblId.TabIndex = 31;
            this.lblId.Text = "Prescription ID :";
            // 
            // uc_rx_edit
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_rx_edit";
            this.Size = new System.Drawing.Size(755, 560);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.ParentForm.DialogResult = DialogResult.OK;
			this.ParentForm.Close();
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			this.ParentForm.DialogResult = DialogResult.Cancel;
			this.ParentForm.Close();
		}

		private void txtRxId_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void panel1_Resize(object sender, System.EventArgs e)
		{
			
		}
		public void loadvalues()
		{
			
			if (this.ReferenceFormRxItem.m_strAction=="edit")
			{
				this.lblDesc.Top = this.cmbRxId.Top + 50;
				this.lblPackageMember.Top = this.lblDesc.Top;
				this.txtDesc.Top = this.lblDesc.Top + this.lblDesc.Height + 5;
				this.txtDesc.Height = this.ClientSize.Height - this.txtDesc.Top - 5;
				this.lstPackageMember.Top = this.txtDesc.Top;
				this.lstPackageMember.Height = this.txtDesc.Height;
				this.txtDesc.Text = this.ReferenceFormRxItem.m_oRxItem.Description;
                this.cmbRxId.Enabled = false;
                this.cmbRxId.Text = this.ReferenceFormRxItem.m_oRxItem.RxId;

                if (this.ReferenceFormRxItem.m_oRxItem.RxPackageMemberList.Trim().Length > 0)
                {
                    string[] strArray = frmMain.g_oUtils.ConvertListToArray(ReferenceFormRxItem.m_oRxItem.RxPackageMemberList, ",");
                    for (int x = 0; x <= strArray.Length - 1; x++)
                    {
                        this.lstPackageMember.Items.Add(strArray[x]);
                    }
                }
			}
			else
			{
                //this.m_oQueries.m_oFvs.LoadDatasource=true;
                //this.m_oQueries.LoadDatasources(true);
                //this.m_oAdo.OpenConnection(m_oAdo.getMDBConnString(this.m_oQueries.m_strTempDbFile,"",""));
                this.LoadAvailableRxIdComboBox();
                this.cmbRxId.Enabled = true;
			}
		}
		public void savevalues()
		{
			if (string.IsNullOrEmpty(this.cmbRxId.Text))
			{
				this.ReferenceFormRxItem.m_strError="Select a treatment";
				this.ReferenceFormRxItem.m_intError=-1;
				return;
			}
			if (this.ReferenceFormRxItem.m_strAction=="new")
			{
                this.ReferenceFormRxItem.m_oRxItem.RxId = this.cmbRxId.Text;
                this.ReferenceFormRxItem.m_oRxItem.Description=this.txtDesc.Text;
			}
			else
			{
				this.ReferenceFormRxItem.m_oRxItem.Description=this.txtDesc.Text;
			}
		    
		}

		public FIA_Biosum_Manager.frmRxItem ReferenceFormRxItem
		{
			get {return this._frmRxItem;}
			set {this._frmRxItem=value;}

		}

        private void LoadAvailableRxIdComboBox()
        {
            this.cmbRxId.Items.Clear();

            int x = 0;
            int y = 0;
            int intMin = 1;
            int intMax = 999;

            string[] strUsedRxIdArray = frmMain.g_oUtils.ConvertListToArray(this.ReferenceFormRxItem.UsedRxList, ",");

            for (x = intMin; x <= intMax; x++)
            {
                if (this.ReferenceFormRxItem.UsedRxList.Trim().Length > 0)
                {
                    for (y = 0; y <= strUsedRxIdArray.Length - 1; y++)
                    {
                        if (Convert.ToInt32(strUsedRxIdArray[y]) == x) break;
                    }
                    if (y > strUsedRxIdArray.Length - 1)
                    {
                        this.cmbRxId.Items.Add(Convert.ToString(x).PadLeft(3, '0'));
                    }
                }
                else
                {
                    this.cmbRxId.Items.Add(Convert.ToString(x).PadLeft(3, '0'));
                }
            }
        }

    }
}

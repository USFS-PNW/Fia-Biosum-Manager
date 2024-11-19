using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_rx_fvscmd_list.
	/// </summary>
	public class uc_rx_fvscmd_list : System.Windows.Forms.UserControl
	{
		public int m_intError=0;
		public string m_strError="";
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.ListView lvRxFVSCmd;
		//public FIA_Biosum_Manager.ResizeFormUsingVisibleScrollBars m_oResizeForm = new ResizeFormUsingVisibleScrollBars();
		private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvAlternateColors = new ListViewAlternateBackgroundColors();
		private System.Windows.Forms.Button btnUp;
		private System.Windows.Forms.Button btnDown;
		private FIA_Biosum_Manager.frmRxItem _frmRxItem=null;
		//private FIA_Biosum_Manager.frmRxPackageItem _frmRxPackageItem=null;
		//private bool _bRxPackageEdit=false;
		private int m_intCurrSelect=-1;
		//rx edit
		const int COLUMN_NULL=0;
		const int COLUMN_RX=1;
		const int COLUMN_FVSCMD=2;
		const int COLUMN_P1=3;
		const int COLUMN_P2=4;
		const int COLUMN_P3=5;
		const int COLUMN_P4=6;
		const int COLUMN_P5=7;
		const int COLUMN_P6=8;
		const int COLUMN_P7=9;
		const int COLUMN_OTHER=10;
		const int COLUMN_FVSCMDID=11;
        private Label lblDesc;
		
		


		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_rx_fvscmd_list()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			//m_oResizeForm.ScrollBarParentControl=panel1;

			// TODO: Add any initialization after the InitializeComponent call

			m_oLvAlternateColors.ReferenceAlternateBackgroundColor=frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvAlternateColors.ReferenceListView=this.lvRxFVSCmd;
			this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvAlternateColors.CustomFullRowSelect=true;
			if (frmMain.g_oGridViewFont != null) this.lvRxFVSCmd.Font = frmMain.g_oGridViewFont;


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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_rx_fvscmd_list));
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnDown = new System.Windows.Forms.Button();
            this.btnUp = new System.Windows.Forms.Button();
            this.lvRxFVSCmd = new System.Windows.Forms.ListView();
            this.lblDesc = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(688, 424);
            this.panel1.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblDesc);
            this.groupBox1.Controls.Add(this.btnDown);
            this.groupBox1.Controls.Add(this.btnUp);
            this.groupBox1.Controls.Add(this.lvRxFVSCmd);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(688, 424);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // btnDown
            // 
            this.btnDown.Image = ((System.Drawing.Image)(resources.GetObject("btnDown.Image")));
            this.btnDown.Location = new System.Drawing.Point(20, 396);
            this.btnDown.Name = "btnDown";
            this.btnDown.Size = new System.Drawing.Size(48, 24);
            this.btnDown.TabIndex = 25;
            this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
            // 
            // btnUp
            // 
            this.btnUp.Image = ((System.Drawing.Image)(resources.GetObject("btnUp.Image")));
            this.btnUp.Location = new System.Drawing.Point(20, 372);
            this.btnUp.Name = "btnUp";
            this.btnUp.Size = new System.Drawing.Size(48, 24);
            this.btnUp.TabIndex = 24;
            this.btnUp.Click += new System.EventHandler(this.btnUp_Click);
            // 
            // lvRxFVSCmd
            // 
            this.lvRxFVSCmd.GridLines = true;
            this.lvRxFVSCmd.Location = new System.Drawing.Point(16, 37);
            this.lvRxFVSCmd.MultiSelect = false;
            this.lvRxFVSCmd.Name = "lvRxFVSCmd";
            this.lvRxFVSCmd.Size = new System.Drawing.Size(656, 331);
            this.lvRxFVSCmd.TabIndex = 17;
            this.lvRxFVSCmd.UseCompatibleStateImageBehavior = false;
            this.lvRxFVSCmd.View = System.Windows.Forms.View.Details;
            this.lvRxFVSCmd.SelectedIndexChanged += new System.EventHandler(this.lvRxFVSCmd_SelectedIndexChanged);
            this.lvRxFVSCmd.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lvRxFVSCmd_MouseUp);
            // 
            // lblDesc
            // 
            this.lblDesc.AutoSize = true;
            this.lblDesc.Location = new System.Drawing.Point(13, 16);
            this.lblDesc.Name = "lblDesc";
            this.lblDesc.Size = new System.Drawing.Size(607, 13);
            this.lblDesc.TabIndex = 27;
            this.lblDesc.Text = "This documentation screen provides a place to manually enter information on the F" +
    "VS commands associated with the treatment.\r\n";
            // 
            // uc_rx_fvscmd_list
            // 
            this.Controls.Add(this.panel1);
            this.Name = "uc_rx_fvscmd_list";
            this.Size = new System.Drawing.Size(688, 424);
            this.Resize += new System.EventHandler(this.uc_rx_fvscmd_list_Resize);
            this.panel1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void uc_rx_fvscmd_list_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.lvRxFVSCmd.Left = 5;
				this.lvRxFVSCmd.Width = this.Width - 10;

				
				btnDown.Top = this.ClientSize.Height - btnDown.Height - 5;
				btnUp.Top = btnDown.Top - btnDown.Height;
				


				this.lvRxFVSCmd.Height = this.btnUp.Top - this.lvRxFVSCmd.Top - 10;
				
		
				
			}
			catch
			{
			}
		}
		private void btnUp_Click(object sender, EventArgs e)
		{
			if (this.lvRxFVSCmd.SelectedItems.Count == 0) return;
			if (this.lvRxFVSCmd.SelectedItems[0].Index == 0) return;

			int intSaveCurrSelect = this.m_intCurrSelect;
			
			int x,y;
			string[] strValuesArray = new string[this.lvRxFVSCmd.Columns.Count];
			//load into an array the values from the selected row
			for (x = 0; x <= this.lvRxFVSCmd.Columns.Count - 1; x++)
			{
				strValuesArray[x] = this.lvRxFVSCmd.SelectedItems[0].SubItems[x].Text;
			}
			//swap values between the selected row and the row above
			y = this.lvRxFVSCmd.SelectedItems[0].Index;
			for (x = 0; x <= this.lvRxFVSCmd.Columns.Count - 1; x++)
			{
				this.lvRxFVSCmd.SelectedItems[0].SubItems[x].Text =
					this.lvRxFVSCmd.Items[y - 1].SubItems[x].Text;

				this.lvRxFVSCmd.Items[y - 1].SubItems[x].Text = 
					strValuesArray[x];
			}
			//adjust the selected row
			this.lvRxFVSCmd.Items[y - 1].Selected = true;
			this.lvRxFVSCmd.Focus();

		}

		private void btnDown_Click(object sender, EventArgs e)
		{
			if (this.lvRxFVSCmd.SelectedItems.Count == 0) return;
			if (this.lvRxFVSCmd.SelectedItems[0].Index == this.lvRxFVSCmd.Items.Count-1) return;

			int intSaveCurrSelect = this.m_intCurrSelect;
			int x, y;

			string[] strValuesArray = new string[this.lvRxFVSCmd.Columns.Count];
			//load into an array the values from the selected row
			for (x = 0; x <= this.lvRxFVSCmd.Columns.Count - 1; x++)
			{
				strValuesArray[x] = this.lvRxFVSCmd.SelectedItems[0].SubItems[x].Text;
			}
			//swap values between the selected row and the row below
			y = this.lvRxFVSCmd.SelectedItems[0].Index;
			for (x = 0; x <= this.lvRxFVSCmd.Columns.Count - 1; x++)
			{
				this.lvRxFVSCmd.SelectedItems[0].SubItems[x].Text =
					this.lvRxFVSCmd.Items[y + 1].SubItems[x].Text;

				this.lvRxFVSCmd.Items[y + 1].SubItems[x].Text =
					strValuesArray[x];
			}
			//adjust the selected row
			this.lvRxFVSCmd.Items[y + 1].Selected = true;
			this.lvRxFVSCmd.Focus();
		}
		public void loadvalues()
		{
			int intRowCount=0;
			int x,y;
			
			this.lvRxFVSCmd.Clear();
			this.m_oLvAlternateColors.InitializeRowCollection();      
			
			
			this.lvRxFVSCmd.Columns.Add(" ",2,HorizontalAlignment.Left);
			this.lvRxFVSCmd.Columns.Add("Rx", 60, HorizontalAlignment.Left);
			this.lvRxFVSCmd.Columns.Add("FVSCmd", 100, HorizontalAlignment.Left);
			this.lvRxFVSCmd.Columns.Add("P1", 50, HorizontalAlignment.Left);
			this.lvRxFVSCmd.Columns.Add("P2", 50, HorizontalAlignment.Left);
			this.lvRxFVSCmd.Columns.Add("P3", 50, HorizontalAlignment.Left);
			this.lvRxFVSCmd.Columns.Add("P4", 50, HorizontalAlignment.Left);
			this.lvRxFVSCmd.Columns.Add("P5", 50, HorizontalAlignment.Left);
			this.lvRxFVSCmd.Columns.Add("P6", 50, HorizontalAlignment.Left);
			this.lvRxFVSCmd.Columns.Add("P7", 50, HorizontalAlignment.Left);
			this.lvRxFVSCmd.Columns.Add("Other", 100, HorizontalAlignment.Left);
			this.lvRxFVSCmd.Columns.Add("FVSCmdId",50, HorizontalAlignment.Left);

				
			this.m_oLvAlternateColors.ListView();

			if (this.lvRxFVSCmd.Items.Count > 0) this.lvRxFVSCmd.Items[0].Selected=true;
			
		}

		private void lvRxFVSCmd_SelectedIndexChanged(object sender, System.EventArgs e)
		{

			ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_FVSCMD,frmRxItem.BUTTON_OPEN]=true;
			ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_FVSCMD,frmRxItem.BUTTON_NEW]=true;
			ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_FVSCMD,frmRxItem.BUTTON_CLEARALL]=true;
			ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_FVSCMD,frmRxItem.BUTTON_DELETE]=true;
			ReferenceFormRxItem.m_bToolBarButtonEnabled[frmRxItem.UC_FVSCMD,frmRxItem.BUTTON_EDIT]=true;
			ReferenceFormRxItem.SetToolBarButtonsEnabled(frmRxItem.UC_FVSCMD);





			
			if (this.lvRxFVSCmd.SelectedItems.Count > 0)
			{
				m_intCurrSelect = lvRxFVSCmd.SelectedItems[0].Index;
				this.m_oLvAlternateColors.DelegateListViewItem(lvRxFVSCmd.SelectedItems[0]);
			}
		}

		private void lvRxFVSCmd_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = this.lvRxFVSCmd.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lvRxFVSCmd.Items[this.lvRxFVSCmd.TopItem.Index + (int)dblRow-1].Selected=true;
				}
			}
			catch 
			{
			}
		}
		public void EditItem()
		{
			int x;
			if (this.lvRxFVSCmd.SelectedItems.Count == 0) return;
			string strFvsCmd="";
			FIA_Biosum_Manager.frmRxItemFvsCmdItem frmFvsCmdItem1 = new frmRxItemFvsCmdItem();
			FIA_Biosum_Manager.RxPackageItemFvsCommandItem oRxPackageItemFvsCmdItem=null;
			frmFvsCmdItem1.MaximizeBox = true;
			frmFvsCmdItem1.BackColor = System.Drawing.SystemColors.Control;
			frmFvsCmdItem1.Text = "FVS: FVS Command Item (Edit)";
			

			//frmFvsCmdItem1.Initialize_Rx_User_Control();

			

			
			
			//frmFvsCmdItem1.uc_rx_edit1.m_oResizeForm.ScrollBarParentControl=frmFvsCmdItem1.uc_rx_edit1.ParentForm;
			


			
			//frmFvsCmdItem1.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmFvsCmdItem1.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			//find the current rxid
			//for (x=0;x<=this.m_oRxItem_Collection.Count-1;x++)
			//{
			//	if (this.m_oRxItem_Collection.Item(x).RxId.Trim() == 
			//		this.lstRx.SelectedItems[0].SubItems[COLUMN_RX].Text.Trim())
			//	{
			//		frmFvsCmdItem1.ReferenceRxItem = this.m_oRxItem_Collection.Item(x);
			//		break;
			//	}
			//}
			frmFvsCmdItem1.ReferenceUserControlFvsCmdList=this;
			
			frmFvsCmdItem1.m_strAction="edit";

		}

		public void EditItem_old()
		{
		}
		public void AddItem()
		{
			int x;
			RxTools oRxTools = new RxTools();
			
			FIA_Biosum_Manager.frmRxItemFvsCmdItem frmFvsCmdItem1 = new frmRxItemFvsCmdItem();
			frmFvsCmdItem1.MaximizeBox = true;
			frmFvsCmdItem1.BackColor = System.Drawing.SystemColors.Control;
			frmFvsCmdItem1.Text = "FVS: FVS Command Item (New)";
			

			
			frmFvsCmdItem1.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;			
			frmFvsCmdItem1.ReferenceUserControlFvsCmdList=this;
			oRxTools=null;
		}
		
		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			EditItem();
		}

			
		public FIA_Biosum_Manager.frmRxItem ReferenceFormRxItem
		{
			get {return this._frmRxItem;}
			set {this._frmRxItem=value;}

		}
		

	}

}

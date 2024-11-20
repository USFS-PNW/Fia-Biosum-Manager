using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for frmRxItemFvsCmdItem.
	/// </summary>
	public class frmRxItemFvsCmdItem : System.Windows.Forms.Form
	{
		public int m_intError=0;
		public string m_strError="";
		public string m_strAction="new";
		private FIA_Biosum_Manager.uc_rx_fvscmd_list _uc_rx_fvscmd_list;
		private FIA_Biosum_Manager.uc_rx_package_fvscmd_list _uc_rx_package_fvscmd_list;
		
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ToolBarButton btnOk;
		private System.Windows.Forms.ToolBarButton btnCancel;
		private System.Windows.Forms.ToolBarButton btnSeparator;
		private System.Windows.Forms.ToolBarButton btnOpen;
		private System.Windows.Forms.ToolBar tlbFvsCmdItem;
		public FIA_Biosum_Manager.uc_rx_fvscmd_edit uc_rx_fvscmd_edit1;
		private System.ComponentModel.IContainer components;

		public frmRxItemFvsCmdItem()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			uc_rx_fvscmd_edit1.m_oResizeForm.ControlToResize=this;

			uc_rx_fvscmd_edit1.m_oResizeForm.ResizeControl();
			

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
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

		public FIA_Biosum_Manager.uc_rx_fvscmd_list ReferenceUserControlFvsCmdList
		{
			get {return this._uc_rx_fvscmd_list;}
			set {this._uc_rx_fvscmd_list=value;}
		}
		public FIA_Biosum_Manager.uc_rx_package_fvscmd_list ReferenceUserControlPackageFvsCmdList
		{
			get {return this._uc_rx_package_fvscmd_list;}
			set {this._uc_rx_package_fvscmd_list=value;}
		}
		

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmRxItemFvsCmdItem));
            this.tlbFvsCmdItem = new System.Windows.Forms.ToolBar();
            this.btnOk = new System.Windows.Forms.ToolBarButton();
            this.btnCancel = new System.Windows.Forms.ToolBarButton();
            this.btnSeparator = new System.Windows.Forms.ToolBarButton();
            this.btnOpen = new System.Windows.Forms.ToolBarButton();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.uc_rx_fvscmd_edit1 = new FIA_Biosum_Manager.uc_rx_fvscmd_edit();
            this.SuspendLayout();
            // 
            // tlbFvsCmdItem
            // 
            this.tlbFvsCmdItem.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
            this.btnOk,
            this.btnCancel,
            this.btnSeparator,
            this.btnOpen});
            this.tlbFvsCmdItem.ButtonSize = new System.Drawing.Size(60, 45);
            this.tlbFvsCmdItem.DropDownArrows = true;
            this.tlbFvsCmdItem.ImageList = this.imageList1;
            this.tlbFvsCmdItem.Location = new System.Drawing.Point(0, 0);
            this.tlbFvsCmdItem.Name = "tlbFvsCmdItem";
            this.tlbFvsCmdItem.ShowToolTips = true;
            this.tlbFvsCmdItem.Size = new System.Drawing.Size(720, 42);
            this.tlbFvsCmdItem.TabIndex = 0;
            this.tlbFvsCmdItem.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.tlbFvsCmdItem_ButtonClick);
            // 
            // btnOk
            // 
            this.btnOk.ImageIndex = 0;
            this.btnOk.Name = "btnOk";
            this.btnOk.Text = "OK";
            // 
            // btnCancel
            // 
            this.btnCancel.ImageIndex = 1;
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Text = "Cancel";
            // 
            // btnSeparator
            // 
            this.btnSeparator.Name = "btnSeparator";
            this.btnSeparator.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
            // 
            // btnOpen
            // 
            this.btnOpen.ImageIndex = 2;
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Text = "Open";
            this.btnOpen.ToolTipText = "Open FVS kcp/key file";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "");
            this.imageList1.Images.SetKeyName(1, "");
            this.imageList1.Images.SetKeyName(2, "");
            // 
            // uc_rx_fvscmd_edit1
            // 
            this.uc_rx_fvscmd_edit1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.uc_rx_fvscmd_edit1.Location = new System.Drawing.Point(0, 42);
            this.uc_rx_fvscmd_edit1.Name = "uc_rx_fvscmd_edit1";
            this.uc_rx_fvscmd_edit1.ReferenceFormFvsCmdItem = null;
            this.uc_rx_fvscmd_edit1.RxPackageEdit = false;
            this.uc_rx_fvscmd_edit1.Size = new System.Drawing.Size(720, 372);
            this.uc_rx_fvscmd_edit1.TabIndex = 1;
            // 
            // frmRxItemFvsCmdItem
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(720, 414);
            this.Controls.Add(this.uc_rx_fvscmd_edit1);
            this.Controls.Add(this.tlbFvsCmdItem);
            this.Name = "frmRxItemFvsCmdItem";
            this.Text = "frmRxItemFvsCmdItem";
            this.Closing += new System.ComponentModel.CancelEventHandler(this.frmRxItemFvsCmdItem_Closing);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		private void frmRxItemFvsCmdItem_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void tlbFvsCmdItem_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			switch (e.Button.Text)
			{
				case "Cancel":
					this.DialogResult = DialogResult.Cancel;
					this.Close();
					break;
				case "Open":
					OpenKCPFile();
					break;
				case "OK":
					//this.uc_rx_fvscmd_edit1.savevalues();
					this.DialogResult=DialogResult.OK;
					this.Close();
					break;

			} 
		}
		private void OpenKCPFile()
		{
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "Open FVS KCP/KEY File";
			
			OpenFileDialog1.Filter = "FVS KCP/KEY File (*.KCP,*.KEY) |*.kcp;*.key";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK)
			{
				string strArg = OpenFileDialog1.FileName.Trim();
				System.Diagnostics.Process proc = new System.Diagnostics.Process();
				proc.StartInfo.FileName = "wordpad.exe";
				
				proc.StartInfo.Arguments = (char)34 + strArg + (char)34;
				proc.Start();
			}
		}
	}
}

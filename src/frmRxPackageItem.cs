using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for frmRx.
	/// </summary>
	public class frmRxPackageItem : System.Windows.Forms.Form
	{
		
		public int m_intError=0;
		public string m_strError="";
		public string m_strAction="new";
		
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.TabControl tabControl1;
		private FIA_Biosum_Manager.RxPackageItem  _oRxPackageItem=null;
		private FIA_Biosum_Manager.RxItem_Collection _oRxItemCollection=null;
		private FIA_Biosum_Manager.uc_rx_package_list _uc_rx_package_list;
		public FIA_Biosum_Manager.RxPackageItem m_oRxPackageItem=null;

		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ToolBarButton btnOk;
		private System.Windows.Forms.ToolBarButton btnCancel;
		private System.Windows.Forms.ToolBarButton btnSeparator1;
		private System.Windows.Forms.ToolBarButton btnFvsCmdEdit;
		private System.Windows.Forms.ToolBarButton btnFvsCmdDelete;
		private System.Windows.Forms.ToolBarButton btnFvsCmdClearAll;
		private System.Windows.Forms.ToolBarButton btnFvsCmdAdd;
		private System.Windows.Forms.ToolBarButton btnProperties;
		private System.Windows.Forms.ToolBarButton btnFvsCmdOpenKCPFile;
		public FIA_Biosum_Manager.uc_rx_package_edit uc_rx_package_edit1;
		private System.Windows.Forms.TabPage tbPgRxPackageItem;
		private System.Windows.Forms.ToolBar tlbRxPackageItem;
		private System.ComponentModel.IContainer components;
        private TabPage tbPgRxItemHarvestCostColumn;
        public uc_rx_package_harvest_cost_column_list uc_rx_package_harvest_cost_column_list1;
		private string _strUsedRxPackageList="";

        public bool[,] m_bToolBarButtonEnabled = new bool[3, 5] {{false,false,false,false,true},
																 {false,false,false,false,true},
																 {true,false,false,true,true}};
		                                                       
        public const byte UC_RXPACKAGE = 0;
        public const byte UC_HARVESTCOST = 1;
        public const byte UC_FVSCMD = 2;
        public const byte BUTTON_NEW = 0;
        public const byte BUTTON_EDIT = 1;
        public const byte BUTTON_DELETE = 2;
        public const byte BUTTON_CLEARALL = 3;
        public const byte BUTTON_OPEN = 4;

        private Control _oParentControl = null;
        private bool _bMinimizeMainForm = false;


		public frmRxPackageItem()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			this.btnFvsCmdAdd.Enabled=false;
			this.btnFvsCmdClearAll.Enabled=false;
			this.btnFvsCmdDelete.Enabled=false;
			
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
		

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmRxPackageItem));
            this.panel1 = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tbPgRxPackageItem = new System.Windows.Forms.TabPage();
            this.uc_rx_package_edit1 = new FIA_Biosum_Manager.uc_rx_package_edit();
            this.tbPgRxItemHarvestCostColumn = new System.Windows.Forms.TabPage();
            this.uc_rx_package_harvest_cost_column_list1 = new FIA_Biosum_Manager.uc_rx_package_harvest_cost_column_list();
            this.tlbRxPackageItem = new System.Windows.Forms.ToolBar();
            this.btnOk = new System.Windows.Forms.ToolBarButton();
            this.btnCancel = new System.Windows.Forms.ToolBarButton();
            this.btnProperties = new System.Windows.Forms.ToolBarButton();
            this.btnSeparator1 = new System.Windows.Forms.ToolBarButton();
            this.btnFvsCmdAdd = new System.Windows.Forms.ToolBarButton();
            this.btnFvsCmdEdit = new System.Windows.Forms.ToolBarButton();
            this.btnFvsCmdDelete = new System.Windows.Forms.ToolBarButton();
            this.btnFvsCmdClearAll = new System.Windows.Forms.ToolBarButton();
            this.btnFvsCmdOpenKCPFile = new System.Windows.Forms.ToolBarButton();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.panel1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tbPgRxPackageItem.SuspendLayout();
            this.tbPgRxItemHarvestCostColumn.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.tabControl1);
            this.panel1.Controls.Add(this.tlbRxPackageItem);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(808, 462);
            this.panel1.TabIndex = 0;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tbPgRxPackageItem);
            this.tabControl1.Controls.Add(this.tbPgRxItemHarvestCostColumn);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 44);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(808, 418);
            this.tabControl1.TabIndex = 3;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tbPgRxPackageItem
            // 
            this.tbPgRxPackageItem.Controls.Add(this.uc_rx_package_edit1);
            this.tbPgRxPackageItem.Location = new System.Drawing.Point(4, 22);
            this.tbPgRxPackageItem.Name = "tbPgRxPackageItem";
            this.tbPgRxPackageItem.Size = new System.Drawing.Size(800, 392);
            this.tbPgRxPackageItem.TabIndex = 0;
            this.tbPgRxPackageItem.Text = "Package";
            this.tbPgRxPackageItem.UseVisualStyleBackColor = true;
            // 
            // uc_rx_package_edit1
            // 
            this.uc_rx_package_edit1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.uc_rx_package_edit1.Location = new System.Drawing.Point(0, 0);
            this.uc_rx_package_edit1.Name = "uc_rx_package_edit1";
            this.uc_rx_package_edit1.ReferenceFormRxPackageItem = null;
            this.uc_rx_package_edit1.RxPackageId = "";
            this.uc_rx_package_edit1.Size = new System.Drawing.Size(800, 392);
            this.uc_rx_package_edit1.TabIndex = 0;
            // 
            // tbPgRxItemHarvestCostColumn
            // 
            this.tbPgRxItemHarvestCostColumn.Controls.Add(this.uc_rx_package_harvest_cost_column_list1);
            this.tbPgRxItemHarvestCostColumn.Location = new System.Drawing.Point(4, 22);
            this.tbPgRxItemHarvestCostColumn.Name = "tbPgRxItemHarvestCostColumn";
            this.tbPgRxItemHarvestCostColumn.Size = new System.Drawing.Size(800, 385);
            this.tbPgRxItemHarvestCostColumn.TabIndex = 2;
            this.tbPgRxItemHarvestCostColumn.Text = "Additional CPA";
            this.tbPgRxItemHarvestCostColumn.UseVisualStyleBackColor = true;
            // 
            // uc_rx_package_harvest_cost_column_list1
            // 
            this.uc_rx_package_harvest_cost_column_list1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.uc_rx_package_harvest_cost_column_list1.Location = new System.Drawing.Point(0, 0);
            this.uc_rx_package_harvest_cost_column_list1.Name = "uc_rx_package_harvest_cost_column_list1";
            this.uc_rx_package_harvest_cost_column_list1.ReferenceFormRxPackageItem = null;
            this.uc_rx_package_harvest_cost_column_list1.Size = new System.Drawing.Size(800, 385);

            // 
            // tlbRxPackageItem
            // 
            this.tlbRxPackageItem.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
            this.btnOk,
            this.btnCancel,
            this.btnProperties,
            this.btnSeparator1,
            this.btnFvsCmdAdd,
            this.btnFvsCmdEdit,
            this.btnFvsCmdDelete,
            this.btnFvsCmdClearAll,
            this.btnFvsCmdOpenKCPFile});
            this.tlbRxPackageItem.ButtonSize = new System.Drawing.Size(60, 45);
            this.tlbRxPackageItem.DropDownArrows = true;
            this.tlbRxPackageItem.ImageList = this.imageList1;
            this.tlbRxPackageItem.Location = new System.Drawing.Point(0, 0);
            this.tlbRxPackageItem.Name = "tlbRxPackageItem";
            this.tlbRxPackageItem.ShowToolTips = true;
            this.tlbRxPackageItem.Size = new System.Drawing.Size(808, 44);
            this.tlbRxPackageItem.TabIndex = 2;
            this.tlbRxPackageItem.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.tlbRxItem_ButtonClick);
            this.tlbRxPackageItem.Click += new System.EventHandler(this.tlbRxItem_Click);
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
            // btnProperties
            // 
            this.btnProperties.ImageIndex = 2;
            this.btnProperties.Name = "btnProperties";
            this.btnProperties.Text = "Properties";
            // 
            // btnSeparator1
            // 
            this.btnSeparator1.Name = "btnSeparator1";
            this.btnSeparator1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
            // 
            // btnFvsCmdAdd
            // 
            this.btnFvsCmdAdd.Enabled = false;
            this.btnFvsCmdAdd.ImageIndex = 3;
            this.btnFvsCmdAdd.Name = "btnFvsCmdAdd";
            this.btnFvsCmdAdd.Text = "New";
            // 
            // btnFvsCmdEdit
            // 
            this.btnFvsCmdEdit.Enabled = false;
            this.btnFvsCmdEdit.ImageIndex = 4;
            this.btnFvsCmdEdit.Name = "btnFvsCmdEdit";
            this.btnFvsCmdEdit.Text = "Edit";
            // 
            // btnFvsCmdDelete
            // 
            this.btnFvsCmdDelete.Enabled = false;
            this.btnFvsCmdDelete.ImageIndex = 5;
            this.btnFvsCmdDelete.Name = "btnFvsCmdDelete";
            this.btnFvsCmdDelete.Text = "Delete";
            // 
            // btnFvsCmdClearAll
            // 
            this.btnFvsCmdClearAll.Enabled = false;
            this.btnFvsCmdClearAll.ImageIndex = 6;
            this.btnFvsCmdClearAll.Name = "btnFvsCmdClearAll";
            this.btnFvsCmdClearAll.Text = "Clear All";
            // 
            // btnFvsCmdOpenKCPFile
            // 
            this.btnFvsCmdOpenKCPFile.Enabled = false;
            this.btnFvsCmdOpenKCPFile.ImageIndex = 7;
            this.btnFvsCmdOpenKCPFile.Name = "btnFvsCmdOpenKCPFile";
            this.btnFvsCmdOpenKCPFile.Text = "Open";
            this.btnFvsCmdOpenKCPFile.ToolTipText = "Open FVS Kcp/Key file";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "");
            this.imageList1.Images.SetKeyName(1, "");
            this.imageList1.Images.SetKeyName(2, "");
            this.imageList1.Images.SetKeyName(3, "");
            this.imageList1.Images.SetKeyName(4, "");
            this.imageList1.Images.SetKeyName(5, "");
            this.imageList1.Images.SetKeyName(6, "");
            this.imageList1.Images.SetKeyName(7, "");
            // 
            // frmRxPackageItem
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(808, 462);
            this.Controls.Add(this.panel1);
            this.Name = "frmRxPackageItem";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.Text = "Rx Package Item";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmRxPackageItem_FormClosing);
            this.Resize += new System.EventHandler(this.frmRxPackageItem_Resize);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tbPgRxPackageItem.ResumeLayout(false);
            this.tbPgRxItemHarvestCostColumn.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion
		public void loadvalues()
		{
			
			this.uc_rx_package_edit1.ReferenceFormRxPackageItem=this;
            this.uc_rx_package_harvest_cost_column_list1.ReferenceFormRxPackageItem = this;
            
			
			this.m_oRxPackageItem = new  RxPackageItem();

			this.m_oRxPackageItem.CopyProperties(ReferenceRxPackageItem,m_oRxPackageItem);

			

			if (m_strAction=="edit")
			{				
				this.uc_rx_package_edit1.loadvalues();
                this.uc_rx_package_harvest_cost_column_list1.loadvalues();
			}
			else
			{				
				this.uc_rx_package_edit1.loadvalues();
                this.uc_rx_package_harvest_cost_column_list1.loadvalues();
			}
		}
		private void savevalues()
		{
			this.uc_rx_package_edit1.savevalues();

			if (this.m_intError==0)
				this.ReferenceRxPackageItem.CopyProperties(this.m_oRxPackageItem,ReferenceRxPackageItem);
		}
					
		private void tlbRxItem_Click(object sender, System.EventArgs e)
		{
			 
		}
		public void SetToolBarFVSCommandButtonsEnabled(string p_strOpenYN,string p_strDeleteYN,string p_strClearAllYN, string p_strEditYN,string p_strNewYN)
		{
			if (p_strOpenYN.Trim().Length > 0)
			{
				if (p_strOpenYN == "Y")
				{
					if (!this.btnFvsCmdOpenKCPFile.Enabled) 
					{
						this.btnFvsCmdOpenKCPFile.Enabled=true;
					}
				}
				else 
				{
					if (this.btnFvsCmdOpenKCPFile.Enabled) 
						this.btnFvsCmdOpenKCPFile.Enabled= false;
				}
			}
			if (p_strDeleteYN.Trim().Length > 0)
			{
				if (p_strDeleteYN =="Y")
				{
					if (!this.btnFvsCmdDelete.Enabled) this.btnFvsCmdDelete.Enabled=true;
				}
				else
				{
					if (this.btnFvsCmdDelete.Enabled) this.btnFvsCmdDelete.Enabled=false;
				}
			}
			if (p_strClearAllYN.Trim().Length > 0)
			{
				if (p_strClearAllYN=="Y")
				{
					if (!this.btnFvsCmdClearAll.Enabled) btnFvsCmdClearAll.Enabled=true;
				}
				else
				{
					if (this.btnFvsCmdClearAll.Enabled) btnFvsCmdClearAll.Enabled=false;
				}
			}
			if (p_strEditYN.Trim().Length > 0)
			{
				if (p_strEditYN=="Y") 
				{
					if (!this.btnFvsCmdEdit.Enabled) this.btnFvsCmdEdit.Enabled=true;
				}
				else 
				{
					if (this.btnFvsCmdEdit.Enabled) this.btnFvsCmdEdit.Enabled=false;
				}
			}
			if (p_strNewYN.Trim().Length > 0)
			{
				if (p_strNewYN=="Y") 
				{	
					if (!this.btnFvsCmdAdd.Enabled) this.btnFvsCmdAdd.Enabled=true;
				}
				else 
				{ 
					if (this.btnFvsCmdAdd.Enabled) this.btnFvsCmdAdd.Enabled=false;
				}
			}
		}
        public void SetToolBarButtonsEnabled(byte p_byteUC)
        {
            this.btnFvsCmdAdd.Enabled = m_bToolBarButtonEnabled[p_byteUC, BUTTON_NEW];
            this.btnFvsCmdClearAll.Enabled = m_bToolBarButtonEnabled[p_byteUC, BUTTON_CLEARALL]; ;
            this.btnFvsCmdDelete.Enabled = m_bToolBarButtonEnabled[p_byteUC, BUTTON_DELETE]; ;
            this.btnFvsCmdEdit.Enabled = m_bToolBarButtonEnabled[p_byteUC, BUTTON_EDIT]; ;
            this.btnFvsCmdOpenKCPFile.Enabled = m_bToolBarButtonEnabled[p_byteUC, BUTTON_OPEN]; ;
        }
		private void tlbRxItem_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			switch (e.Button.Text)
			{
				case "Cancel":
					this.DialogResult = DialogResult.Cancel;
					this.Close();
					break;
				case "OK":
					savevalues();
					if (m_intError==0) 
					{
                        if (m_strAction == "new")
                            ReferenceUserControlRxPackageList.AddItem(ReferenceRxPackageItem);
                        else
                            ReferenceUserControlRxPackageList.UpdateItem(ReferenceRxPackageItem, this.uc_rx_package_edit1.m_bSave);     
						this.DialogResult = DialogResult.OK;
						this.Close();
					}
					break;
				case "Open":
					OpenKCPFile();
					break;
                case "Properties":
					Properties();
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

		private void tabControl1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
            switch (tabControl1.SelectedTab.Text.Trim().ToUpper())
            {
                case "PACKAGE":

                    this.btnFvsCmdAdd.Enabled = m_bToolBarButtonEnabled[UC_RXPACKAGE, BUTTON_NEW];
                    this.btnFvsCmdClearAll.Enabled = m_bToolBarButtonEnabled[UC_RXPACKAGE, BUTTON_CLEARALL]; ;
                    this.btnFvsCmdDelete.Enabled = m_bToolBarButtonEnabled[UC_RXPACKAGE, BUTTON_DELETE]; ;
                    this.btnFvsCmdEdit.Enabled = m_bToolBarButtonEnabled[UC_RXPACKAGE, BUTTON_EDIT]; ;
                    this.btnFvsCmdOpenKCPFile.Enabled = m_bToolBarButtonEnabled[UC_RXPACKAGE, BUTTON_OPEN]; ;

                    break;
                case "ASSOCIATED FVS COMMAND(S)":
                    this.btnFvsCmdAdd.Enabled = m_bToolBarButtonEnabled[UC_FVSCMD, BUTTON_NEW];
                    this.btnFvsCmdClearAll.Enabled = m_bToolBarButtonEnabled[UC_FVSCMD, BUTTON_CLEARALL]; ;
                    this.btnFvsCmdDelete.Enabled = m_bToolBarButtonEnabled[UC_FVSCMD, BUTTON_DELETE]; ;
                    this.btnFvsCmdEdit.Enabled = m_bToolBarButtonEnabled[UC_FVSCMD, BUTTON_EDIT]; ;
                    this.btnFvsCmdOpenKCPFile.Enabled = m_bToolBarButtonEnabled[UC_FVSCMD, BUTTON_OPEN]; ;

                    break;
                case "HARVEST COSTS":
                    this.btnFvsCmdAdd.Enabled = m_bToolBarButtonEnabled[UC_HARVESTCOST, BUTTON_NEW];
                    this.btnFvsCmdClearAll.Enabled = m_bToolBarButtonEnabled[UC_HARVESTCOST, BUTTON_CLEARALL]; ;
                    this.btnFvsCmdDelete.Enabled = m_bToolBarButtonEnabled[UC_HARVESTCOST, BUTTON_DELETE]; ;
                    this.btnFvsCmdEdit.Enabled = m_bToolBarButtonEnabled[UC_HARVESTCOST, BUTTON_EDIT]; ;
                    this.btnFvsCmdOpenKCPFile.Enabled = m_bToolBarButtonEnabled[UC_HARVESTCOST, BUTTON_OPEN]; ;

                    break;
               
            }
           
		}
		private void Properties()
		{


			

			FIA_Biosum_Manager.RxItem_Collection oRxColl = new RxItem_Collection();

			
            frmMain.g_sbpInfo.Text = "Creating Rx Package Properties Report...Stand By";
			for (int x=0;x<=ReferenceRxItemCollection.Count-1;x++)
			{
				if (ReferenceRxItemCollection.Item(x).Delete==false)
				{
					if (this.m_oRxPackageItem.SimulationYear1Rx.Trim()==
						ReferenceRxItemCollection.Item(x).RxId.Trim() ||
						this.m_oRxPackageItem.SimulationYear2Rx.Trim()==
						ReferenceRxItemCollection.Item(x).RxId.Trim() ||
						this.m_oRxPackageItem.SimulationYear3Rx.Trim()==
						ReferenceRxItemCollection.Item(x).RxId.Trim() ||
						this.m_oRxPackageItem.SimulationYear4Rx.Trim()==
						ReferenceRxItemCollection.Item(x).RxId.Trim())
					{

					
						RxItem oItem = new RxItem();
						oItem.CopyProperties(ReferenceRxItemCollection.Item(x),oItem);
						oRxColl.Add(oItem);
					}

				}
			}
			
			FIA_Biosum_Manager.RxPackageItem_Collection oRxPackageCollection = new RxPackageItem_Collection();
			oRxPackageCollection.Add(this.m_oRxPackageItem);

            FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
            frmTemp.Text = "FIA Biosum";
            frmTemp.AutoScroll = false;
            uc_textbox uc_textbox1 = new uc_textbox();
            frmTemp.Controls.Add(uc_textbox1);
            uc_textbox1.Dock = DockStyle.Fill;
            uc_textbox1.lblTitle.Text = "Package Properties";
            RxTools oRxTools = new RxTools();
            uc_textbox1.TextValue = oRxTools.PackageProperties(oRxPackageCollection, oRxColl);
            frmTemp.Show();

			frmMain.g_sbpInfo.Text = "Ready";


			

		}
	
		
        public uc_rx_package_list ReferenceUserControlRxPackageList
        {
            get { return _uc_rx_package_list; }
            set { _uc_rx_package_list = value; }
        }
		public FIA_Biosum_Manager.RxPackageItem ReferenceRxPackageItem
		{
			get {return _oRxPackageItem;}
			set {_oRxPackageItem=value;}
		}
		public FIA_Biosum_Manager.RxItem_Collection ReferenceRxItemCollection
		{
			get {return _oRxItemCollection;}
			set {_oRxItemCollection=value;}
		}
		public string UsedRxPackageList
		{
			set {this._strUsedRxPackageList=value;}
			get {return this._strUsedRxPackageList;}
		}
        public Control ParentControl
        {
            get { return _oParentControl; }
            set { _oParentControl = value; }
        }
        public bool MinimizeMainForm
        {
            set { _bMinimizeMainForm = value; }
            get { return _bMinimizeMainForm; }
        }

        private void frmRxPackageItem_Resize(object sender, EventArgs e)
        {
            if (this.MinimizeMainForm)
            {
                if (this.WindowState != System.Windows.Forms.FormWindowState.Minimized)
                {
                    frmMain.g_oFrmMain.WindowState = System.Windows.Forms.FormWindowState.Maximized;
                    ((frmDialog)ParentControl).WindowState = System.Windows.Forms.FormWindowState.Normal;
                }
                else
                {
                    frmMain.g_oFrmMain.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                    ((frmDialog)ParentControl).WindowState = System.Windows.Forms.FormWindowState.Minimized;

                }

            }
        }

        private void frmRxPackageItem_FormClosing(object sender, FormClosingEventArgs e)
        {
            ParentControl.Enabled = true;
        }

		
	}
}

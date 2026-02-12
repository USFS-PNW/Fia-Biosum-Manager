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
	public class frmRxItem : System.Windows.Forms.Form
	{
		
		public int m_intError=0;
		public string m_strError="";
		public string m_strAction="new";
		
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tbPgRxItem;
		private System.Windows.Forms.TabPage tbPgRxItemFvsCmd;
		public FIA_Biosum_Manager.uc_rx_edit uc_rx_edit1;
		
		private FIA_Biosum_Manager.RxItem  _oRxItem=null;
		private FIA_Biosum_Manager.uc_rx_list _uc_rx_list;
		public FIA_Biosum_Manager.RxItem m_oRxItem=null;
		private string _strUsedRxList="";
		private System.Windows.Forms.ToolBar tlbRxItem;
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
		private System.Windows.Forms.TabPage tbPgRxItemHarvestMethod;
		private FIA_Biosum_Manager.uc_rx_harvest_method uc_rx_harvest_method1;
		private System.Windows.Forms.TabPage tbPgRxItemAdditionalCpa;
		private FIA_Biosum_Manager.uc_rx_harvest_cost_column_list uc_rx_harvest_cost_column_list1;
		private System.ComponentModel.IContainer components;
		public  bool[,] m_bToolBarButtonEnabled = new bool[4,5] {{false,false,false,false,true},
																 {false,false,false,false,true},
																 {true,false,false,false,true},
		                                                         {true,false,false,false,true}};
		public const byte UC_RX=0;
		public const byte UC_HARVESTMETHOD=1;
		public const byte UC_HARVESTCOST=2;
		public const byte UC_FVSCMD=3;
		public const byte BUTTON_NEW=0;
		public const byte BUTTON_EDIT=1;
		public const byte BUTTON_DELETE=2;
		public const byte BUTTON_CLEARALL=3;
		public const byte BUTTON_OPEN=4;

        private Control _oParentControl = null;
        private bool _bMinimizeMainForm = false;

		public frmRxItem()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//

            // GitHub issue #70: Can't hide a tab page, but you can remove it
            // Remove the line below to reinstate
            tabControl1.TabPages.Remove(this.tbPgRxItemFvsCmd);

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmRxItem));
            this.panel1 = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tbPgRxItem = new System.Windows.Forms.TabPage();
            this.uc_rx_edit1 = new FIA_Biosum_Manager.uc_rx_edit();
            this.tbPgRxItemHarvestMethod = new System.Windows.Forms.TabPage();
            this.uc_rx_harvest_method1 = new FIA_Biosum_Manager.uc_rx_harvest_method();
            this.tbPgRxItemAdditionalCpa = new System.Windows.Forms.TabPage();
            this.uc_rx_harvest_cost_column_list1 = new FIA_Biosum_Manager.uc_rx_harvest_cost_column_list();
            this.tbPgRxItemFvsCmd = new System.Windows.Forms.TabPage();
            this.tlbRxItem = new System.Windows.Forms.ToolBar();
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
            this.tbPgRxItem.SuspendLayout();
            this.tbPgRxItemHarvestMethod.SuspendLayout();
            this.tbPgRxItemAdditionalCpa.SuspendLayout();
            this.tbPgRxItemFvsCmd.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.tabControl1);
            this.panel1.Controls.Add(this.tlbRxItem);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(808, 462);
            this.panel1.TabIndex = 0;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tbPgRxItem);
            this.tabControl1.Controls.Add(this.tbPgRxItemHarvestMethod);
            this.tabControl1.Controls.Add(this.tbPgRxItemAdditionalCpa);
            this.tabControl1.Controls.Add(this.tbPgRxItemFvsCmd);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 47);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(808, 415);
            this.tabControl1.TabIndex = 3;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tbPgRxItem
            // 
            this.tbPgRxItem.Controls.Add(this.uc_rx_edit1);
            this.tbPgRxItem.Location = new System.Drawing.Point(4, 25);
            this.tbPgRxItem.Name = "tbPgRxItem";
            this.tbPgRxItem.Size = new System.Drawing.Size(800, 386);
            this.tbPgRxItem.TabIndex = 0;
            this.tbPgRxItem.Text = "Treatment";
            // 
            // uc_rx_edit1
            // 
            this.uc_rx_edit1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.uc_rx_edit1.Location = new System.Drawing.Point(0, 0);
            this.uc_rx_edit1.Name = "uc_rx_edit1";
            this.uc_rx_edit1.ReferenceFormRxItem = null;
            this.uc_rx_edit1.Size = new System.Drawing.Size(800, 386);
            this.uc_rx_edit1.TabIndex = 0;
            // 
            // tbPgRxItemHarvestMethod
            // 
            this.tbPgRxItemHarvestMethod.Controls.Add(this.uc_rx_harvest_method1);
            this.tbPgRxItemHarvestMethod.Location = new System.Drawing.Point(4, 25);
            this.tbPgRxItemHarvestMethod.Name = "tbPgRxItemHarvestMethod";
            this.tbPgRxItemHarvestMethod.Size = new System.Drawing.Size(800, 386);
            this.tbPgRxItemHarvestMethod.TabIndex = 2;
            this.tbPgRxItemHarvestMethod.Text = "Harvest Method";
            // 
            // uc_rx_harvest_method1
            // 
            this.uc_rx_harvest_method1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.uc_rx_harvest_method1.Location = new System.Drawing.Point(0, 0);
            this.uc_rx_harvest_method1.Name = "uc_rx_harvest_method1";
            this.uc_rx_harvest_method1.ReferenceFormRxItem = null;
            this.uc_rx_harvest_method1.Size = new System.Drawing.Size(800, 386);
            this.uc_rx_harvest_method1.TabIndex = 0;
            // 
            // tbPgRxItemAdditionalCpa
            // 
            this.tbPgRxItemAdditionalCpa.Controls.Add(this.uc_rx_harvest_cost_column_list1);
            this.tbPgRxItemAdditionalCpa.Location = new System.Drawing.Point(4, 25);
            this.tbPgRxItemAdditionalCpa.Name = "tbPgRxItemAdditionalCpa";
            this.tbPgRxItemAdditionalCpa.Size = new System.Drawing.Size(800, 386);
            this.tbPgRxItemAdditionalCpa.TabIndex = 3;
            this.tbPgRxItemAdditionalCpa.Text = "Additional CPA";
            // 
            // uc_rx_harvest_cost_column_list1
            // 
            this.uc_rx_harvest_cost_column_list1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.uc_rx_harvest_cost_column_list1.Location = new System.Drawing.Point(0, 0);
            this.uc_rx_harvest_cost_column_list1.Name = "uc_rx_harvest_cost_column_list1";
            this.uc_rx_harvest_cost_column_list1.ReferenceFormRxItem = null;
            this.uc_rx_harvest_cost_column_list1.Size = new System.Drawing.Size(800, 386);
            this.uc_rx_harvest_cost_column_list1.TabIndex = 0;
            // 
            // tlbRxItem
            // 
            this.tlbRxItem.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
            this.btnOk,
            this.btnCancel,
            this.btnProperties,
            this.btnSeparator1,
            this.btnFvsCmdAdd,
            this.btnFvsCmdEdit,
            this.btnFvsCmdDelete,
            this.btnFvsCmdClearAll,
            this.btnFvsCmdOpenKCPFile});
            this.tlbRxItem.ButtonSize = new System.Drawing.Size(60, 45);
            this.tlbRxItem.DropDownArrows = true;
            this.tlbRxItem.ImageList = this.imageList1;
            this.tlbRxItem.Location = new System.Drawing.Point(0, 0);
            this.tlbRxItem.Name = "tlbRxItem";
            this.tlbRxItem.ShowToolTips = true;
            this.tlbRxItem.Size = new System.Drawing.Size(808, 47);
            this.tlbRxItem.TabIndex = 2;
            this.tlbRxItem.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.tlbRxItem_ButtonClick);
            this.tlbRxItem.Click += new System.EventHandler(this.tlbRxItem_Click);
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
            this.btnProperties.ImageIndex = 7;
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
            this.btnFvsCmdAdd.ImageIndex = 5;
            this.btnFvsCmdAdd.Name = "btnFvsCmdAdd";
            this.btnFvsCmdAdd.Text = "New";
            // 
            // btnFvsCmdEdit
            // 
            this.btnFvsCmdEdit.Enabled = false;
            this.btnFvsCmdEdit.ImageIndex = 6;
            this.btnFvsCmdEdit.Name = "btnFvsCmdEdit";
            this.btnFvsCmdEdit.Text = "Edit";
            // 
            // btnFvsCmdDelete
            // 
            this.btnFvsCmdDelete.Enabled = false;
            this.btnFvsCmdDelete.ImageIndex = 4;
            this.btnFvsCmdDelete.Name = "btnFvsCmdDelete";
            this.btnFvsCmdDelete.Text = "Delete";
            // 
            // btnFvsCmdClearAll
            // 
            this.btnFvsCmdClearAll.Enabled = false;
            this.btnFvsCmdClearAll.ImageIndex = 3;
            this.btnFvsCmdClearAll.Name = "btnFvsCmdClearAll";
            this.btnFvsCmdClearAll.Text = "Clear All";
            // 
            // btnFvsCmdOpenKCPFile
            // 
            this.btnFvsCmdOpenKCPFile.Enabled = false;
            this.btnFvsCmdOpenKCPFile.ImageIndex = 8;
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
            this.imageList1.Images.SetKeyName(8, "");
            // 
            // frmRxItem
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(808, 462);
            this.Controls.Add(this.panel1);
            this.Name = "frmRxItem";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show;
            this.Text = "Rx Item";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmRxItem_FormClosing);
            this.Resize += new System.EventHandler(this.frmRxItem_Resize);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tbPgRxItem.ResumeLayout(false);
            this.tbPgRxItemHarvestMethod.ResumeLayout(false);
            this.tbPgRxItemAdditionalCpa.ResumeLayout(false);
            this.tbPgRxItemFvsCmd.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion
		public void loadvalues()
		{
			
			this.uc_rx_edit1.ReferenceFormRxItem=this;
			this.uc_rx_harvest_method1.ReferenceFormRxItem=this;
			this.uc_rx_harvest_cost_column_list1.ReferenceFormRxItem=this;

			
			
			this.m_oRxItem = new RxItem();

			this.m_oRxItem.CopyProperties(ReferenceRxItem,m_oRxItem);

			
			if (m_strAction=="edit")
			{
				
				this.uc_rx_edit1.loadvalues();
				this.uc_rx_harvest_method1.loadvalues();
				this.uc_rx_harvest_cost_column_list1.loadvalues();

			}
			else
			{
				this.uc_rx_edit1.loadvalues();
				this.uc_rx_harvest_method1.loadvalues();
				this.uc_rx_harvest_cost_column_list1.loadvalues();
			}
		}
		private void savevalues()
		{
			this.uc_rx_edit1.savevalues();

			if (this.m_intError==0)
			{
                this.m_oRxItem.HarvestMethodLowSlope = this.uc_rx_harvest_method1.MethodLowSlope.Trim();
				this.m_oRxItem.HarvestMethodSteepSlope = this.uc_rx_harvest_method1.MethodSteepSlope.Trim();
			}
            if (this.m_intError == 0)
                this.uc_rx_harvest_cost_column_list1.savevalues();


			if (this.m_intError==0)
				this.ReferenceRxItem.CopyProperties(this.m_oRxItem,this.ReferenceRxItem);
			

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
                            this.ReferenceUserControlRxList.AddItem(this.ReferenceRxItem);
                        else
                            this.ReferenceUserControlRxList.Description(this.ReferenceRxItem.Description);                            
						this.DialogResult = DialogResult.OK;
						this.ReferenceUserControlRxList.btnSave.Enabled=true;;
						this.Close();
					}
					break;
				case "Open":
					OpenKCPFile();
					
					break;
                case "Properties":
					Properties();
					break;
				case "Edit":										
                    this.uc_rx_harvest_cost_column_list1.EditItem();
                    break;
                case "Delete":
                    this.uc_rx_harvest_cost_column_list1.RemoveItem();
					break;
				case "Clear All":
                    this.uc_rx_harvest_cost_column_list1.RemoveAllItems();
					break;
				case "New":
                    this.uc_rx_harvest_cost_column_list1.AddItem();
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
				case "RX":
					
					this.btnFvsCmdAdd.Enabled=m_bToolBarButtonEnabled[UC_RX,BUTTON_NEW];
					this.btnFvsCmdClearAll.Enabled=m_bToolBarButtonEnabled[UC_RX,BUTTON_CLEARALL];;
					this.btnFvsCmdDelete.Enabled=m_bToolBarButtonEnabled[UC_RX,BUTTON_DELETE];;
					this.btnFvsCmdEdit.Enabled=m_bToolBarButtonEnabled[UC_RX,BUTTON_EDIT];;
					this.btnFvsCmdOpenKCPFile.Enabled=m_bToolBarButtonEnabled[UC_FVSCMD,BUTTON_OPEN];;
					
					break;
				case "ASSOCIATED FVS COMMAND(S)":
					this.btnFvsCmdAdd.Enabled=m_bToolBarButtonEnabled[UC_FVSCMD,BUTTON_NEW];
					this.btnFvsCmdClearAll.Enabled=m_bToolBarButtonEnabled[UC_FVSCMD,BUTTON_CLEARALL];;
					this.btnFvsCmdDelete.Enabled=m_bToolBarButtonEnabled[UC_FVSCMD,BUTTON_DELETE];;
					this.btnFvsCmdEdit.Enabled=m_bToolBarButtonEnabled[UC_FVSCMD,BUTTON_EDIT];;
					this.btnFvsCmdOpenKCPFile.Enabled=m_bToolBarButtonEnabled[UC_FVSCMD,BUTTON_OPEN];;

					break;
                case "ADDITIONAL CPA":
					this.btnFvsCmdAdd.Enabled=m_bToolBarButtonEnabled[UC_HARVESTCOST,BUTTON_NEW];
					this.btnFvsCmdClearAll.Enabled=m_bToolBarButtonEnabled[UC_HARVESTCOST,BUTTON_CLEARALL];;
					this.btnFvsCmdDelete.Enabled=m_bToolBarButtonEnabled[UC_HARVESTCOST,BUTTON_DELETE];;
					this.btnFvsCmdEdit.Enabled=m_bToolBarButtonEnabled[UC_HARVESTCOST,BUTTON_EDIT];;
					this.btnFvsCmdOpenKCPFile.Enabled=m_bToolBarButtonEnabled[UC_HARVESTCOST,BUTTON_OPEN];;

					break;
				case "HARVEST METHOD":
					this.uc_rx_harvest_method1.RxDescription=this.uc_rx_edit1.txtDesc.Text;
					this.btnFvsCmdAdd.Enabled=m_bToolBarButtonEnabled[UC_HARVESTMETHOD,BUTTON_NEW];
					this.btnFvsCmdClearAll.Enabled=m_bToolBarButtonEnabled[UC_HARVESTMETHOD,BUTTON_CLEARALL];;
					this.btnFvsCmdDelete.Enabled=m_bToolBarButtonEnabled[UC_HARVESTMETHOD,BUTTON_DELETE];;
					this.btnFvsCmdEdit.Enabled=m_bToolBarButtonEnabled[UC_HARVESTMETHOD,BUTTON_EDIT];;
					this.btnFvsCmdOpenKCPFile.Enabled=m_bToolBarButtonEnabled[UC_HARVESTMETHOD,BUTTON_OPEN];;
					break;
			}
		}
		public void SetToolBarButtonsEnabled(byte p_byteUC)
		{
			this.btnFvsCmdAdd.Enabled=m_bToolBarButtonEnabled[p_byteUC,BUTTON_NEW];
			this.btnFvsCmdClearAll.Enabled=m_bToolBarButtonEnabled[p_byteUC,BUTTON_CLEARALL];;
			this.btnFvsCmdDelete.Enabled=m_bToolBarButtonEnabled[p_byteUC,BUTTON_DELETE];;
			this.btnFvsCmdEdit.Enabled=m_bToolBarButtonEnabled[p_byteUC,BUTTON_EDIT];;
			this.btnFvsCmdOpenKCPFile.Enabled=m_bToolBarButtonEnabled[p_byteUC,BUTTON_OPEN];;
		}

		private void Properties()
		{
			if (this.m_oRxItem.RxId.Trim().Length == 0) return;
			frmMain.g_sbpInfo.Text = "Creating Rx Properties Report...Stand By";
			FIA_Biosum_Manager.RxItem_Collection oRxCollection = new RxItem_Collection();
			FIA_Biosum_Manager.RxItem oRxItem = new RxItem();

			oRxItem.CopyProperties(m_oRxItem,oRxItem);
			oRxCollection.Add(oRxItem);
            FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
            frmTemp.Text = "FIA Biosum";
            frmTemp.AutoScroll = false;
            uc_textbox uc_textbox1 = new uc_textbox();
            frmTemp.Controls.Add(uc_textbox1);
            uc_textbox1.Dock = DockStyle.Fill;
            uc_textbox1.lblTitle.Text = "Treatment Properties";
            RxTools oRxTools = new RxTools();
            uc_textbox1.TextValue = oRxTools.TreatmentProperties(oRxCollection);
            frmTemp.Show();
			frmMain.g_sbpInfo.Text = "Ready";

		}
	
		public FIA_Biosum_Manager.uc_rx_list ReferenceUserControlRxList
		{
			get {return this._uc_rx_list;}
			set {this._uc_rx_list=value;}
		}
		public FIA_Biosum_Manager.RxItem ReferenceRxItem
		{
			get {return _oRxItem;}
			set {_oRxItem=value;}
		}
		public string UsedRxList
		{
			set {this._strUsedRxList=value;}
			get {return this._strUsedRxList;}
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

        private void frmRxItem_Resize(object sender, EventArgs e)
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

        private void frmRxItem_FormClosing(object sender, FormClosingEventArgs e)
        {
            ParentControl.Enabled = true;
        }
		
	}
}

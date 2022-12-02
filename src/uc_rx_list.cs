using System;
using System.Collections;
using System.ComponentModel;
using SQLite.ADO;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_rx_list.
	/// </summary>
	public class uc_rx_list : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnEdit;
		private System.Windows.Forms.Button btnNew;
		public int m_intDialogHt;
		public int m_intDialogWd;
		private System.Windows.Forms.ListView lstRx;
		private int m_intError;
		private System.Windows.Forms.Button btnDefault;
		public System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnClear;
		private System.Windows.Forms.Button btnHelp;
		private FIA_Biosum_Manager.ado_data_access m_ado;

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultFvsXPSFile;

		private Queries m_oQueries; 
		string m_strTable="";		
		private string m_strConn;
		private System.Windows.Forms.Button btnDelete;
		private ListViewAlternateBackgroundColors m_oLvAlternateColors = new FIA_Biosum_Manager.ListViewAlternateBackgroundColors();
		const int COLUMN_NULL=0;
		const int COLUMN_RX=1;
		const int COLUMN_DESC=2;
		private RxTools m_oRxTools = new RxTools();
		private FIA_Biosum_Manager.RxItem_Collection m_oRxItem_Collection = new RxItem_Collection();
        private FIA_Biosum_Manager.RxPackageItem_Collection m_oRxPackageItem_Collection = new RxPackageItem_Collection();
		private FIA_Biosum_Manager.RxItemFvsCommandItem_Collection m_oRxItemFvsCommandItem_Collection = new RxItemFvsCommandItem_Collection();
		private System.Windows.Forms.Button btnProperties;
		private frmMain _frmMain=null;
        private frmDialog _frmDialog = null;
        
	

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_rx_list()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

            this.m_oEnv = new env();

			this.m_intDialogHt = this.groupBox1.Top + this.btnClose.Top + this.btnClose.Height + 20;
			this.m_intDialogWd = this.groupBox1.Left + this.btnClose.Left + this.btnClose.Width + 20;

			m_oLvAlternateColors.ReferenceAlternateBackgroundColor=frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvAlternateColors.ReferenceListView=this.lstRx;
			this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvAlternateColors.CustomFullRowSelect=true;
			if (frmMain.g_oGridViewFont != null) lstRx.Font = frmMain.g_oGridViewFont;


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
		public frmMain ReferenceMainForm
		{
			set {_frmMain=value;}
			get {return _frmMain;}
		}
		public frmDialog ReferenceParentDialogForm
		{
			set {_frmDialog=value;}
			get {return _frmDialog;}
		}
		public void loadvalues()
		{
			int x;
			int y;
			int intIndex=0;
			string strRx="";
			string strDesc="";

			this.m_oQueries = new Queries();
			m_oQueries.m_oFvs.LoadDatasource=true;
			m_oQueries.LoadDatasources(true);

			
    

			this.m_oLvAlternateColors.InitializeRowCollection();      
			this.lstRx.Clear();
			this.lstRx.Columns.Add(" ",2,HorizontalAlignment.Left);
			this.lstRx.Columns.Add("Rx", 60, HorizontalAlignment.Left);
			this.lstRx.Columns.Add("Description", 300, HorizontalAlignment.Left);
			
					

			this.m_intError=0;

			this.m_oRxItem_Collection.Clear();
			this.m_oRxItemFvsCommandItem_Collection.Clear();

			this.m_oRxTools.LoadAllRxItemsFromTableIntoRxCollection(m_oQueries.m_strTempDbFile,this.m_oQueries,this.m_oRxItem_Collection);
			this.lstRx.BeginUpdate();
			for (x=0;x<=this.m_oRxItem_Collection.Count-1;x++)
			{
				if (m_oRxItem_Collection.Item(x).Delete==false)
				{
					lstRx.Items.Add("");
					lstRx.Items[intIndex].UseItemStyleForSubItems=false;
					lstRx.Items[intIndex].SubItems.Add(m_oRxItem_Collection.Item(x).RxId);
					lstRx.Items[intIndex].SubItems.Add(m_oRxItem_Collection.Item(x).Description);

					m_oLvAlternateColors.AddRow();
					m_oLvAlternateColors.AddColumns(intIndex,lstRx.Columns.Count);

					intIndex++;
				}
			}
			this.lstRx.EndUpdate();

			this.m_oLvAlternateColors.ListView();

            //Only enable 'Use Default Values' if no items in lstRx
            if (lstRx.Items.Count < 1)
                btnDefault.Enabled = true;

		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnProperties = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnHelp = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnDefault = new System.Windows.Forms.Button();
            this.btnNew = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.lstRx = new System.Windows.Forms.ListView();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnProperties);
            this.groupBox1.Controls.Add(this.btnDelete);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.btnClear);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnSave);
            this.groupBox1.Controls.Add(this.btnDefault);
            this.groupBox1.Controls.Add(this.btnNew);
            this.groupBox1.Controls.Add(this.btnEdit);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.lstRx);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(672, 480);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // btnProperties
            // 
            this.btnProperties.Location = new System.Drawing.Point(400, 392);
            this.btnProperties.Name = "btnProperties";
            this.btnProperties.Size = new System.Drawing.Size(114, 32);
            this.btnProperties.TabIndex = 11;
            this.btnProperties.Text = "Properties";
            this.btnProperties.Click += new System.EventHandler(this.btnProperties_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Enabled = false;
            this.btnDelete.Location = new System.Drawing.Point(272, 392);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(64, 32);
            this.btnDelete.TabIndex = 10;
            this.btnDelete.Text = "Delete";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(16, 432);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(96, 32);
            this.btnHelp.TabIndex = 8;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(336, 392);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(64, 32);
            this.btnClear.TabIndex = 7;
            this.btnClear.Text = "Clear All";
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(576, 392);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(64, 32);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Enabled = false;
            this.btnSave.Location = new System.Drawing.Point(512, 392);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(64, 32);
            this.btnSave.TabIndex = 5;
            this.btnSave.Text = "Save";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnDefault
            // 
            this.btnDefault.Enabled = false;
            this.btnDefault.Location = new System.Drawing.Point(32, 392);
            this.btnDefault.Name = "btnDefault";
            this.btnDefault.Size = new System.Drawing.Size(114, 32);
            this.btnDefault.TabIndex = 2;
            this.btnDefault.Text = "Use Default Values";
            this.btnDefault.Click += new System.EventHandler(this.btnDefault_Click);
            // 
            // btnNew
            // 
            this.btnNew.Location = new System.Drawing.Point(144, 392);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(64, 32);
            this.btnNew.TabIndex = 3;
            this.btnNew.Text = "New";
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Enabled = false;
            this.btnEdit.Location = new System.Drawing.Point(208, 392);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(64, 32);
            this.btnEdit.TabIndex = 4;
            this.btnEdit.Text = "Edit";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(560, 432);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 9;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lstRx
            // 
            this.lstRx.GridLines = true;
            this.lstRx.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstRx.HideSelection = false;
            this.lstRx.Location = new System.Drawing.Point(16, 48);
            this.lstRx.MultiSelect = false;
            this.lstRx.Name = "lstRx";
            this.lstRx.Size = new System.Drawing.Size(640, 336);
            this.lstRx.TabIndex = 1;
            this.lstRx.UseCompatibleStateImageBehavior = false;
            this.lstRx.View = System.Windows.Forms.View.Details;
            this.lstRx.SelectedIndexChanged += new System.EventHandler(this.lstRx_SelectedIndexChanged);
            this.lstRx.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstRx_MouseUp);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(666, 24);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "Treatment List";
            // 
            // uc_rx_list
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_rx_list";
            this.Size = new System.Drawing.Size(672, 480);
            this.Resize += new System.EventHandler(this.uc_rx_list_Resize);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			if (this.btnSave.Enabled==true)
			{
				DialogResult result = MessageBox.Show("Save Changes Y/N","Plot Treatments",System.Windows.Forms.MessageBoxButtons.YesNo,System.Windows.Forms.MessageBoxIcon.Question);
				if (result == System.Windows.Forms.DialogResult.Yes)
				{
					this.savevalues();
				}
			}
            ((frmDialog)ParentForm).ParentControl.Enabled = true;
			this.ParentForm.Close();
		}

		private void btnNew_Click(object sender, System.EventArgs e)
		{

			string strDesc="";
			string strRx = "";
			int x;

			

			FIA_Biosum_Manager.frmRxItem frmRxItem1 = new frmRxItem();
			frmRxItem1.MaximizeBox = true;
			frmRxItem1.BackColor = System.Drawing.SystemColors.Control;
			frmRxItem1.Text = "FVS: Treatment (New)";


			//frmRxItem1.Initialize_Rx_User_Control();

			

			frmRxItem1.uc_rx_edit1.m_oResizeForm.ControlToResize = frmRxItem1;
			
			//frmRxItem1.uc_rx_edit1.m_oResizeForm.ScrollBarParentControl=frmRxItem1.uc_rx_edit1.ParentForm;
			


			frmRxItem1.uc_rx_edit1.m_oResizeForm.ResizeControl();

			//frmRxItem1.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmRxItem1.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			
			

			
			
			frmRxItem1.ReferenceUserControlRxList=this;
			frmRxItem1.UsedRxList=this.m_oRxTools.GetUsedRxList(this.m_oRxItem_Collection);
			RxItem oRxItem = new RxItem();
		    frmRxItem1.ReferenceRxItem = oRxItem;
			
			frmRxItem1.m_strAction="new";
			frmRxItem1.loadvalues();
            frmRxItem1.ParentControl = (frmDialog)ParentForm;
            frmRxItem1.ParentControl.Enabled = false;
            frmRxItem1.MinimizeMainForm = true;
            frmRxItem1.Show();


		}
		public void AddItem(RxItem p_oRxItem)
		{
			
			this.lstRx.Items.Add("");
			lstRx.Items[lstRx.Items.Count-1].UseItemStyleForSubItems=false;
			this.lstRx.Items[lstRx.Items.Count-1].SubItems.Add(p_oRxItem.RxId);
			this.lstRx.Items[lstRx.Items.Count-1].SubItems.Add(p_oRxItem.Description);
			
			this.m_oLvAlternateColors.AddRow();
			this.m_oLvAlternateColors.AddColumns(lstRx.Items.Count-1,lstRx.Columns.Count);

			this.m_oRxItem_Collection.Add(p_oRxItem);

			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
		}

		private void btnDefault_Click(object sender, System.EventArgs e)
		{
			ado_data_access oAdo = new ado_data_access();
			oAdo.OpenConnection(oAdo.getMDBConnString(m_oQueries.m_strTempDbFile,"",""));

			for (int x=0;x<=this.m_oRxItem_Collection.Count-1;x++)
			{
				this.m_oRxItem_Collection.Item(x).Delete=true;
				this.m_oRxItem_Collection.Item(x).Index=-1;
			}
			

			this.lstRx.Items.Clear();
			this.m_oLvAlternateColors.InitializeRowCollection();


			this.m_intError=0;
			this.lstRx.BeginUpdate();
			//
			//1st Default Treatment
			//
			this.lstRx.Items.Add("");
			lstRx.Items[0].UseItemStyleForSubItems=false;
			this.lstRx.Items[0].SubItems.Add("050");
			this.lstRx.Items[0].SubItems.Add("Thin-from-below: Applies to all trees 1 to 21 inches " + 
				                             "DBH to a target residual BA of 90 sq. ft. Trees >21 " + 
				                             "inches DBH will not be harvested. Leave 20% of the material " + 
											 "harvested in the woods. Leave all hardwoods standing.");
			
			this.m_oLvAlternateColors.AddRow();
			this.m_oLvAlternateColors.AddColumns(0,lstRx.Columns.Count);

			FIA_Biosum_Manager.RxItem oItem = new RxItem();
			oItem.RxId = "050";
			oItem.Description = "Thin-from-below: Applies to all trees 1 to 21 inches " + 
				                             "DBH to a target residual BA of 90 sq. ft. Trees >21 " + 
				                             "inches DBH will not be harvested. Leave 20% of the material " + 
											 "harvested in the woods. Leave all hardwoods standing.";
			oItem.Index = 0;
			//oItem.ReferenceFvsCommandsCollection = this.m_oRxItemFvsCommandItem_Collection;
			oItem.Category = getFvsCategoryDescription(oAdo,oItem.RxId);
			oItem.SubCategory = getFvsSubCategoryDescription(oAdo,oItem.RxId);
			oItem.Add=true;
			this.m_oRxItem_Collection.Add(oItem);

			m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1 = new RxItemFvsCommandItem_Collection();
			m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).ReferenceFvsCommandsCollection = this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1;
			FIA_Biosum_Manager.RxItemFvsCommandItem oFvsCmdItem = new RxItemFvsCommandItem();
			oFvsCmdItem.Index = 0;
			oFvsCmdItem.RxId = "050";
			oFvsCmdItem.FVSCommand = "SpGroup";
			oFvsCmdItem.FVSCommandId=1;
			oFvsCmdItem.Parameter1 = "LeaveTrees";
			oFvsCmdItem.Other = "OH BM BU RA FL TO CY BL WO BO VO IO WI CL WJ BR CP CN MA GC DG WA AS CW CH OT";
			oFvsCmdItem.Add=true;

			this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1.Add(oFvsCmdItem);


			oFvsCmdItem = new RxItemFvsCommandItem();
			oFvsCmdItem.Index = 1;
			oFvsCmdItem.RxId = "050";
			oFvsCmdItem.FVSCommand = "SpecPref";
			oFvsCmdItem.FVSCommandId=1;
			oFvsCmdItem.Parameter1 = "NA";
			oFvsCmdItem.Parameter1Description = "Cycle: Defined in Rx Package";
			oFvsCmdItem.Parameter2 = "LeaveTrees";
			oFvsCmdItem.Parameter3 = "-200";
			oFvsCmdItem.Add=true;


			this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1.Add(oFvsCmdItem);

			oFvsCmdItem = new RxItemFvsCommandItem();
			oFvsCmdItem.Index = 2;
			oFvsCmdItem.RxId = "050";
		    oFvsCmdItem.FVSCommand = "ThinBBA";
			oFvsCmdItem.FVSCommandId=1;
			oFvsCmdItem.Parameter1 = "NA";
			oFvsCmdItem.Parameter1Description = "Cycle: Defined in Rx Package";
			oFvsCmdItem.Parameter2 = "90.";
			oFvsCmdItem.Parameter3 = "1.";
			oFvsCmdItem.Parameter4 = "1.";
			oFvsCmdItem.Parameter5 = "21.";
			oFvsCmdItem.Parameter6 = "0.";
			oFvsCmdItem.Parameter7 = "999.";
			oFvsCmdItem.Add=true;

			this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1.Add(oFvsCmdItem);

			oFvsCmdItem = new RxItemFvsCommandItem();
			oFvsCmdItem.Index = 3;
			oFvsCmdItem.RxId = "050";
			oFvsCmdItem.FVSCommand = "Yardloss";
			oFvsCmdItem.FVSCommandId=1;
			oFvsCmdItem.Parameter1 = "NA";
			oFvsCmdItem.Parameter1Description = "Cycle: Defined in Rx Package";
			oFvsCmdItem.Parameter2 = "0.";
			oFvsCmdItem.Parameter3 = ".2";
			oFvsCmdItem.Parameter4 = "1";
			oFvsCmdItem.Parameter5 = "1";
			oFvsCmdItem.Add=true;

			this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1.Add(oFvsCmdItem);
			//
			//2nd Default Treatment
			//
			this.lstRx.Items.Add("");
			lstRx.Items[1].UseItemStyleForSubItems=false;
			this.lstRx.Items[1].SubItems.Add("051");
			this.lstRx.Items[1].SubItems.Add("Thin-from-below: Applies to all trees 1 to 15 inches " + 
				"DBH to a target residual BA of 60 sq. ft. Trees >15 " + 
				"inches DBH will not be harvested. Leave 20% of the material " + 
				"harvested in the woods. Leave all hardwoods standing.");
			
			this.m_oLvAlternateColors.AddRow();
			this.m_oLvAlternateColors.AddColumns(1,this.lstRx.Columns.Count);

			oItem = new RxItem();
			oItem.RxId = "051";
			oItem.Description = "Thin-from-below: Applies to all trees 1 to 15 inches " + 
				"DBH to a target residual BA of 60 sq. ft. Trees >15 " + 
				"inches DBH will not be harvested. Leave 20% of the material " + 
				"harvested in the woods. Leave all hardwoods standing.";
			oItem.Index = 1;
			//oItem.ReferenceFvsCommandsCollection = this.m_oRxItemFvsCommandItem_Collection;
			oItem.Category = getFvsCategoryDescription(oAdo,oItem.RxId);
			oItem.SubCategory = getFvsSubCategoryDescription(oAdo,oItem.RxId);
			oItem.Add=true;
			this.m_oRxItem_Collection.Add(oItem);

			m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1 = new RxItemFvsCommandItem_Collection();
			m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).ReferenceFvsCommandsCollection = this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1;

			oFvsCmdItem = new RxItemFvsCommandItem();
			oFvsCmdItem.Index = 0;
			oFvsCmdItem.RxId = "051";
			oFvsCmdItem.FVSCommand = "SpGroup";
			oFvsCmdItem.FVSCommandId=1;
			oFvsCmdItem.Parameter1 = "LeaveTrees";
			oFvsCmdItem.Other = "OH BM BU RA FL TO CY BL WO BO VO IO WI CL WJ BR CP CN MA GC DG WA AS CW CH OT";
			oFvsCmdItem.Add=true;

			this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1.Add(oFvsCmdItem);

			oFvsCmdItem = new RxItemFvsCommandItem();
			oFvsCmdItem.Index = 1;
			oFvsCmdItem.RxId = "051";
			oFvsCmdItem.FVSCommand = "SpecPref";
			oFvsCmdItem.FVSCommandId=1;
			oFvsCmdItem.Parameter1 = "NA";
			oFvsCmdItem.Parameter1Description = "Cycle: Defined in Rx Package";
			oFvsCmdItem.Parameter2 = "LeaveTrees";
			oFvsCmdItem.Parameter3 = "-200";
			oFvsCmdItem.Add=true;

			this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1.Add(oFvsCmdItem);

			oFvsCmdItem = new RxItemFvsCommandItem();
			oFvsCmdItem.Index = 2;
			oFvsCmdItem.RxId = "051";
			oFvsCmdItem.FVSCommand = "ThinBBA";
			oFvsCmdItem.FVSCommandId=1;
			oFvsCmdItem.Parameter1 = "NA";
			oFvsCmdItem.Parameter1Description = "Cycle: Defined in Rx Package";
			oFvsCmdItem.Parameter2 = "60.";
			oFvsCmdItem.Parameter3 = "1.";
			oFvsCmdItem.Parameter4 = "1.";
			oFvsCmdItem.Parameter5 = "15.";
			oFvsCmdItem.Parameter6 = "0.";
			oFvsCmdItem.Parameter7 = "999.";
			oFvsCmdItem.Add=true;

			this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1.Add(oFvsCmdItem);

			oFvsCmdItem = new RxItemFvsCommandItem();
			oFvsCmdItem.Index = 3;
			oFvsCmdItem.RxId = "051";
			oFvsCmdItem.FVSCommand = "Yardloss";
			oFvsCmdItem.FVSCommandId=1;
			oFvsCmdItem.Parameter1 = "NA";
			oFvsCmdItem.Parameter1Description = "Cycle: Defined in Rx Package";
			oFvsCmdItem.Parameter2 = "0.";
			oFvsCmdItem.Parameter3 = ".2";
			oFvsCmdItem.Parameter4 = "1";
			oFvsCmdItem.Parameter5 = "1";
			oFvsCmdItem.Add=true;

			this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1.Add(oFvsCmdItem);
			
			//
			//3rd Default Treatment
			//
			this.lstRx.Items.Add("");
			lstRx.Items[2].UseItemStyleForSubItems=false;
			this.lstRx.Items[2].SubItems.Add("052");
			this.lstRx.Items[2].SubItems.Add("Thin-from-below: Applies to all trees with a  " + 
				"DBH to a target residual BA of 60 sq. ft. All trees will be harvested. " + 
				"Leave 20% of the material " + 
				"harvested in the woods. Leave all hardwoods standing.");
			
			this.m_oLvAlternateColors.AddRow();
			this.m_oLvAlternateColors.AddColumns(2,lstRx.Columns.Count);

			oItem = new RxItem();
			oItem.RxId = "052";
			oItem.Description = "Thin-from-below: Applies to all trees with a  " + 
				"DBH to a target residual BA of 60 sq. ft. All trees will be harvested. " + 
				"Leave 20% of the material " + 
				"harvested in the woods. Leave all hardwoods standing.";
			oItem.Index = 2;
			oItem.Category = getFvsCategoryDescription(oAdo,oItem.RxId);
			oItem.SubCategory = getFvsSubCategoryDescription(oAdo,oItem.RxId);
			oItem.Add=true;
			//oItem.ReferenceFvsCommandsCollection = this.m_oRxItemFvsCommandItem_Collection;
			this.m_oRxItem_Collection.Add(oItem);

			m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1 = new RxItemFvsCommandItem_Collection();
			m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).ReferenceFvsCommandsCollection = this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1;
			oFvsCmdItem = new RxItemFvsCommandItem();
			oFvsCmdItem.Index = 0;
			oFvsCmdItem.RxId = "052";
			oFvsCmdItem.FVSCommand = "SpGroup";
			oFvsCmdItem.FVSCommandId=1;
			oFvsCmdItem.Parameter1 = "LeaveTrees";
			oFvsCmdItem.Other = "OH BM BU RA FL TO CY BL WO BO VO IO WI CL WJ BR CP CN MA GC DG WA AS CW CH OT";
			oFvsCmdItem.Add=true;

			this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1.Add(oFvsCmdItem);

			oFvsCmdItem = new RxItemFvsCommandItem();
			oFvsCmdItem.Index = 1;
			oFvsCmdItem.RxId = "052";
			oFvsCmdItem.FVSCommand = "SpecPref";
			oFvsCmdItem.FVSCommandId=1;
			oFvsCmdItem.Parameter1 = "NA";
			oFvsCmdItem.Parameter1Description = "Cycle: Defined in Rx Package";
			oFvsCmdItem.Parameter2 = "LeaveTrees";
			oFvsCmdItem.Parameter3 = "-200";
			oFvsCmdItem.Add=true;

			this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1.Add(oFvsCmdItem);

			oFvsCmdItem = new RxItemFvsCommandItem();
			oFvsCmdItem.Index = 2;
			oFvsCmdItem.RxId = "052";
			oFvsCmdItem.FVSCommand = "ThinBBA";
			oFvsCmdItem.FVSCommandId=1;
			oFvsCmdItem.Parameter1 = "NA";
			oFvsCmdItem.Parameter1Description = "Cycle: Defined in Rx Package";
			oFvsCmdItem.Parameter2 = "60.";
			oFvsCmdItem.Parameter3 = "1.";
			oFvsCmdItem.Parameter4 = "1.";
			oFvsCmdItem.Parameter5 = "999.";
			oFvsCmdItem.Parameter6 = "0.";
			oFvsCmdItem.Parameter7 = "999.";
			oFvsCmdItem.Add=true;

			this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1.Add(oFvsCmdItem);

			oFvsCmdItem = new RxItemFvsCommandItem();
			oFvsCmdItem.Index = 3;
			oFvsCmdItem.RxId = "052";
			oFvsCmdItem.FVSCommand = "Yardloss";
			oFvsCmdItem.FVSCommandId=1;
			oFvsCmdItem.Parameter1 = "NA";
			oFvsCmdItem.Parameter1Description = "Cycle: Defined in Rx Package";
			oFvsCmdItem.Parameter2 = "0.";
			oFvsCmdItem.Parameter3 = ".2";
			oFvsCmdItem.Parameter4 = "1";
			oFvsCmdItem.Parameter5 = "1";
			oFvsCmdItem.Add=true;

			this.m_oRxItem_Collection.Item(m_oRxItem_Collection.Count-1).m_oFvsCommandItem_Collection1.Add(oFvsCmdItem);

			



			this.lstRx.Items[0].Selected = true;





			this.lstRx.EndUpdate();
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			this.m_oLvAlternateColors.ListView();

			oAdo.CloseConnection(oAdo.m_OleDbConnection);
			oAdo=null;


		}

		private void btnClear_Click(object sender, System.EventArgs e)
		{
			if (this.lstRx.Items.Count > 0) this.btnSave.Enabled=true;
			this.btnDelete.Enabled=false;
			this.btnEdit.Enabled=false;
			

			this.lstRx.Clear();
			this.lstRx.Columns.Add("",2,HorizontalAlignment.Left);
			this.lstRx.Columns.Add("Rx ID", 60, HorizontalAlignment.Left);
			this.lstRx.Columns.Add("Description", 150, HorizontalAlignment.Left);
			this.m_oLvAlternateColors.InitializeRowCollection();

			for (int x=0;x<=this.m_oRxItem_Collection.Count-1;x++)
			{
				this.m_oRxItem_Collection.Item(x).Delete=true;
				this.m_oRxItem_Collection.Item(x).Index=-1;
			}

			
		}

		private void btnEdit_Click(object sender, System.EventArgs e)
		{

			string strDesc="";
			string strRxList = "";
			int x;

			if (this.lstRx.SelectedItems.Count == 0)
				return;

			FIA_Biosum_Manager.frmRxItem frmRxItem1 = new frmRxItem();
			RxItem oRxItem;
			frmRxItem1.MaximizeBox = true;
			frmRxItem1.BackColor = System.Drawing.SystemColors.Control;
			frmRxItem1.Text = "FVS: Treatment (Edit)";


			

			

			frmRxItem1.uc_rx_edit1.m_oResizeForm.ControlToResize = frmRxItem1;
			
			
			


			frmRxItem1.uc_rx_edit1.m_oResizeForm.ResizeControl();

			
			frmRxItem1.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			//find the current rxid
			for (x=0;x<=this.m_oRxItem_Collection.Count-1;x++)
			{
				if (this.m_oRxItem_Collection.Item(x).RxId.Trim() == 
					this.lstRx.SelectedItems[0].SubItems[COLUMN_RX].Text.Trim())
				{
					
					frmRxItem1.ReferenceRxItem = this.m_oRxItem_Collection.Item(x);
					
				}
				strRxList = strRxList + this.m_oRxItem_Collection.Item(x).RxId + ",";
			}
			if (strRxList.Trim().Length > 0) strRxList = strRxList.Substring(0,strRxList.Length - 1);
			frmRxItem1.ReferenceUserControlRxList=this;
			frmRxItem1.UsedRxList=strRxList;
            frmRxItem1.ParentControl = (frmDialog)ParentForm;
            frmRxItem1.ParentControl.Enabled = false;
			frmRxItem1.m_strAction="edit";
			frmRxItem1.loadvalues();
            frmRxItem1.MinimizeMainForm = true;
            frmRxItem1.Show();

		}
        public void Description(string p_strDesc)
        {
            this.lstRx.SelectedItems[0].SubItems[2].Text = p_strDesc;
        }
		private void btnSave_Click(object sender, System.EventArgs e)
		{
		  this.savevalues();
		}
		private void val_data()
		{
            

		}
		/// <summary>
		/// save the rx list items
		/// </summary>
		public void savevalues()
		{
			int x;
			int y;
			
			int intIndex=0;
			string strFields;
			string strValues;
			string strCatId;
			string strSubCatId;
			

			ado_data_access oAdo = new ado_data_access();
			oAdo.OpenConnection(oAdo.getMDBConnString(m_oQueries.m_strTempDbFile,"",""));
			if (oAdo.m_intError==0)
			{
				//delete all records from rx table
				oAdo.m_strSQL = "DELETE FROM " + this.m_oQueries.m_oFvs.m_strRxTable;
				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
				if (oAdo.m_intError==0)
				{
					//delete all records from the rx fvs commands table
					oAdo.m_strSQL="DELETE FROM " + this.m_oQueries.m_oFvs.m_strRxFvsCmdTable;
					oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
					if (oAdo.m_intError==0)
					{
						//delete all records from the rx harvest cost columns table
						oAdo.m_strSQL="DELETE FROM " + this.m_oQueries.m_oFvs.m_strRxHarvestCostColumnsTable;
						oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
						if (oAdo.m_intError==0)
						{
							//insert all the records
							for (x=0;x<=this.m_oRxItem_Collection.Count-1;x++)
							{
								if (m_oRxItem_Collection.Item(x).Delete==false)
								{
									//insert the rx record
									strFields="";
									strValues="";

                                    strFields = "rx,description,catid,subcatid,harvestmethodLowSlope,harvestmethodsteepslope";
									strValues = "'" +  m_oRxItem_Collection.Item(x).RxId.Trim() + "',";
									strValues = strValues + "'" + oAdo.FixString(m_oRxItem_Collection.Item(x).Description.Trim(),"'","''") + "',";
									oAdo.m_strSQL =this.m_oQueries.m_oFvs.GetCategoryIdFromDescriptionSQL(m_oRxItem_Collection.Item(x).Category);
									strCatId = oAdo.getSingleStringValueFromSQLQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL,"temp");
									if (strCatId.Trim().Length == 0) strCatId = "null";
									strValues = strValues + strCatId + ",";
									oAdo.m_strSQL = this.m_oQueries.m_oFvs.GetSubCategoryIdFromDescriptionSQL(m_oRxItem_Collection.Item(x).SubCategory);
									strSubCatId = oAdo.getSingleStringValueFromSQLQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL,"temp");
									if (strSubCatId.Trim().Length ==0) strSubCatId = "null";
									strValues = strValues + strSubCatId + ",";
                                    strValues = strValues + "'" + oAdo.FixString(m_oRxItem_Collection.Item(x).HarvestMethodLowSlope.Trim(), "'", "''") + "',";
									strValues = strValues + "'" + oAdo.FixString(m_oRxItem_Collection.Item(x).HarvestMethodSteepSlope.Trim(),"'","''") + "'";
									oAdo.m_strSQL = Queries.GetInsertSQL(strFields,strValues,m_oQueries.m_oFvs.m_strRxTable);
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);

									intIndex=0;
									//insert all the fvs commands associated with the rx
									if (this.m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection != null)
									{
										for (y=0;y<=this.m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Count-1;y++)
										{
											if (m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).Delete==false)
											{
												if (this.m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).RxId==
													this.m_oRxItem_Collection.Item(x).RxId && 
													this.m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).Index==intIndex)
												{
									
													strFields="rx_fvscmd_index,rx,fvscmd,fvscmd_id,p1,p2,p3,p4,p5,p6,p7,other";
													strValues=Convert.ToString(intIndex) + ","; 
													strValues=strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).RxId.Trim() + "',";
													strValues=strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommand.Trim() + "',";
													strValues=strValues + Convert.ToString(m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommandId).Trim() + ",";
													strValues=strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).Parameter1.Trim() + "',";
													strValues=strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).Parameter2.Trim() + "',";
													strValues=strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).Parameter3.Trim() + "',";
													strValues=strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).Parameter4.Trim() + "',";
													strValues=strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).Parameter5.Trim() + "',";
													strValues=strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).Parameter6.Trim() + "',";
													strValues=strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).Parameter7.Trim() + "',";
													strValues=strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).Other.Trim() + "'";
													oAdo.m_strSQL = Queries.GetInsertSQL(strFields,strValues,m_oQueries.m_oFvs.m_strRxFvsCmdTable);
													oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
													intIndex++;
													y=-1;
									
												}
											}
											else
											{
												//delete rx fvs command from the package table
												oAdo.m_strSQL = "DELETE FROM " + this.m_oQueries.m_oFvs.m_strRxPackageFvsCmdOrderTable + " " + 
													"WHERE TRIM(RX)='" + m_oRxItem_Collection.Item(x).RxId.Trim() + "' AND " + 
													"TRIM(FVSCMD)='" + m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommand.Trim() + "' AND " + 
													"FVSCMD_ID = " + Convert.ToString(m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommandId);
											
												oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);

											}
										}
									}
									
									//insert all the harvest cost columns for the rx
									if (this.m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection != null)
									{
										for (y=0;y<=this.m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Count-1;y++)
										{
											if (m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).Delete==false)
											{
												if (this.m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).RxId==
													this.m_oRxItem_Collection.Item(x).RxId)
												{
													strValues="";
													strFields="rx,ColumnName,description";
													strValues=strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).RxId.Trim() + "',";
													strValues=strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).HarvestCostColumn.Trim() + "',";
													strValues=strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).Description.Trim() + "'";
													oAdo.m_strSQL = Queries.GetInsertSQL(strFields,strValues,m_oQueries.m_oFvs.m_strRxHarvestCostColumnsTable);
													oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
													
									
												}
											}
											else
											{
												//delete rx + column_name from the harvest cost column table
												oAdo.m_strSQL = "DELETE FROM " + this.m_oQueries.m_oFvs.m_strRxHarvestCostColumnsTable + " " + 
													"WHERE TRIM(RX)='" + m_oRxItem_Collection.Item(x).RxId.Trim() + "' AND " + 
													"TRIM(ColumnName)='" + m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).HarvestCostColumn.Trim() + "'";
												oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);

											}
										}
									}
								}
								else
								{
	

									//delete the rx from the package
									oAdo.m_strSQL = "DELETE FROM " + this.m_oQueries.m_oFvs.m_strRxPackageFvsCmdOrderTable + " " + 
										"WHERE TRIM(RX)='" + m_oRxItem_Collection.Item(x).RxId.Trim() + "'";
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);

									//delete all rx items from the harvest cost column table
									oAdo.m_strSQL = "DELETE FROM " + this.m_oQueries.m_oFvs.m_strRxHarvestCostColumnsTable + " " + 
										"WHERE TRIM(RX)='" + m_oRxItem_Collection.Item(x).RxId.Trim() + "'";
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
								}
								

						

							}
							//update the package table with fvs commands that were added to the rx 
							this.UpdatePackageItems(oAdo);
						}
					}
				}
				
				oAdo.CloseConnection(oAdo.m_OleDbConnection);
				if (this.m_intError==0 && oAdo.m_intError==0)
				{
					this.btnSave.Enabled=false;
				}
				
			}

		
		}
		/// <summary>
		/// If an fvs command was added to the rx then the packages that contain the rx must be updated with the additional
		/// fvs command.
		/// </summary>
		/// <param name="p_oAdo"></param>
		private void UpdatePackageItems(ado_data_access p_oAdo)
		{
			//get a list packages where an rx fvs command was added
			string strRxList="";
			string strPackageList="";
			string strValues="";
			string strFields="";
			string strRx="";
			string strCurrRx="";
			string strFvsCycle="";
			string strCurrFvsCycle="";
			string strFvsCmd="";
			string strFvsCmdId="";
			string strRxPackage="";
			string strCurrRxPackage="";

			string strSavRx="";
			string strSavFvsCycle="";
			string strSavFvsCmd="";
			string strSavFvsCmdId="";
			string strSavRxPackage="";
			
			

			
			int y,x,z;
			//get the list of rx fvs commands that were added
			for (x=0;x<=m_oRxItem_Collection.Count-1;x++)
			{
				if (this.m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection != null)
				{
					for (y=0;y<=this.m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Count-1;y++)
					{
						if (m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).Add==true)
						{
							if (strRxList.IndexOf(m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).RxId.Trim(),0)<0)
									strRxList = strRxList + "'" + m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).RxId.Trim() + "',";
						}
					}
				}
			}
			if (strRxList.Trim().Length > 0)
			{
				//get the list of packages affected by the rx fvs commands added
				strRxList=strRxList.Substring(0,strRxList.Length -1);
				p_oAdo.m_strSQL = "SELECT DISTINCT rxpackage FROM " + this.m_oQueries.m_oFvs.m_strRxPackageFvsCmdOrderTable + " " + 
					              "WHERE rx IN (" + strRxList + ")";
				strPackageList = p_oAdo.CreateCommaDelimitedList(p_oAdo.m_OleDbConnection,p_oAdo.m_strSQL,"'");
				if (strPackageList.Trim().Length > 0)
				{
					

					//get a count of effected packages
					p_oAdo.m_strSQL = "SELECT COUNT(*) FROM " + this.m_oQueries.m_oFvs.m_strRxPackageFvsCmdOrderTable + " " + 
						"WHERE rxpackage IN (" + strPackageList + ")";
					int intRecTtl = Convert.ToInt32(p_oAdo.getRecordCount(p_oAdo.m_OleDbConnection,"SELECT COUNT(*) FROM " + m_oQueries.m_oFvs.m_strRxPackageFvsCmdOrderTable,"TEMP"));
				    

					if (intRecTtl > 0)
					{
						//create an array object that will hold the insert commands
						string[] strSqlArray = new string[intRecTtl + 500];
						int intSqlCount=0;
						strSqlArray.Initialize();

						//query the record set into a data reader
						p_oAdo.m_strSQL = "SELECT * FROM " + this.m_oQueries.m_oFvs.m_strRxPackageFvsCmdOrderTable + " " + 
							"WHERE rxpackage IN (" + strPackageList + ")";

						p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection,p_oAdo.m_strSQL);
						if (p_oAdo.m_OleDbDataReader.HasRows)
						{

							strFields="rxpackage,rx,fvscycle,fvscmd,fvscmd_id";
							while (p_oAdo.m_OleDbDataReader.Read())
							{
								strValues="";
								strRx="";
								strFvsCycle="";
								strFvsCmd="";
								strFvsCmdId="";
								strRxPackage="";
								strSavRx="";
								strSavFvsCycle="";
								strSavFvsCmd="";
								strSavFvsCmdId="";
								strSavRxPackage="";

								strRxPackage = p_oAdo.m_OleDbDataReader["rxpackage"].ToString().Trim();
								strFvsCmd= p_oAdo.m_OleDbDataReader["fvscmd"].ToString().Trim();
								strFvsCmdId = Convert.ToString(p_oAdo.m_OleDbDataReader["fvscmd_id"]).Trim();

								if (p_oAdo.m_OleDbDataReader["rx"] != System.DBNull.Value)
								{
									strRx = p_oAdo.m_OleDbDataReader["rx"].ToString().Trim();
								}
								if (p_oAdo.m_OleDbDataReader["fvscycle"] != System.DBNull.Value)
								{
									strFvsCycle = p_oAdo.m_OleDbDataReader["fvscycle"].ToString().Trim();
								}
								strSavRxPackage = strRxPackage;
								strSavRx = strRx;
								strSavFvsCycle=strFvsCycle;
								strSavFvsCmdId=strFvsCmdId;
								strSavFvsCmd=strFvsCmd;
								//bypass loading the rx records from the datareader since the 
								//rx records to load will be from the fvs commands collection
								if (strRx.Trim().Length  > 0 && 
									strFvsCycle.Trim().Length > 0 && 
									(strRx != strCurrRx || strFvsCycle != strCurrFvsCycle))
							
								{	

									strCurrRx=strRx;
									strCurrRxPackage = strRxPackage;
									strCurrFvsCycle=strFvsCycle;
									//update all the fvs commands for the rx by using the fvs commands collection									
									for (x=0;x<=m_oRxItem_Collection.Count-1;x++)
									{
										//find the current rx
										if (this.m_oRxItem_Collection.Item(x).RxId.Trim() == strCurrRx.Trim())
										{
											//find the fvs commands to add
											for (z=0;z<=m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Count-1;z++)
											{
												for (y=0;y<=this.m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Count-1;y++)
												{
													
													if (m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).Delete==false)
													{
														//load them up in index order
														if (m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).Index == z)
														{
															//insert added record
															strRxPackage =strCurrRxPackage;
															strRx=strCurrRx;
															strFvsCycle=strCurrFvsCycle;
															strFvsCmd=m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommand.Trim();
															strFvsCmdId = Convert.ToString(m_oRxItem_Collection.Item(x).ReferenceFvsCommandsCollection.Item(y).FVSCommandId);
															strValues="'" + strCurrRxPackage + "',";
															strValues=strValues + "'" + strCurrRx + "',";
															strValues=strValues + "'" + strCurrFvsCycle + "',";
															strValues=strValues + "'" + strFvsCmd + "',";
															strValues=strValues +  strFvsCmdId;
															p_oAdo.m_strSQL = Queries.GetInsertSQL(strFields,strValues,m_oQueries.m_oFvs.m_strRxPackageFvsCmdOrderTable);
															strSqlArray[intSqlCount] = p_oAdo.m_strSQL;
															intSqlCount++;
														}
													}
												}
											}
										}
									}
										
								}
								
								if (strRx.Trim().Length == 0)
								{
									//insert current package fvs command
									strValues="'" + strSavRxPackage + "',";
									strValues=strValues + "'" + strSavRx + "',";
									strValues=strValues + "'" + strSavFvsCycle + "',";
									strValues=strValues + "'" + strSavFvsCmd + "',";
									strValues=strValues +  strSavFvsCmdId;
									p_oAdo.m_strSQL = Queries.GetInsertSQL(strFields,strValues,m_oQueries.m_oFvs.m_strRxPackageFvsCmdOrderTable);
									strSqlArray[intSqlCount] = p_oAdo.m_strSQL;
									intSqlCount++;
								}
								
							}
							p_oAdo.m_OleDbDataReader.Close();
							//delete all the packages listed in the strPackageList variable from the table
							p_oAdo.m_strSQL = "DELETE FROM " + this.m_oQueries.m_oFvs.m_strRxPackageFvsCmdOrderTable + " " + 
								"WHERE rxpackage IN (" + strPackageList + ")";
							p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection,p_oAdo.m_strSQL);

							//insert the packages back into the table that now list the additional rx fvs commands
							for (x=0;x<=strSqlArray.Length - 1;x++)
							{
								if (strSqlArray[x] != null)
								{
									if (strSqlArray[x].Trim().Length > 0)
									{
										p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection,strSqlArray[x]);
									}
									else break;
								}
								else break;
							}
						}
						

					}
					
				}
				
			}
			

			
		}
		


		private void btnDelete_Click(object sender, System.EventArgs e)
		{
			if (this.lstRx.SelectedItems.Count==0) return;
			int x;
			for (x=0;x<=m_oRxItem_Collection.Count-1;x++)
			{
				if (m_oRxItem_Collection.Item(x).RxId.Trim()==
					lstRx.SelectedItems[0].SubItems[COLUMN_RX].Text.Trim())
				{
					//m_oRxItem_Collection.Remove(x);
					m_oRxItem_Collection.Item(x).Delete=true;
					break;
				}
			}

			int intSelected = this.lstRx.SelectedItems[0].Index;
			this.m_oLvAlternateColors.m_oRowCollection.Remove(intSelected);
			this.m_oLvAlternateColors.m_intSelectedRow=-1;
			this.lstRx.SelectedItems[0].Remove();
			if (this.lstRx.Items.Count > 0)
			{
				if (intSelected == 0)
				{
					this.lstRx.Items[intSelected].Selected=true;
				}
				else if (intSelected == this.lstRx.Items.Count-1)
				{
					this.lstRx.Items[intSelected].Selected=true;
				}
				else
				{
					this.lstRx.Items[intSelected-1].Selected=true;
				}
				
			}
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			this.lstRx.Focus();
		}

		private void lstRx_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.btnDelete.Enabled==false) this.btnDelete.Enabled=true;
			if (this.btnEdit.Enabled==false) this.btnEdit.Enabled=true;

			if (this.lstRx.SelectedItems.Count > 0)
				this.m_oLvAlternateColors.DelegateListViewItem(lstRx.SelectedItems[0]);
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{

			this.ParentForm.Close();

			
		}

		private void uc_rx_list_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
				this.lstRx.Left = 5;
				this.lstRx.Width = this.Width - 10;

				
				
				this.lstRx.Height = this.btnClose.Top - this.lstRx.Top - (this.btnEdit.Height * 2);
				
				this.btnDefault.Left = (int)(this.Width * .5) - (int)((this.btnDefault.Width + 
					                                                  this.btnNew.Width + 
																	  this.btnEdit.Width + 
																	  this.btnClear.Width + 
																	  this.btnProperties.Width + 
																	  this.btnSave.Width + 
																	  this.btnDelete.Width + 
																	  this.btnCancel.Width) * .5);

				this.btnNew.Left = this.btnDefault.Left + this.btnDefault.Width;
				this.btnEdit.Left = this.btnNew.Left + this.btnNew.Width;
				this.btnDelete.Left = this.btnEdit.Left + this.btnEdit.Width;

				this.btnEdit.Top = this.lstRx.Top + this.lstRx.Height + 5;
				this.btnClear.Left = this.btnDelete.Left + this.btnDelete.Width;
				this.btnProperties.Left = this.btnClear.Left + this.btnClear.Width;
				this.btnSave.Left = this.btnProperties.Left + this.btnProperties.Width;
				this.btnCancel.Left = this.btnSave.Left + this.btnSave.Width;


				
				this.btnCancel.Top = this.btnEdit.Top;
				this.btnProperties.Top = this.btnEdit.Top;
				this.btnSave.Top = this.btnEdit.Top;
				this.btnClear.Top = this.btnEdit.Top;
				this.btnDelete.Top = this.btnEdit.Top;
				this.btnNew.Top =this.btnEdit.Top;
				this.btnDefault.Top = this.btnEdit.Top;
				
				this.btnHelp.Top = this.btnClose.Top;
				this.btnHelp.Left = this.lstRx.Left;
			}
			catch
			{
			}
		}


		private void lstRx_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = lstRx.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lstRx.Items[lstRx.TopItem.Index + (int)dblRow-1].Selected=true;
				}
			}
			catch 
			{
			}
		}
		private string getFvsCategoryDescription(ado_data_access p_oAdo,string p_strRxId)
		{
			
			string strDesc = p_oAdo.getSingleStringValueFromSQLQuery(p_oAdo.m_OleDbConnection,this.m_oQueries.m_oFvs.GetRxItemCategoryDescriptionSQL(p_strRxId),"temp");
			
			return strDesc;
		}
		private string getFvsSubCategoryDescription(ado_data_access p_oAdo, string p_strRxId)
		{
			
			
			string strDesc = p_oAdo.getSingleStringValueFromSQLQuery(p_oAdo.m_OleDbConnection,this.m_oQueries.m_oFvs.GetRxItemSubCategoryDescriptionSQL(p_strRxId),"temp");
			
			return strDesc;
		}

		//Note: The HTML properties page takes too long to load; This has been replaced by
        //RxTools.TreatmentProperties()
        private void btnProperties_ClickHtml(object sender, System.EventArgs e)
		{
			if  (this.lstRx.Items.Count==0) return;

			frmMain.g_sbpInfo.Text = "Creating Rx Properties Report...Stand By";				
			FIA_Biosum_Manager.RxItem_Collection oColl = new RxItem_Collection();
			for (int x=0;x<=m_oRxItem_Collection.Count-1;x++)
			{
				if (m_oRxItem_Collection.Item(x).Delete==false)
				{
					RxItem oItem = new RxItem();
					oItem.CopyProperties(m_oRxItem_Collection.Item(x),oItem);
					if (oItem.m_oFvsCommandItem_Collection1 != null)
					{
						for (int y=0;y<=oItem.m_oFvsCommandItem_Collection1.Count-1;y++)
						{
							if (oItem.m_oFvsCommandItem_Collection1.Item(y).Delete==true)
							{
								oItem.m_oFvsCommandItem_Collection1.Remove(y);
							}
						}
					}
					oColl.Add(oItem);

				}
			}
			
			FIA_Biosum_Manager.project_properties_html_report oRpt = new project_properties_html_report();
			oRpt.ProcessTreatments=true;
			oRpt.ProcessPackages=false;
			oRpt.RxCollection = oColl;
			oRpt.ReportHeader = "FIA Biosum Treatments";
			oRpt.WindowTitle = "FIA Biosum Treatment Properties";
			oRpt.ProjectName = frmMain.g_oFrmMain.frmProject.uc_project1.txtProjectId.Text;
			oRpt.CreateReport();
			frmMain.g_sbpInfo.Text = "Ready";


		}

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "FVS", "RX_TREATMENT_LIST" });
        }

        private void btnProperties_Click(object sender, EventArgs e)
        {
            if (this.lstRx.Items.Count == 0) return;

            frmMain.g_sbpInfo.Text = "Creating Rx Properties Report...Stand By";
            FIA_Biosum_Manager.RxItem_Collection oColl = new RxItem_Collection();
            for (int x = 0; x <= m_oRxItem_Collection.Count - 1; x++)
            {
                if (m_oRxItem_Collection.Item(x).Delete == false)
                {
                    RxItem oItem = new RxItem();
                    oItem.CopyProperties(m_oRxItem_Collection.Item(x), oItem);
                    if (oItem.m_oFvsCommandItem_Collection1 != null)
                    {
                        for (int y = 0; y <= oItem.m_oFvsCommandItem_Collection1.Count - 1; y++)
                        {
                            if (oItem.m_oFvsCommandItem_Collection1.Item(y).Delete == true)
                            {
                                oItem.m_oFvsCommandItem_Collection1.Remove(y);
                            }
                        }
                    }
                    oColl.Add(oItem);

                }
            }
            
            FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
            frmTemp.Text = "FIA Biosum";
            frmTemp.AutoScroll = false;
            uc_textbox uc_textbox1 = new uc_textbox();
            frmTemp.Controls.Add(uc_textbox1);
            uc_textbox1.Dock = DockStyle.Fill;
            uc_textbox1.lblTitle.Text = "Treatment Properties";
            uc_textbox1.TextValue = m_oRxTools.TreatmentProperties(oColl);
            frmTemp.Show();

        }
		
		
		
	}
	public class RxTools
	{
		public int m_intError=0;
		public string m_strError="";

		public RxTools()
		{
		}

		public void LoadAllRxItemsFromTableIntoRxCollection(string p_strDbFile,Queries p_oQueries,RxItem_Collection p_oRxItemCollection)
		{
			ado_data_access oAdo = new ado_data_access();
			oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile,"",""));
			if (oAdo.m_intError==0)
			{ 
				this.LoadAllRxItemsFromTableIntoRxCollection(oAdo,oAdo.m_OleDbConnection,p_oQueries,p_oRxItemCollection);
			}
			m_intError=oAdo.m_intError;
			oAdo.CloseConnection(oAdo.m_OleDbConnection);
			oAdo=null;
		}
        /// <summary>
        /// Load the user defined Treatment into the Rx Collection.
        /// Variables Loaded:
        /// 1. rx - treatment (3 number id)
        /// 2. description
        /// 3. catid - category id
        /// 4. subcatid - subcategory id
        /// 5. Harvest Method for Low Slope
        /// 6. Harvest Method for Steep Slope
        /// 7. FVS Commands representing this RX are loaded into fvs commands collection
        /// 8. Additional RX Harvest Costs that are added after FRCS is run.
        /// 9. List of RX Packages that this RX is a member.
        /// </summary>
        /// <param name="p_oAdo"></param>
        /// <param name="p_oConn"></param>
        /// <param name="p_oQueries"></param>
        /// <param name="p_oRxItemCollection"></param>
		public void LoadAllRxItemsFromTableIntoRxCollection(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,Queries p_oQueries,RxItem_Collection p_oRxItemCollection)
		{
			
			int x;   	
			int intIndex=0;
			p_oAdo.m_strSQL = "SELECT * FROM " + p_oQueries.m_oFvs.m_strRxTable;
			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				while (p_oAdo.m_OleDbDataReader.Read())
				{
						
					if (p_oAdo.m_OleDbDataReader["rx"] != System.DBNull.Value)
					{
							
						RxItem oRxItem = new RxItem();
						oRxItem.m_oFvsCommandItem_Collection1 = new RxItemFvsCommandItem_Collection();
						oRxItem.ReferenceFvsCommandsCollection = oRxItem.m_oFvsCommandItem_Collection1;
						oRxItem.m_oHarvestCostColumnItem_Collection1 = new RxItemHarvestCostColumnItem_Collection();
						oRxItem.ReferenceHarvestCostColumnCollection=oRxItem.m_oHarvestCostColumnItem_Collection1;
						oRxItem.Index = intIndex;
						oRxItem.RxId = Convert.ToString(p_oAdo.m_OleDbDataReader["rx"]);
							
						if (p_oAdo.m_OleDbDataReader["description"] != System.DBNull.Value)
						{
							oRxItem.Description = Convert.ToString(p_oAdo.m_OleDbDataReader["description"]);
						}
						if (p_oAdo.m_OleDbDataReader["catid"] != System.DBNull.Value &&
                            p_oQueries.m_oFvs.m_strFvsCatTable.Trim().Length > 0)
						{
							oRxItem.Category = p_oQueries.m_oFvs.GetCategoryDescriptionFromCategoryIdSQL(Convert.ToString(p_oAdo.m_OleDbDataReader["catid"]).Trim());
						}
						if (p_oAdo.m_OleDbDataReader["subcatid"] != System.DBNull.Value &&
                            p_oQueries.m_oFvs.m_strFvsSubCatTable.Trim().Length > 0)
						{
							oRxItem.SubCategory = p_oQueries.m_oFvs.GetSubCategoryDescriptionFromCategoryIdAndSubCategoryIdSQL(
								Convert.ToString(p_oAdo.m_OleDbDataReader["catid"]).Trim(),
								Convert.ToString(p_oAdo.m_OleDbDataReader["subcatid"]).Trim());
						}
                        if (p_oAdo.m_OleDbDataReader["HarvestMethodLowSlope"] != System.DBNull.Value)
						{
                            oRxItem.HarvestMethodLowSlope = Convert.ToString(p_oAdo.m_OleDbDataReader["HarvestMethodLowSlope"]).Trim();
						}
						if (p_oAdo.m_OleDbDataReader["HarvestMethodSteepSlope"] != System.DBNull.Value)
						{
							oRxItem.HarvestMethodSteepSlope = Convert.ToString(p_oAdo.m_OleDbDataReader["HarvestMethodSteepSlope"]).Trim();
						}
						
						p_oRxItemCollection.Add(oRxItem);
							
							
						intIndex++;
							


					}
				}
					
			}
				
			intIndex=0;
			p_oAdo.m_OleDbDataReader.Close();
			//
			//FVS COMMANDS
			//
            if (p_oQueries.m_oFvs.m_strRxFvsCmdTable.Trim().Length > 0)
            {
                p_oAdo.m_strSQL = "SELECT * FROM " + p_oQueries.m_oFvs.m_strRxFvsCmdTable;
                p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);
                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        if (p_oAdo.m_OleDbDataReader["rx"] != System.DBNull.Value &&
                            p_oAdo.m_OleDbDataReader["fvscmd"] != System.DBNull.Value)
                        {
                            for (x = 0; x <= p_oRxItemCollection.Count - 1; x++)
                            {
                                if (p_oRxItemCollection.Item(x).RxId ==
                                    p_oAdo.m_OleDbDataReader["rx"].ToString().Trim())
                                {
                                    FIA_Biosum_Manager.RxItemFvsCommandItem oItem = new RxItemFvsCommandItem();
                                    oItem.Index = p_oRxItemCollection.Item(x).m_oFvsCommandItem_Collection1.Count;
                                    oItem.SaveIndex = p_oRxItemCollection.Item(x).m_oFvsCommandItem_Collection1.Count;
                                    oItem.RxId = p_oRxItemCollection.Item(x).RxId;
                                    oItem.FVSCommand = p_oAdo.m_OleDbDataReader["fvscmd"].ToString().Trim();
                                    oItem.FVSCommandId = Convert.ToByte(p_oAdo.m_OleDbDataReader["fvscmd_id"]);
                                    if (p_oAdo.m_OleDbDataReader["p1"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter1 = p_oAdo.m_OleDbDataReader["p1"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["p2"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter2 = p_oAdo.m_OleDbDataReader["p2"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["p3"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter3 = p_oAdo.m_OleDbDataReader["p3"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["p4"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter4 = p_oAdo.m_OleDbDataReader["p4"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["p5"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter5 = p_oAdo.m_OleDbDataReader["p5"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["p6"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter6 = p_oAdo.m_OleDbDataReader["p6"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["p7"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter7 = p_oAdo.m_OleDbDataReader["p7"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["other"] != System.DBNull.Value)
                                    {
                                        oItem.Other = p_oAdo.m_OleDbDataReader["other"].ToString();
                                    }
                                    p_oRxItemCollection.Item(x).m_oFvsCommandItem_Collection1.Add(oItem);
                                    intIndex++;
                                }
                            }

                        }
                    }
                }
                p_oAdo.m_OleDbDataReader.Close();
            }
			intIndex=0;
			//
			//HARVEST COST COLUMNS
			//
			p_oAdo.m_strSQL = "SELECT * FROM " + p_oQueries.m_oFvs.m_strRxHarvestCostColumnsTable;
			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				while (p_oAdo.m_OleDbDataReader.Read())
				{
					if (p_oAdo.m_OleDbDataReader["rx"] != System.DBNull.Value &&
						p_oAdo.m_OleDbDataReader["ColumnName"] != System.DBNull.Value)
					{
						for (x=0;x<=p_oRxItemCollection.Count-1;x++)
						{
							if (p_oRxItemCollection.Item(x).RxId == 
								p_oAdo.m_OleDbDataReader["rx"].ToString().Trim())
							{
								FIA_Biosum_Manager.RxItemHarvestCostColumnItem oItem = new RxItemHarvestCostColumnItem();
								oItem.Index = p_oRxItemCollection.Item(x).m_oHarvestCostColumnItem_Collection1.Count;
								oItem.SaveIndex = p_oRxItemCollection.Item(x).m_oHarvestCostColumnItem_Collection1.Count;
								oItem.RxId = p_oRxItemCollection.Item(x).RxId;
								oItem.HarvestCostColumn=p_oAdo.m_OleDbDataReader["ColumnName"].ToString().Trim();
								
								if (p_oAdo.m_OleDbDataReader["description"] != System.DBNull.Value)
								{
									oItem.Description = p_oAdo.m_OleDbDataReader["description"].ToString().Trim();
								}
								p_oRxItemCollection.Item(x).m_oHarvestCostColumnItem_Collection1.Add(oItem);
								intIndex++;
							}
						}
							
					}
				}
			}
			p_oAdo.m_OleDbDataReader.Close();
            //
            //PACKAGE MEMBERSHIP
            //
            for (x = 0; x <= p_oRxItemCollection.Count - 1; x++)
            {
                p_oRxItemCollection.Item(x).RxPackageMemberList = "";
            }
            p_oAdo.m_strSQL = "SELECT * FROM " + p_oQueries.m_oFvs.m_strRxPackageTable;
            p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
               
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    string strPackage = "";
                    string strCycle1Rx = "";
                    string strCycle2Rx = "";
                    string strCycle3Rx = "";
                    string strCycle4Rx = "";

                    if (p_oAdo.m_OleDbDataReader["rxpackage"] != System.DBNull.Value)
                    {
                        strPackage = p_oAdo.m_OleDbDataReader["rxpackage"].ToString().Trim();
                        if (strPackage.Trim().Length > 0)
                        {
                            
                            //sim year 1
                            if (p_oAdo.m_OleDbDataReader["simyear1_rx"] != System.DBNull.Value)
                            {
                                strCycle1Rx= Convert.ToString(p_oAdo.m_OleDbDataReader["simyear1_rx"]).Trim();
                            }
                            //sim year 2
                            if (p_oAdo.m_OleDbDataReader["simyear2_rx"] != System.DBNull.Value)
                            {
                                strCycle2Rx = Convert.ToString(p_oAdo.m_OleDbDataReader["simyear2_rx"]).Trim();
                            }
                            //sim year 3
                            if (p_oAdo.m_OleDbDataReader["simyear3_rx"] != System.DBNull.Value)
                            {
                                strCycle3Rx = Convert.ToString(p_oAdo.m_OleDbDataReader["simyear3_rx"]).Trim();
                            }
                            //sim year 4
                            if (p_oAdo.m_OleDbDataReader["simyear4_rx"] != System.DBNull.Value)
                            {
                                strCycle4Rx = Convert.ToString(p_oAdo.m_OleDbDataReader["simyear4_rx"]).Trim();
                            }
                            
                            for (x = 0; x <= p_oRxItemCollection.Count - 1; x++)
                            {
                                string strLine = "";
                                
                                if (strCycle1Rx == p_oRxItemCollection.Item(x).RxId)
                                {
                                    strLine = "1-";
                                }
                                if (strCycle2Rx == p_oRxItemCollection.Item(x).RxId)
                                {
                                    strLine = strLine + "2-";
                                }
                                if (strCycle3Rx == p_oRxItemCollection.Item(x).RxId)
                                {
                                    strLine = strLine + "3-";
                                }
                                if (strCycle4Rx == p_oRxItemCollection.Item(x).RxId)
                                {
                                    strLine = strLine + "4-";
                                }
                                if (strLine.Trim().Length > 0)
                                {
                                    strLine = strLine.Substring(0, strLine.Length - 1);
                                    strLine = "Package:" + strPackage + " Simulation Cycle(s):" + strLine + ",";

                                    p_oRxItemCollection.Item(x).RxPackageMemberList =
                                        p_oRxItemCollection.Item(x).RxPackageMemberList + strLine;

                                }
                               
                               
                            }
                            
                        }
                    }
                }

            }
            p_oAdo.m_OleDbDataReader.Close();

			for (x=0;x<=p_oRxItemCollection.Count-1;x++)
			{
				if (p_oRxItemCollection.Item(x).Category.Trim().Length > 0)
				{
					p_oRxItemCollection.Item(x).Category = p_oAdo.getSingleStringValueFromSQLQuery(
						p_oConn,p_oRxItemCollection.Item(x).Category,"temp");
				}
				if (p_oRxItemCollection.Item(x).SubCategory.Trim().Length > 0)
				{
					p_oRxItemCollection.Item(x).SubCategory = p_oAdo.getSingleStringValueFromSQLQuery(
						p_oConn,p_oRxItemCollection.Item(x).SubCategory,"temp");
				}
                //remove the last comma from the end of the string
                if (p_oRxItemCollection.Item(x).RxPackageMemberList.Trim().Length > 0)
                    p_oRxItemCollection.Item(x).RxPackageMemberList = p_oRxItemCollection.Item(x).RxPackageMemberList.Substring(0, p_oRxItemCollection.Item(x).RxPackageMemberList.Length - 1);
			}
		}

        public void LoadAllRxPackageItems(RxPackageItem_Collection p_oRxPackageItemCollection)
        {
            Queries oQueries = new Queries();
            oQueries.m_oFvs.LoadDatasource = true;
            oQueries.LoadDatasources(true);
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(oQueries.m_strTempDbFile, "", ""));
            if (oAdo.m_intError == 0)
            {
                this.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(oAdo, oAdo.m_OleDbConnection, oQueries, p_oRxPackageItemCollection);
            }
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            m_intError = oAdo.m_intError;
            oQueries = null;
            oAdo = null;


        }
		public void LoadAllRxPackageItemsFromTableIntoRxPackageCollection(string p_strDbFile,Queries p_oQueries,RxPackageItem_Collection p_oRxPackageItemCollection)
		{
			ado_data_access oAdo = new ado_data_access();
			oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile,"",""));
			if (oAdo.m_intError==0)
			{ 
				this.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(oAdo,oAdo.m_OleDbConnection,p_oQueries,p_oRxPackageItemCollection);
			}
			oAdo.CloseConnection(oAdo.m_OleDbConnection);
			m_intError=oAdo.m_intError;
			oAdo=null;
		}
	    public void LoadAllRxPackageItemsFromTableIntoRxPackageCollection(ado_data_access p_oAdo,
			                                                              System.Data.OleDb.OleDbConnection p_oConn,
			                                                              Queries p_oQueries,
			                                                              RxPackageItem_Collection p_oRxPackageItemCollection)
		{
			int x;   	
			int intIndex=0;
			p_oAdo.m_strSQL = "SELECT * FROM " + p_oQueries.m_oFvs.m_strRxPackageTable;
			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				
				while (p_oAdo.m_OleDbDataReader.Read())
				{
						
					if (p_oAdo.m_OleDbDataReader["rxpackage"] != System.DBNull.Value)
					{
							
						RxPackageItem oRxPackageItem = new RxPackageItem();
						
						oRxPackageItem.m_oFvsCommandItem_Collection1 = new RxPackageItemFvsCommandItem_Collection();
						oRxPackageItem.ReferenceFvsCommandsCollection = oRxPackageItem.m_oFvsCommandItem_Collection1;
						oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1 = new RxPackageCombinedFVSCommandsItem_Collection();
						oRxPackageItem.ReferenceRxPackageCombinedFVSCommandsItemCollection = oRxPackageItem.m_oRxPackageCombinedFVSCommandsItem_Collection1;

										
						
						oRxPackageItem.Index = intIndex;

						oRxPackageItem.RxPackageId = Convert.ToString(p_oAdo.m_OleDbDataReader["rxpackage"]);
							
						if (p_oAdo.m_OleDbDataReader["description"] != System.DBNull.Value)
						{
							oRxPackageItem.Description = Convert.ToString(p_oAdo.m_OleDbDataReader["description"]);
						}
						//treatment cycle length
						if (p_oAdo.m_OleDbDataReader["rxcycle_length"] != System.DBNull.Value)
						{
							oRxPackageItem.RxCycleLength = Convert.ToInt32(p_oAdo.m_OleDbDataReader["rxcycle_length"]);
						}
						//sim year 1
						if (p_oAdo.m_OleDbDataReader["simyear1_rx"] != System.DBNull.Value)
						{
							oRxPackageItem.SimulationYear1Rx = Convert.ToString(p_oAdo.m_OleDbDataReader["simyear1_rx"]);
						}
						if (p_oAdo.m_OleDbDataReader["simyear1_fvscycle"] != System.DBNull.Value)
						{
							oRxPackageItem.SimulationYear1Fvs = Convert.ToString(p_oAdo.m_OleDbDataReader["simyear1_fvscycle"]);
						}
						//sim year 2
						if (p_oAdo.m_OleDbDataReader["simyear2_rx"] != System.DBNull.Value)
						{
							oRxPackageItem.SimulationYear2Rx = Convert.ToString(p_oAdo.m_OleDbDataReader["simyear2_rx"]);
						}
						if (p_oAdo.m_OleDbDataReader["simyear2_fvscycle"] != System.DBNull.Value)
						{
							oRxPackageItem.SimulationYear2Fvs = Convert.ToString(p_oAdo.m_OleDbDataReader["simyear2_fvscycle"]);
						}
						//sim year 3
						if (p_oAdo.m_OleDbDataReader["simyear3_rx"] != System.DBNull.Value)
						{
							oRxPackageItem.SimulationYear3Rx = Convert.ToString(p_oAdo.m_OleDbDataReader["simyear3_rx"]);
						}
						if (p_oAdo.m_OleDbDataReader["simyear3_fvscycle"] != System.DBNull.Value)
						{
							oRxPackageItem.SimulationYear3Fvs = Convert.ToString(p_oAdo.m_OleDbDataReader["simyear3_fvscycle"]);
						}
						//sim year 4
						if (p_oAdo.m_OleDbDataReader["simyear4_rx"] != System.DBNull.Value)
						{
							oRxPackageItem.SimulationYear4Rx = Convert.ToString(p_oAdo.m_OleDbDataReader["simyear4_rx"]);
						}
						if (p_oAdo.m_OleDbDataReader["simyear4_fvscycle"] != System.DBNull.Value)
						{
							oRxPackageItem.SimulationYear4Fvs = Convert.ToString(p_oAdo.m_OleDbDataReader["simyear4_fvscycle"]);
						}
						if (p_oAdo.m_OleDbDataReader["kcpfile"] != System.DBNull.Value)
						{
							oRxPackageItem.KcpFile = Convert.ToString(p_oAdo.m_OleDbDataReader["kcpfile"]);
						}
                        if (oRxPackageItem.SimulationYear1Rx.Trim().Length == 0) oRxPackageItem.SimulationYear1Rx = "000";
                        if (oRxPackageItem.SimulationYear2Rx.Trim().Length == 0) oRxPackageItem.SimulationYear2Rx = "000";
                        if (oRxPackageItem.SimulationYear3Rx.Trim().Length == 0) oRxPackageItem.SimulationYear3Rx = "000";
                        if (oRxPackageItem.SimulationYear4Rx.Trim().Length == 0) oRxPackageItem.SimulationYear4Rx = "000";

						p_oRxPackageItemCollection.Add(oRxPackageItem);
						intIndex++;
					}
				}
				
			}
            p_oAdo.m_OleDbDataReader.Close();

            if (p_oQueries.m_oFvs.m_strRxPackageFvsCmdTable.Trim().Length > 0)
            {
                p_oAdo.m_strSQL = "SELECT * FROM " + p_oQueries.m_oFvs.m_strRxPackageFvsCmdTable;
                p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);
                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        if (p_oAdo.m_OleDbDataReader["rxpackage"] != System.DBNull.Value &&
                            p_oAdo.m_OleDbDataReader["fvscmd"] != System.DBNull.Value)
                        {
                            for (x = 0; x <= p_oRxPackageItemCollection.Count - 1; x++)
                            {
                                if (p_oRxPackageItemCollection.Item(x).RxPackageId ==
                                    p_oAdo.m_OleDbDataReader["rxpackage"].ToString().Trim())
                                {
                                    FIA_Biosum_Manager.RxPackageItemFvsCommandItem oItem = new RxPackageItemFvsCommandItem();
                                    oItem.Index = p_oRxPackageItemCollection.Item(x).m_oFvsCommandItem_Collection1.Count;
                                    oItem.SaveIndex = p_oRxPackageItemCollection.Item(x).m_oFvsCommandItem_Collection1.Count;
                                    oItem.RxPackageId = p_oRxPackageItemCollection.Item(x).RxPackageId;
                                    oItem.FVSCommand = p_oAdo.m_OleDbDataReader["fvscmd"].ToString().Trim();
                                    oItem.FVSCommandId = Convert.ToByte(p_oAdo.m_OleDbDataReader["fvscmd_id"]);
                                    oItem.ListViewIndex = Convert.ToInt32(p_oAdo.m_OleDbDataReader["list_index"]);


                                    if (p_oAdo.m_OleDbDataReader["p1"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter1 = p_oAdo.m_OleDbDataReader["p1"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["p2"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter2 = p_oAdo.m_OleDbDataReader["p2"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["p3"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter3 = p_oAdo.m_OleDbDataReader["p3"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["p4"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter4 = p_oAdo.m_OleDbDataReader["p4"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["p5"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter5 = p_oAdo.m_OleDbDataReader["p5"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["p6"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter6 = p_oAdo.m_OleDbDataReader["p6"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["p7"] != System.DBNull.Value)
                                    {
                                        oItem.Parameter7 = p_oAdo.m_OleDbDataReader["p7"].ToString();
                                    }
                                    if (p_oAdo.m_OleDbDataReader["other"] != System.DBNull.Value)
                                    {
                                        oItem.Other = p_oAdo.m_OleDbDataReader["other"].ToString();
                                    }
                                    p_oRxPackageItemCollection.Item(x).m_oFvsCommandItem_Collection1.Add(oItem);
                                    intIndex++;
                                }
                            }

                        }
                    }
                }
                p_oAdo.m_OleDbDataReader.Close();
            }
            if (p_oQueries.m_oFvs.m_strRxPackageFvsCmdOrderTable.Trim().Length > 0)
            {
                p_oAdo.m_strSQL = "SELECT * FROM " + p_oQueries.m_oFvs.m_strRxPackageFvsCmdOrderTable;
                p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);
                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        if (p_oAdo.m_OleDbDataReader["rxpackage"] != System.DBNull.Value &&
                            p_oAdo.m_OleDbDataReader["fvscmd"] != System.DBNull.Value)
                        {
                            for (x = 0; x <= p_oRxPackageItemCollection.Count - 1; x++)
                            {
                                if (p_oRxPackageItemCollection.Item(x).RxPackageId ==
                                    p_oAdo.m_OleDbDataReader["rxpackage"].ToString().Trim())
                                {
                                    FIA_Biosum_Manager.RxPackageCombinedFVSCommandsItem oItem = new RxPackageCombinedFVSCommandsItem();
                                    oItem.Index = p_oRxPackageItemCollection.Item(x).m_oRxPackageCombinedFVSCommandsItem_Collection1.Count;
                                    oItem.SaveIndex = p_oRxPackageItemCollection.Item(x).m_oRxPackageCombinedFVSCommandsItem_Collection1.Count;
                                    oItem.RxPackageId = p_oRxPackageItemCollection.Item(x).RxPackageId;
                                    oItem.FVSCommand = p_oAdo.m_OleDbDataReader["fvscmd"].ToString().Trim();
                                    oItem.FVSCommandId = Convert.ToByte(p_oAdo.m_OleDbDataReader["fvscmd_id"]);
                                    if (p_oAdo.m_OleDbDataReader["rx"] != System.DBNull.Value)
                                    {
                                        oItem.RxId = Convert.ToString(p_oAdo.m_OleDbDataReader["rx"]);
                                    }
                                    if (p_oAdo.m_OleDbDataReader["fvscycle"] != System.DBNull.Value)
                                    {
                                        oItem.FVSCycle = Convert.ToString(p_oAdo.m_OleDbDataReader["fvscycle"]);
                                    }


                                    p_oRxPackageItemCollection.Item(x).m_oRxPackageCombinedFVSCommandsItem_Collection1.Add(oItem);
                                    intIndex++;
                                }
                            }

                        }
                    }
                }
                p_oAdo.m_OleDbDataReader.Close();
            }

			
	
			
		}
		public void LoadAllRxPackageCombinedFvsCommandsIntoCollection(string p_strDbFile,
			Queries p_oQueries,
			FIA_Biosum_Manager.RxPackageCombinedFVSCommandsItem_Collection p_oCombinedFvsCommandsCollection)
		{
			ado_data_access oAdo = new ado_data_access();
			oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile,"",""));
			if (oAdo.m_intError==0)
			{ 
				this.LoadAllRxPackageCombinedFvsCommandsIntoCollection(oAdo,
					oAdo.m_OleDbConnection,p_oQueries,
					p_oCombinedFvsCommandsCollection);
			}
			m_intError=oAdo.m_intError;
			oAdo.CloseConnection(oAdo.m_OleDbConnection);
			oAdo=null;
		}
		public void LoadAllRxPackageCombinedFvsCommandsIntoCollection(ado_data_access p_oAdo,
			System.Data.OleDb.OleDbConnection p_oConn,
			Queries p_oQueries,
			RxPackageCombinedFVSCommandsItem_Collection p_oCombinedFvsCommandsCollection)
		{
			
			int x;   	
			
			p_oAdo.m_strSQL = "SELECT * FROM " + p_oQueries.m_oFvs.m_strRxPackageFvsCmdOrderTable;
				             
			p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
			if (p_oAdo.m_OleDbDataReader.HasRows)
			{
				x=0;
				while (p_oAdo.m_OleDbDataReader.Read())
				{
						
					if (p_oAdo.m_OleDbDataReader["rxpackage"] != System.DBNull.Value)
					{
						FIA_Biosum_Manager.RxPackageCombinedFVSCommandsItem oItem = new RxPackageCombinedFVSCommandsItem();	
						oItem.Index = x;
						oItem.SaveIndex = oItem.Index;
						oItem.RxPackageId = Convert.ToString(p_oAdo.m_OleDbDataReader["rxpackage"]);
						x++;
							
						
						if (p_oAdo.m_OleDbDataReader["rx"] != System.DBNull.Value)
						{
							oItem.RxId = Convert.ToString(p_oAdo.m_OleDbDataReader["rx"]);
						}
						if (p_oAdo.m_OleDbDataReader["fvscmd"] != System.DBNull.Value)
						{
							oItem.FVSCommand = Convert.ToString(p_oAdo.m_OleDbDataReader["fvscmd"]);
						}
						if (p_oAdo.m_OleDbDataReader["fvscmd_id"] != System.DBNull.Value)
						{
							oItem.FVSCommandId = Convert.ToByte(p_oAdo.m_OleDbDataReader["fvscmd_id"]);
						}
						if (p_oAdo.m_OleDbDataReader["fvscycle"] != System.DBNull.Value)
						{
							oItem.FVSCycle = Convert.ToString(p_oAdo.m_OleDbDataReader["fvscycle"]);
						}


												
						p_oCombinedFvsCommandsCollection.Add(oItem);
										


					}
				}
					
			}
				
		    p_oAdo.m_OleDbDataReader.Close();
		
		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="p_strDbFile"></param>
		/// <param name="p_strRxPackageTable"></param>
		/// <param name="p_strPlotTable"></param>
		/// <param name="p_strFilter"></param>
		/// <returns></returns>
		public string GetRxPackageFvsOutDbFileName(string p_strDbFile,
												   string p_strRxPackageTable,
			                                       string p_strPlotTable,
			                                       string p_strFilter)
		{
			string strFile="";
			ado_data_access oAdo = new ado_data_access();
			oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile,"",""));
			if (oAdo.m_intError==0)
			{
				strFile = this.GetRxPackageFvsOutDbFileName(oAdo.m_OleDbConnection,p_strRxPackageTable,p_strPlotTable,p_strFilter);
				oAdo.CloseConnection(oAdo.m_OleDbConnection);
			}
			oAdo=null;
			
			return strFile;
			
		}
		/// <summary>
		/// 
		/// </summary>
		/// <param name="p_oConn"></param>
		/// <param name="p_strRxPackageTable"></param>
		/// <param name="p_strPlotTable"></param>
		/// <param name="p_strFilter"></param>
		/// <returns></returns>
		public string GetRxPackageFvsOutDbFileName(System.Data.OleDb.OleDbConnection p_oConn,
													string p_strRxPackageTable,
													string p_strPlotTable,
													string p_strFilter)
		{
			string strFile="";
			ado_data_access oAdo = new ado_data_access();
			oAdo.m_strSQL = Queries.FVS.GetFVSVariantRxPackageSQL(p_strPlotTable,p_strRxPackageTable,p_strFilter);
			oAdo.SqlQueryReader(p_oConn,oAdo.m_strSQL);
			if (oAdo.m_intError==0)
			{
				if (oAdo.m_OleDbDataReader.HasRows)
				{
					oAdo.m_OleDbDataReader.Read();
					strFile = this.GetRxPackageFvsOutDbFileName(oAdo.m_OleDbDataReader);
				}
				oAdo.m_OleDbDataReader.Close();
			}
			oAdo=null;

			return strFile;

		}
		public string GetRxPackageFvsOutDbFileName(System.Data.OleDb.OleDbDataReader p_OleDbReader)
		{

			string strFile = "FVSOUT_" + p_OleDbReader["fvs_variant"].ToString().Trim() + "_P" + p_OleDbReader["rxpackage"].ToString().Trim();
			if (p_OleDbReader["simyear1_rx"] == System.DBNull.Value ||
				p_OleDbReader["simyear1_rx"].ToString().Trim().Length == 0 ||
                p_OleDbReader["simyear1_rx"].ToString().Trim() == "GP")
			{
				strFile = strFile + "-000";
			}
			else
			{
				strFile = strFile + "-" + p_OleDbReader["simyear1_rx"].ToString().Trim();
			}
			if (p_OleDbReader["simyear2_rx"] == System.DBNull.Value ||
				p_OleDbReader["simyear2_rx"].ToString().Trim().Length == 0 ||
                p_OleDbReader["simyear2_rx"].ToString().Trim() == "GP")
			{
				strFile = strFile + "-000";
			}
			else
			{
				//GP if (p_OleDbReader["simyear2_rx"].ToString().Trim()=="GP")
				//GP {
				//GP 	strFile = strFile + "-0GP";
				//GP }
				//GP else
				//GP {
					strFile = strFile + "-" + p_OleDbReader["simyear2_rx"].ToString().Trim();
				//GP }
					
			}
			if (p_OleDbReader["simyear3_rx"] == System.DBNull.Value ||
				p_OleDbReader["simyear3_rx"].ToString().Trim().Length == 0 ||
                p_OleDbReader["simyear3_rx"].ToString().Trim() == "GP")
			{
				strFile = strFile + "-000";
			}
			else
			{
				//GP if (p_OleDbReader["simyear3_rx"].ToString().Trim()=="GP")
				//GP {
				//GP	strFile = strFile + "-0GP";
				//GP }
				//GP else
				//GP {
					strFile = strFile + "-" + p_OleDbReader["simyear3_rx"].ToString().Trim();
				//GP }
					
			}
			if (p_OleDbReader["simyear4_rx"] == System.DBNull.Value ||
				p_OleDbReader["simyear4_rx"].ToString().Trim().Length == 0 ||
                p_OleDbReader["simyear4_rx"].ToString().Trim() == "GP")
			{
				strFile = strFile + "-000";
			}
			else
			{
				//GP if (p_OleDbReader["simyear4_rx"].ToString().Trim()=="GP")
				//GP {
				//GP	strFile = strFile + "-0GP";
				//GP }
				//GP else
				//GP {
					strFile = strFile + "-" + p_OleDbReader["simyear4_rx"].ToString().Trim();
				//GP }
					
			}
			strFile=strFile + ".MDB";
			return strFile;

		}
		public byte AssignFvsCommandId(FIA_Biosum_Manager.RxItemFvsCommandItem_Collection p_FvsCollection, string p_strFvsCmd)
		{
			byte id=1;
			for (int x=0;x<=p_FvsCollection.Count-1;x++)
			{
				if (p_FvsCollection.Item(x).FVSCommand.Trim().ToUpper() == p_strFvsCmd.Trim().ToUpper())
				{
					if (p_FvsCollection.Item(x).FVSCommandId==id)
					{
						id=Convert.ToByte(id+1);
						x=-1;
					}
					else 
					{
					}
			    }
			}
			return id;

		}
		public byte AssignFvsCommandId(FIA_Biosum_Manager.RxPackageItemFvsCommandItem_Collection p_FvsCollection, string p_strFvsCmd)
		{
			byte id=1;
			for (int x=0;x<=p_FvsCollection.Count-1;x++)
			{
				if (p_FvsCollection.Item(x).FVSCommand.Trim().ToUpper() == p_strFvsCmd.Trim().ToUpper())
				{
					if (p_FvsCollection.Item(x).FVSCommandId==id)
					{
						id=Convert.ToByte(id+1);
						x=-1;
					}
					else 
					{
					}
				}
			}
			return id;
		}


		public string GetUsedRxPackageList(FIA_Biosum_Manager.RxPackageItem_Collection p_oRxPackageItemCollection)
		{
			int x;
			string strRxPackageList="";

			if (p_oRxPackageItemCollection != null)
			{
				for (x=0;x<=p_oRxPackageItemCollection.Count-1;x++)
				{
					strRxPackageList = strRxPackageList + p_oRxPackageItemCollection.Item(x).RxPackageId + ",";
				}
				if (strRxPackageList.Trim().Length > 0) strRxPackageList = strRxPackageList.Substring(0,strRxPackageList.Length - 1);
			}
			return strRxPackageList;
		}
		public string GetUsedRxList(FIA_Biosum_Manager.RxItem_Collection p_oRxItemCollection)
		{
			int x;
			string strRxList="";

			if (p_oRxItemCollection != null)
			{
				for (x=0;x<=p_oRxItemCollection.Count-1;x++)
				{
					strRxList = strRxList + p_oRxItemCollection.Item(x).RxId + ",";
				}
				if (strRxList.Trim().Length > 0) strRxList = strRxList.Substring(0,strRxList.Length - 1);
			}
			return strRxList;
		}
		public void CopyRxItemsToPackageItem(FIA_Biosum_Manager.RxItem_Collection p_RxItemCollection, FIA_Biosum_Manager.RxPackageItem p_oRxPackageItem)
		{

		}
		public void CreateTableLinksToFVSOutTreeListTables(Queries p_oQueries,string p_strDbFile)
		{
			
			ado_data_access oAdo = new ado_data_access();
			oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile,"",""));
			CreateTableLinksToFVSOutTreeListTables(oAdo,oAdo.m_OleDbConnection,p_oQueries,p_strDbFile);
			oAdo.CloseConnection(oAdo.m_OleDbConnection);
		}
		public string GetListOfFVSVariantsInPlotTable(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strPlotTable)
		{
			return p_oAdo.CreateCommaDelimitedList(p_oConn,"SELECT DISTINCT fvs_variant FROM " +  p_strPlotTable +  " WHERE fvs_variant IS NOT NULL","");
		}
        /// <summary>
        /// Create links to all the FVS Output tables for every variant and package
        /// </summary>
        /// <param name="p_oAdo"></param>
        /// <param name="p_oConn"></param>
        /// <param name="p_oQueries">Class that provides query access as well as creating a temp DB file that contains all the needed table links</param>
        /// <param name="p_strDbFile">Access database file that is the target file to contain all the table links</param>
		public void CreateTableLinksToFVSOutTreeListTables(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,Queries p_oQueries,string p_strDbFile)
		{
            RxPackageItem_Collection oRxPackageItemCollection = new RxPackageItem_Collection();
			
			string strVariantsList = GetListOfFVSVariantsInPlotTable(p_oAdo,p_oAdo.m_OleDbConnection,p_oQueries.m_oFIAPlot.m_strPlotTable);
            this.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(p_oAdo, p_oConn, p_oQueries, oRxPackageItemCollection);
			string[] strVariantsArray=frmMain.g_oUtils.ConvertListToArray(strVariantsList,",");
			dao_data_access oDao = new dao_data_access();
			//string strTreeListDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()  + "\\fvs\\data\\TreeLists";
			string strTreeListLinkTableName="";
            if (strVariantsArray != null && strVariantsArray[0] != null)
            {
                for (int x = 0; x <= strVariantsArray.Length - 1; x++)
                {
                    string strTreeListDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\data\\" + strVariantsArray[x].Trim() + "\\BiosumCalc";
                    for (int y = 0; y <= oRxPackageItemCollection.Count - 1; y++)
                    {

                        strTreeListLinkTableName = strVariantsArray[x].Trim() + "_P" + oRxPackageItemCollection.Item(y).RxPackageId + "_TREE_CUTLIST";

                        if (System.IO.File.Exists(strTreeListDir + "\\" + strTreeListLinkTableName.Trim() + ".MDB"))
                        {
                            oDao.CreateTableLink(p_strDbFile, "fvs_tree_IN_" + strTreeListLinkTableName.Trim(), strTreeListDir + "\\" + strTreeListLinkTableName.Trim() + ".MDB", "fvs_tree", true);
                        }
                    }

                }
            }
            oDao.m_DaoWorkspace.Close();
            oRxPackageItemCollection.Clear();
            oRxPackageItemCollection = null;
			oDao=null;
		}

        /// <summary>
        /// Create table links to FVS Output tables
        /// </summary>
        /// <param name="p_strDestinationDbFile">Database file to store the links. If the file does not exist then it will be created.</param>
        /// <param name="p_strSourceFVSOutputDbFile">FVS Output Database file containing the FVS Output tables</param>
        public void CreateFVSOutputTableLinks(string p_strDestinationDbFile, string p_strSourceFVSOutputDbFile)
        {
            m_intError = 0;
            m_strError = "";
            if (System.IO.File.Exists(p_strSourceFVSOutputDbFile) == false)
            {
                m_intError = -1;
                m_strError = "!!Error!!\r\n" +
                            "RxTools:CreateFVSOutputTableLinks\r\n" +
                            p_strSourceFVSOutputDbFile + "\r\n file does not exist";
                return;
            }
            dao_data_access oDao = new dao_data_access();
            string[] strTableArray=null;
            int  z;
            if (!System.IO.File.Exists(p_strDestinationDbFile))
            {
                oDao.CreateMDB(p_strDestinationDbFile);
            }
            oDao.getTableNames(p_strSourceFVSOutputDbFile, ref strTableArray);
            oDao.OpenDb(p_strDestinationDbFile);
            for (z = 0; z <= strTableArray.Length - 1; z++)
            {
                if (strTableArray[z] == null) break;
                if (RxTools.ValidFVSTable(strTableArray[z].Trim().ToUpper()))
                {

                    if (oDao.TableExists(oDao.m_DaoDatabase, strTableArray[z]))
                        oDao.DeleteTableFromMDB(oDao.m_DaoDatabase, strTableArray[z]);
                    oDao.CreateTableLink(oDao.m_DaoDatabase, strTableArray[z].Trim(), p_strSourceFVSOutputDbFile, strTableArray[z].Trim());


                }
            }
            oDao.m_DaoDatabase.Close();
            oDao.m_DaoWorkspace.Close();
            oDao = null;
        }

        /// <summary>
        /// Create links to all the FVS Output FVS_SUMMARY tables for every variant and package
        /// </summary>
        /// <param name="p_oQueries">Class that provides query access as well as creating a temp DB file that contains all the needed table links</param>
        /// <param name="p_strDbFile">Access database file that is the target file to contain all the table links</param>
		public void CreateTableLinksToFVSOutSummaryTables(Queries p_oQueries,string p_strDbFile)
		{
            dao_data_access oDao = new dao_data_access();
            var fvsVariantDataDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\data";
            var regex = new Regex(@"FVSOUT_([A-Z]{2})_(P\d{3})-\d{3}-\d{3}-\d{3}-\d{3}\.[(accdb)|(ACCDB)|(mdb)|(MDB)]");
            var fvsOutputDatabases = Directory.GetFiles(fvsVariantDataDir, "*.*", SearchOption.AllDirectories).Where(s => regex.IsMatch(Path.GetFileName(s))).ToArray();
            foreach (var db in fvsOutputDatabases)
            {
                if (oDao.TableExists(db, "fvs_summary"))
                {
                    var strVariant = regex.Match(db).Groups[1];
                    var strPackage = regex.Match(db).Groups[2];
                    oDao.CreateTableLink(p_strDbFile, $"fvs_summary_IN_{strVariant}_{strPackage}", db, "fvs_summary", true);
                }
            }
            oDao.m_DaoWorkspace.Close();
            oDao = null;
        }

        public static bool ValidFVSTable(string p_strTableName)
        {
            int x;
            
            for (x = 0; x <= Tables.FVS.g_strFVSOutTablesArray.Length - 1; x++)
            {
                if (Tables.FVS.g_strFVSOutTablesArray[x].Trim().ToUpper() ==
                    p_strTableName.Trim().ToUpper())
                {
                    return true;
                }
            }
            return false;
        }
        public void CheckTreeListDbExist()
        {
            // Check for SQLite elements
            DataMgr dataMgr = new DataMgr();
            string strTreeDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSTreeListDbFile;
            if (!File.Exists(strTreeDb))
            {
                dataMgr.CreateDbFile(strTreeDb);
            }
            string strConn = dataMgr.GetConnectionString(strTreeDb);
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                conn.Open();
                if (!dataMgr.TableExist(conn, Tables.FVS.DefaultFVSCutTreeTableName))
                {
                    frmMain.g_oTables.m_oFvs.CreateFVSOutTreeTable(dataMgr, conn, Tables.FVS.DefaultFVSCutTreeTableName);
                }
                if (!dataMgr.TableExist(conn, Tables.FVS.DefaultFVSResidTreeTableName))
                {
                    frmMain.g_oTables.m_oFvs.CreateFVSOutTreeTable(dataMgr, conn, Tables.FVS.DefaultFVSResidTreeTableName);
                }
            }
        }
        /// <summary>
        /// Load the Rx Harvest Methods into the processor scenario dropdown combo boxes for 
        /// low slope and steep slope
        /// </summary>
        /// <param name="p_oAdo"></param>
        /// <param name="p_oConn"></param>
        /// <param name="p_oQueries"></param>
        /// <param name="p_cmbHarvestMethod"></param>
        /// <param name="p_cmbHarvestMethodSteepSlope"></param>
		public void LoadRxHarvestMethods(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,Queries p_oQueries,System.Windows.Forms.ComboBox p_cmbHarvestMethod,System.Windows.Forms.ComboBox p_cmbHarvestMethodSteepSlope)
		{
			p_cmbHarvestMethod.Items.Clear();
			p_cmbHarvestMethodSteepSlope.Items.Clear();
			p_oAdo.m_strSQL = Queries.GenericSelectSQL(p_oQueries.m_oReference.m_strRefHarvestMethodTable,"steep_yn,method,description","steep_yn IN ('Y','N')");
			p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection,p_oAdo.m_strSQL);

			if (p_oAdo.m_intError==0)
			{
				if (p_oAdo.m_OleDbDataReader.HasRows)
				{
					while (p_oAdo.m_OleDbDataReader.Read())
					{
						if (p_oAdo.m_OleDbDataReader["method"] != System.DBNull.Value)
						{
							if (p_oAdo.m_OleDbDataReader["steep_yn"].ToString().Trim()=="Y")
							{
								p_cmbHarvestMethodSteepSlope.Items.Add(p_oAdo.m_OleDbDataReader["method"].ToString().Trim());
														
							}
							else
							{
								p_cmbHarvestMethod.Items.Add(p_oAdo.m_OleDbDataReader["method"].ToString().Trim());
								
							}
							
						}
					}
				}
				p_oAdo.m_OleDbDataReader.Close();
			}
		}

        /// <summary>
        /// Load the Rx Harvest Methods into the processor scenario dropdown combo boxes for 
        /// low slope and steep slope
        /// </summary>
        /// <param name="p_oAdo"></param>
        /// <param name="p_oConn"></param>
        /// <param name="p_oQueries"></param>
        /// <param name="p_cmbHarvestMethod"></param>
        /// <param name="p_cmbHarvestMethodSteepSlope"></param>
        /// @ToDo: This needs to be tested; Harvest methods still in MS Access when written
        public void LoadRxHarvestMethodsSqlite(string p_strDbFile, Queries p_oQueries, ComboBox p_cmbHarvestMethod, ComboBox p_cmbHarvestMethodSteepSlope)
        {
            p_cmbHarvestMethod.Items.Clear();
            p_cmbHarvestMethodSteepSlope.Items.Clear();
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(p_strDbFile);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                dataMgr.m_strSQL = Queries.GenericSelectSQL(p_oQueries.m_oReference.m_strRefHarvestMethodTable, "steep_yn,method,description", "steep_yn IN ('Y','N')");
                dataMgr.SqlQueryReader(oConn, dataMgr.m_strSQL);
                if (dataMgr.m_intError == 0)
                {
                    if (dataMgr.m_DataReader.HasRows)
                    {
                        while (dataMgr.m_DataReader.Read())
                        {
                            if (dataMgr.m_DataReader["method"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["steep_yn"].ToString().Trim() == "Y")
                                {
                                    p_cmbHarvestMethodSteepSlope.Items.Add(dataMgr.m_DataReader["method"].ToString().Trim());

                                }
                                else
                                {
                                    p_cmbHarvestMethod.Items.Add(dataMgr.m_DataReader["method"].ToString().Trim());

                                }
                            }
                        }
                    }
                    dataMgr.m_DataReader.Close();
                }
            }
        }

        public void LoadFVSOutputPrePostRxCycleSeqNum(ado_data_access p_oAdo,
                                        System.Data.OleDb.OleDbConnection p_oConn, 
                                        FVSPrePostSeqNumItem_Collection p_oCollection)
        {
            FVSPrePostSeqNumItem oItem = null;
            int x;

            p_oCollection.Clear();

            p_oAdo.m_strSQL = "SELECT * FROM " + Tables.FVS.DefaultFVSPrePostSeqNumTable + " ORDER BY PREPOST_SEQNUM_ID";
            p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    oItem = new FVSPrePostSeqNumItem();
                    oItem.PrePostSeqNumId = Convert.ToInt32(p_oAdo.m_OleDbDataReader["PREPOST_SEQNUM_ID"]);
                    oItem.TableName = Convert.ToString(p_oAdo.m_OleDbDataReader["TableName"]).Trim();
                    int intAssignedCount = 0;
                    //
                    //PRE
                    //
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE1_PRE_SEQNUM"] != DBNull.Value)
                    {
                        oItem.RxCycle1PreSeqNum = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE1_PRE_SEQNUM"]).Trim();
                        intAssignedCount++;
                    }
                    else
                    {
                        oItem.RxCycle1PreSeqNum = "Not Used";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE2_PRE_SEQNUM"] != DBNull.Value)
                    {
                        oItem.RxCycle2PreSeqNum = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE2_PRE_SEQNUM"]).Trim();
                        intAssignedCount++;
                    }
                    else
                    {
                        oItem.RxCycle2PreSeqNum = "Not Used";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE3_PRE_SEQNUM"] != DBNull.Value)
                    {
                        oItem.RxCycle3PreSeqNum = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE3_PRE_SEQNUM"]).Trim();
                        intAssignedCount++;
                    }
                    else
                    {
                        oItem.RxCycle3PreSeqNum = "Not Used";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE4_PRE_SEQNUM"] != DBNull.Value)
                    {
                        oItem.RxCycle4PreSeqNum = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE4_PRE_SEQNUM"]).Trim();
                        intAssignedCount++;
                    }
                    else
                    {
                        oItem.RxCycle4PreSeqNum = "Not Used";
                    }
                    //
                    //POST
                    //
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE1_POST_SEQNUM"] != DBNull.Value)
                    {
                        oItem.RxCycle1PostSeqNum = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE1_POST_SEQNUM"]).Trim();
                        intAssignedCount++;
                    }
                    else
                    {
                        oItem.RxCycle1PostSeqNum = "Not Used";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE2_POST_SEQNUM"] != DBNull.Value)
                    {
                        oItem.RxCycle2PostSeqNum = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE2_POST_SEQNUM"]).Trim();
                        intAssignedCount++;
                    }
                    else
                    {
                        oItem.RxCycle2PostSeqNum = "Not Used";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE3_POST_SEQNUM"] != DBNull.Value)
                    {
                        oItem.RxCycle3PostSeqNum = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE3_POST_SEQNUM"]).Trim();
                        intAssignedCount++;
                    }
                    else
                    {
                        oItem.RxCycle3PostSeqNum = "Not Used";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE4_POST_SEQNUM"] != DBNull.Value)
                    {
                        oItem.RxCycle4PostSeqNum = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE4_POST_SEQNUM"]).Trim();
                        intAssignedCount++;
                    }
                    else
                    {
                        oItem.RxCycle4PostSeqNum = "Not Used";
                    }
                    //
                    //BASEYEAR
                    //
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE1_PRE_BASEYR_YN"] != DBNull.Value)
                    {
                        oItem.RxCycle1PreSeqNumBaseYearYN = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE1_PRE_BASEYR_YN"]).Trim();
                    }
                    else
                    {
                        oItem.RxCycle1PreSeqNumBaseYearYN = "N";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE2_PRE_BASEYR_YN"] != DBNull.Value)
                    {
                        oItem.RxCycle2PreSeqNumBaseYearYN = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE2_PRE_BASEYR_YN"]).Trim();
                    }
                    else
                    {
                        oItem.RxCycle2PreSeqNumBaseYearYN = "N";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE3_PRE_BASEYR_YN"] != DBNull.Value)
                    {
                        oItem.RxCycle3PreSeqNumBaseYearYN = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE3_PRE_BASEYR_YN"]).Trim();
                    }
                    else
                    {
                        oItem.RxCycle3PreSeqNumBaseYearYN = "N";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE4_PRE_BASEYR_YN"] != DBNull.Value)
                    {
                        oItem.RxCycle4PreSeqNumBaseYearYN = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE4_PRE_BASEYR_YN"]).Trim();
                    }
                    else
                    {
                        oItem.RxCycle4PreSeqNumBaseYearYN = "N";
                    }
                    //
                    //STRCLASS BEFORE REMOVAL
                    //
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE1_PRE_BEFORECUT_YN"] != DBNull.Value)
                    {
                        oItem.RxCycle1PreStrClassBeforeTreeRemovalYN = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE1_PRE_BEFORECUT_YN"]).Trim();
                    }
                    else
                    {
                        oItem.RxCycle1PreStrClassBeforeTreeRemovalYN = "Y";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE2_PRE_BEFORECUT_YN"] != DBNull.Value)
                    {
                        oItem.RxCycle2PreStrClassBeforeTreeRemovalYN = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE2_PRE_BEFORECUT_YN"]).Trim();
                    }
                    else
                    {
                        oItem.RxCycle2PreStrClassBeforeTreeRemovalYN = "Y";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE3_PRE_BEFORECUT_YN"] != DBNull.Value)
                    {
                        oItem.RxCycle3PreStrClassBeforeTreeRemovalYN = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE3_PRE_BEFORECUT_YN"]).Trim();
                    }
                    else
                    {
                        oItem.RxCycle3PreStrClassBeforeTreeRemovalYN = "Y";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE4_PRE_BEFORECUT_YN"] != DBNull.Value)
                    {
                        oItem.RxCycle4PreStrClassBeforeTreeRemovalYN = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE4_PRE_BEFORECUT_YN"]).Trim();
                    }
                    else
                    {
                        oItem.RxCycle4PreStrClassBeforeTreeRemovalYN = "Y";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE1_POST_BEFORECUT_YN"] != DBNull.Value)
                    {
                        oItem.RxCycle1PostStrClassBeforeTreeRemovalYN = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE1_POST_BEFORECUT_YN"]).Trim();
                    }
                    else
                    {
                        oItem.RxCycle1PostStrClassBeforeTreeRemovalYN = "N";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE2_POST_BEFORECUT_YN"] != DBNull.Value)
                    {
                        oItem.RxCycle2PostStrClassBeforeTreeRemovalYN = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE2_POST_BEFORECUT_YN"]).Trim();
                    }
                    else
                    {
                        oItem.RxCycle2PostStrClassBeforeTreeRemovalYN = "N";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE3_POST_BEFORECUT_YN"] != DBNull.Value)
                    {
                        oItem.RxCycle3PostStrClassBeforeTreeRemovalYN = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE3_POST_BEFORECUT_YN"]).Trim();
                    }
                    else
                    {
                        oItem.RxCycle3PostStrClassBeforeTreeRemovalYN = "N";
                    }
                    if (p_oAdo.m_OleDbDataReader["RXCYCLE4_POST_BEFORECUT_YN"] != DBNull.Value)
                    {
                        oItem.RxCycle4PostStrClassBeforeTreeRemovalYN = Convert.ToString(p_oAdo.m_OleDbDataReader["RXCYCLE4_POST_BEFORECUT_YN"]).Trim();
                    }
                    else
                    {
                        oItem.RxCycle4PostStrClassBeforeTreeRemovalYN = "N";
                    }

                    if (p_oAdo.m_OleDbDataReader["USE_SUMMARY_TABLE_SEQNUM_YN"] != DBNull.Value)
                    {
                        oItem.UseSummaryTableSeqNumYN = Convert.ToString(p_oAdo.m_OleDbDataReader["USE_SUMMARY_TABLE_SEQNUM_YN"]).Trim();
                    }
                    else
                    {
                        oItem.UseSummaryTableSeqNumYN = "Y";
                    }

                    oItem.Type = p_oAdo.m_OleDbDataReader["TYPE"].ToString().Trim();
                    // Check for rxpackages in the fvs_output_prepost_seqnum_rxpackage_assignment for assigned package for "C"
                    // tables. If missing report 0 because FVS Out won't use assignment without rxPackages
                    if (oItem.Type.Equals("C"))
                    {
                        string strRxPackageSql = $@"SELECT COUNT(*) FROM {Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable} 
                            WHERE PREPOST_SEQNUM_ID = {oItem.PrePostSeqNumId}";
                        double rxPackageCount = p_oAdo.getSingleDoubleValueFromSQLQuery(p_oConn, strRxPackageSql, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable);
                        if (rxPackageCount == 0)
                        {
                            intAssignedCount = 0;
                        }
                    }
                    oItem.AssignedCount = intAssignedCount;

                    switch (oItem.TableName.Trim().ToUpper())
                    {
                        case "FVS_CUTLIST": oItem.MultipleRecordsForASingleStandYearCombination = true; break;
                        case "FVS_ATRTLIST": oItem.MultipleRecordsForASingleStandYearCombination = true; break;
                        case "FVS_MORTALITY": oItem.MultipleRecordsForASingleStandYearCombination = true; break;
                        case "FVS_SNAG_DET": oItem.MultipleRecordsForASingleStandYearCombination = true; break;
                        case "FVS_TREELIST": oItem.MultipleRecordsForASingleStandYearCombination = true; break;
                        case "FVS_STRCLASS": oItem.MultipleRecordsForASingleStandYearCombination = true; break;
                        default: oItem.MultipleRecordsForASingleStandYearCombination = false; break;                        
                    }



                    p_oCollection.Add(oItem);


                }
                p_oAdo.m_OleDbDataReader.Close();
                //rx package assignments for custom definitions
                for (x = 0; x <= p_oCollection.Count - 1; x++)
                {
                    if (p_oCollection.Item(x).Type == "C")
                    {
                        if (p_oCollection.Item(x).m_FVSPrePostSeqNumRxPackageAssgnItem_Collection1 == null)
                            p_oCollection.Item(x).m_FVSPrePostSeqNumRxPackageAssgnItem_Collection1 = new FVSPrePostSeqNumRxPackageAssgnItem_Collection();
                        else
                            p_oCollection.Item(x).m_FVSPrePostSeqNumRxPackageAssgnItem_Collection1.Clear();

                        p_oAdo.m_strSQL = "SELECT * FROM " + Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable + " " +
                                          "WHERE PREPOST_SEQNUM_ID=" + p_oCollection.Item(x).PrePostSeqNumId;
                        p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);
                        if (p_oAdo.m_OleDbDataReader.HasRows)
                        {
                            string strList = "";
                            while (p_oAdo.m_OleDbDataReader.Read())
                            {
                                
                                FVSPrePostSeqNumRxPackageAssgnItem oPackageItem = new FVSPrePostSeqNumRxPackageAssgnItem();
                                oPackageItem.PrePostSeqNumId = p_oCollection.Item(x).PrePostSeqNumId;
                                oPackageItem.RxPackageId = p_oAdo.m_OleDbDataReader["RXPACKAGE"].ToString().Trim();
                                strList = strList + oPackageItem.RxPackageId + ",";
                                p_oCollection.Item(x).m_FVSPrePostSeqNumRxPackageAssgnItem_Collection1.Add(oPackageItem);
                                
                            }
                            if (strList.Trim().Length > 0) strList = strList.Substring(0, strList.Length - 1);
                            p_oCollection.Item(x).RxPackageList = strList;
                        }
                        p_oAdo.m_OleDbDataReader.Close();
                        p_oCollection.Item(x).ReferenceFVSPrePostSeqNumRxPackageAssgnCollection =
                            p_oCollection.Item(x).m_FVSPrePostSeqNumRxPackageAssgnItem_Collection1;
                    }
                }
            }
        }
        public void CreateFVSPrePostSeqNumTables(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, FVSPrePostSeqNumItem p_oItem, string p_strSourceTableName, string p_strSourceLinkedTableName, bool p_bAudit,bool p_bDebug,string p_strDebugFile)
        {
           


            if (p_strSourceTableName.Trim().ToUpper() == "FVS_CASES") return;
            int x;
            string[] strSQLArray = null;
            string strPrePostSeqNumMatrixTable = p_strSourceTableName + "_PREPOST_SEQNUM_MATRIX";
            string strAuditPrePostSeqNumCountsTable = "audit_" + p_strSourceTableName + "_prepost_seqnum_counts_table";
            string strAuditYearCountsTable = "audit_" + p_strSourceTableName + "_year_counts_table";

            if (p_oAdo.TableExist(p_oConn, strPrePostSeqNumMatrixTable))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE " + strPrePostSeqNumMatrixTable);

            if (p_oAdo.TableExist(p_oConn, strAuditPrePostSeqNumCountsTable))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE " + strAuditPrePostSeqNumCountsTable);

            if (p_oAdo.TableExist(p_oConn, strAuditYearCountsTable))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE " + strAuditYearCountsTable);

            if (p_oAdo.TableExist(p_oConn, "temp_rowcount"))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE temp_rowcount");

           

            //
            //DEFAULT CONFIGURATIONS
            //
            if (p_strSourceTableName.Trim().ToUpper() == "FVS_SUMMARY" ||
                p_strSourceTableName.Trim().ToUpper() == "FVS_CUTLIST" ||
                p_strSourceTableName.Trim().ToUpper().IndexOf("FVS_POTFIRE",0) >=0)
            {
                //create the audit SeqNum table structure
                frmMain.g_oTables.m_oFvs.CreateFVSOutputPrePostSeqNumAuditGenericTable(p_oAdo, p_oConn, strPrePostSeqNumMatrixTable);
                if (p_strSourceTableName.Trim().ToUpper() == "FVS_SUMMARY" ||
                    p_strSourceTableName.Trim().ToUpper().IndexOf("FVS_POTFIRE", 0) >= 0)
                {
                    if (p_oItem.UseSummaryTableSeqNumYN == "Y")
                    {
                        //STANDID + YEAR = ONE RECORD
                        p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditPrePostGenericSQL("", "FVS_SUMMARY", false);
                    }
                    else
                    {
                        //STANDID + YEAR = ONE RECORD
                        p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditPrePostGenericSQL("", p_strSourceTableName, false);
                    }
                }
                else
                {
                    //STANDID + YEAR = MULTIPLE RECORDS
                    p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditPrePostGenericSQL("", "FVS_SUMMARY", false);
                }
                p_oAdo.m_strSQL = "INSERT INTO " + strPrePostSeqNumMatrixTable + " " +
                                  p_oAdo.m_strSQL;

                if (p_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);

                p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditUpdatePrePostGenericSQL(
                    p_oItem, strPrePostSeqNumMatrixTable);

                if (p_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);

                if (p_bAudit)
                {
                    if (p_strSourceTableName == "FVS_SUMMARY")
                    {
                        p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPrePostSeqNumCount
                       (p_oItem, strAuditPrePostSeqNumCountsTable, strPrePostSeqNumMatrixTable);

                        if (p_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                        p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);

                        p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_PrePostGenericSQL(
                           strAuditYearCountsTable, "FVS_SUMMARY", false);

                        if (p_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                        p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);

                    }
                    else
                    {

                        if (p_oItem.UseSummaryTableSeqNumYN == "N")
                        {
                            p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPrePostSeqNumCount
                              (p_oItem, strAuditPrePostSeqNumCountsTable, strPrePostSeqNumMatrixTable);

                            if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                            p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
                        }

                    }
                }
                if (p_strSourceTableName == "FVS_SUMMARY") return;
            }
            if (p_bAudit)
            {
                //
                //stand + year rowcounts for all tables as compared with stand + year in the summary table
                //
                p_oAdo.AddColumn(p_oAdo.m_OleDbConnection, "audit_fvs_summary_year_counts_table", p_strSourceTableName, "INTEGER", "");
                string[] strSQL = Queries.FVS.FVSOutputTable_AuditFVSSummaryTableRowCountsSQL(
                    "temp_rowcount",
                    "audit_FVS_SUMMARY_year_counts_table",
                    p_strSourceTableName,
                    p_strSourceTableName);

                for (x = 0; x <= strSQL.Length - 1; x++)
                {
                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                    p_oAdo.SqlNonQuery(p_oConn, strSQL[x]);
                }
            }

            if (p_strSourceTableName.Trim().ToUpper() == "FVS_TREELIST" ||
                p_strSourceTableName.Trim().ToUpper() == "FVS_CUTLIST" ||
                p_strSourceTableName.Trim().ToUpper().IndexOf("FVS_POTFIRE", 0) >= 0) return;

            //check if this table uses a default configuration of a different table
            if (p_oItem.TableName.Trim().ToUpper() !=
                p_strSourceTableName.Trim().ToUpper() &&
                p_oItem.Type == "D") return;

            //
            //CUSTOM CONFIGURATIONS
            //
            if (p_oItem.TableName.Trim().ToUpper() == "FVS_STRCLASS")
            {
                frmMain.g_oTables.m_oFvs.CreateFVSOutputPrePostSeqNumAuditStrClassTable(p_oAdo, p_oConn, strPrePostSeqNumMatrixTable);
            }
            else
            {
                frmMain.g_oTables.m_oFvs.CreateFVSOutputPrePostSeqNumAuditGenericTable(p_oAdo, p_oConn, strPrePostSeqNumMatrixTable);
            }

            if (p_oItem.TableName.Trim().ToUpper() == "FVS_STRCLASS")
            {
                if (p_oItem.UseSummaryTableSeqNumYN == "N")
                {
                    p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditPrePostStrClassSQL("", p_strSourceTableName, false);
                }
                else
                {
                    strSQLArray = Queries.FVS.FVSOutputTable_AuditPrePostFvsStrClassUsingFVSSummarySQL("", "FVS_SUMMARY", false);
                }
            }
            else
            {
                if (p_oItem.UseSummaryTableSeqNumYN == "N")
                    //STANDID + YEAR = ONE RECORD
                    p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditPrePostGenericSQL("", p_strSourceTableName, false);
                else
                    //STANDID + YEAR = MULTIPLE RECORDS
                    p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditPrePostGenericSQL("", "FVS_SUMMARY", false);

            }

            if (p_oItem.TableName.Trim().ToUpper() == "FVS_STRCLASS")
            {
                if (p_oItem.UseSummaryTableSeqNumYN == "Y")
                {
                    for (x=0;x<=strSQLArray.Length-1;x++)
                    {
                        p_oAdo.m_strSQL = "INSERT INTO " + strPrePostSeqNumMatrixTable + " " +
                                             strSQLArray[x];
                        if (p_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                        p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
                    }
                }
                else
                {
                    p_oAdo.m_strSQL = "INSERT INTO " + strPrePostSeqNumMatrixTable + " " +
                                       p_oAdo.m_strSQL;
                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                    p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
                }
                p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditUpdatePrePostStrClassSQL(
                   p_oItem, strPrePostSeqNumMatrixTable);

                if (p_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);

            }
            else
            {
                p_oAdo.m_strSQL = "INSERT INTO " + strPrePostSeqNumMatrixTable + " " +
                                        p_oAdo.m_strSQL;

                if (p_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);


                p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditUpdatePrePostGenericSQL(
                    p_oItem, strPrePostSeqNumMatrixTable);

                if (p_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
            }



            if (p_oItem.UseSummaryTableSeqNumYN == "N")
            {
                p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPrePostSeqNumCount
                  (p_oItem, strAuditPrePostSeqNumCountsTable, strPrePostSeqNumMatrixTable);

                if (p_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
            }




        }
        /// <summary>
		/// create links to the fvs out pre and post variable tables
		/// </summary>
		public void CreateTableLinksToFVSPrePostTables(string p_strDestinationDbFile)
		{
           
            string strFVSOutPrePostPathAndDbFile;
            string strFVSWeightedPathAndDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile;

            int x;
            dao_data_access oDao = new dao_data_access();
            if (!System.IO.File.Exists(p_strDestinationDbFile)) oDao.CreateMDB(p_strDestinationDbFile);
            for (x = 0; x <= Tables.FVS.g_strFVSOutTablesArray.Length - 1; x++)
            {
                strFVSOutPrePostPathAndDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\db\\PREPOST_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim() + ".ACCDB";

                if (System.IO.File.Exists(strFVSOutPrePostPathAndDbFile))
                {
                    if (oDao.TableExists(strFVSOutPrePostPathAndDbFile, "PRE_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim()) &&
                        oDao.TableExists(strFVSOutPrePostPathAndDbFile, "POST_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim()))
                    {
                        oDao.CreateTableLink(
                            p_strDestinationDbFile,
                            "PRE_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim(),
                            strFVSOutPrePostPathAndDbFile,
                            "PRE_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim(),true);
                        if (oDao.m_intError != 0)
                        {
                            m_strError = "!!Error Creating FVS PrePost Table Link!!!";
                            this.m_intError = oDao.m_intError;
                            break;
                        }
                        oDao.CreateTableLink(
                            p_strDestinationDbFile,
                            "POST_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim(),
                            strFVSOutPrePostPathAndDbFile,
                            "POST_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim(),true);
                        if (oDao.m_intError != 0)
                        {
                             m_strError = "!!Error Creating FVS PrePost Table Link!!!";
                            this.m_intError = oDao.m_intError;
                            break;
                        }

                        // Check for weighted table if pre/post exists
                        if (System.IO.File.Exists(strFVSWeightedPathAndDbFile))
                        {
                            if (oDao.TableExists(strFVSWeightedPathAndDbFile, "PRE_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim() + "_WEIGHTED") &&
                                oDao.TableExists(strFVSWeightedPathAndDbFile, "POST_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim() + "_WEIGHTED"))
                            {
                                oDao.CreateTableLink(
                                    p_strDestinationDbFile,
                                    "PRE_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim() + "_WEIGHTED",
                                    strFVSWeightedPathAndDbFile,
                                    "PRE_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim() + "_WEIGHTED", true);
                                if (oDao.m_intError != 0)
                                {
                                    m_strError = "!!Error Creating FVS PrePost Table Link!!!";
                                    this.m_intError = oDao.m_intError;
                                    break;
                                }
                                oDao.CreateTableLink(
                                    p_strDestinationDbFile,
                                    "POST_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim() + "_WEIGHTED",
                                    strFVSWeightedPathAndDbFile,
                                    "POST_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim() + "_WEIGHTED", true);
                                if (oDao.m_intError != 0)
                                {
                                    m_strError = "!!Error Creating FVS PrePost Table Link!!!";
                                    this.m_intError = oDao.m_intError;
                                    break;
                                }
                            }
                        }
                    }
                }
            }



            oDao.m_DaoWorkspace.Close();
            oDao = null;
			
			
		
        }

        public void DeleteTableLinksToFVSPrePostTables(string p_strDestinationDbFile)
        {

            int x;
            dao_data_access oDao = new dao_data_access();
            string[] arrTableNames = null;
            int y = oDao.getTableNames(p_strDestinationDbFile, ref arrTableNames, true);
            for (x = 0; x <= y - 1; x++)
            {
                if (oDao.TableType(p_strDestinationDbFile, arrTableNames[x]).Equals("L"))
                {
                    oDao.DeleteTableFromMDB(p_strDestinationDbFile, arrTableNames[x]);
                }
                    
            }

            oDao.m_DaoWorkspace.Close();
            oDao = null;

        }

        public string TreatmentProperties(FIA_Biosum_Manager.RxItem_Collection p_oColl)
        {
            string strLine = "";
            int x,y;

            strLine = "Project: " + frmMain.g_oFrmMain.frmProject.uc_project1.txtProjectId.Text + "\r\n";
            strLine = strLine + "-------------------------------------------------\r\n\r\n";

            if (p_oColl.Count == 0)
            {
                strLine = strLine + "No Treatments Defined \r\n\r\n";
            }

            for (x = 0; x <= p_oColl.Count - 1; x++)
            {
                strLine = strLine + "Treatment " + p_oColl.Item(x).RxId + "\r\n";
                strLine = strLine + "---------------------------------------------\r\n";
                strLine = strLine + "Category: " + p_oColl.Item(x).Category + "\r\n";
                strLine = strLine + "Subcategory: " + p_oColl.Item(x).SubCategory + "\r\n";
                strLine = strLine + "Description: " + p_oColl.Item(x).Description + "\r\n";
                strLine = strLine + "Harvest Method Low Slope: " + p_oColl.Item(x).HarvestMethodLowSlope + "\r\n";
                strLine = strLine + "Steep Slope Harvest Method: " + p_oColl.Item(x).HarvestMethodSteepSlope + "\r\n";
                strLine = strLine + "Package Member: " + p_oColl.Item(x).RxPackageMemberList + "\r\n";
                strLine = strLine + "Associated Harvest Cost Columns: \r\n";
                if (p_oColl.Item(x).ReferenceHarvestCostColumnCollection == null ||
                    p_oColl.Item(x).ReferenceHarvestCostColumnCollection.Count == 0)
                {
                    strLine = strLine + "None Defined \r\n";
                }
                else
                {
                    for (y = 0; y <= p_oColl.Item(x).ReferenceHarvestCostColumnCollection.Count - 1; y++)
                    {

                        if (p_oColl.Item(x).RxId ==
                            p_oColl.Item(x).ReferenceHarvestCostColumnCollection.Item(y).RxId)
                        {
                            strLine = strLine + "Cost Component: " + p_oColl.Item(x).ReferenceHarvestCostColumnCollection.Item(y).HarvestCostColumn + "\r\n";
                            if (p_oColl.Item(x).ReferenceHarvestCostColumnCollection.Item(y).Description.Trim().Length != 0)
                            {
                                strLine = strLine + "Description: " + p_oColl.Item(x).ReferenceHarvestCostColumnCollection.Item(y).Description;
                            }
                            else
                            {
                                strLine = strLine + "Description: ";
                            }
                            strLine = strLine + "\r\n";
                        }

                    }
                }

                strLine = strLine + "\r\n\r\n";
            }

            strLine = strLine + "\r\n\r\nEOF";
            return strLine;
        }

        public string PackageProperties(FIA_Biosum_Manager.RxPackageItem_Collection p_oRxPkgColl,
                                        FIA_Biosum_Manager.RxItem_Collection p_oRxColl)
        {
            string strLine = "";
            int x, y, z;

            strLine = "Project: " + frmMain.g_oFrmMain.frmProject.uc_project1.txtProjectId.Text + "\r\n";
            strLine = strLine + "-------------------------------------------------\r\n\r\n";

            if (p_oRxPkgColl.Count == 0)
            {
                strLine = strLine + "No Packages Defined \r\n\r\n";
            }

            for (x = 0; x <= p_oRxPkgColl.Count - 1; x++)
            {
                strLine = strLine + "Package " + p_oRxPkgColl.Item(x).RxPackageId + "\r\n";
                strLine = strLine + "---------------------------------------------\r\n";
                strLine = strLine + "Description: " + p_oRxPkgColl.Item(x).Description + "\r\n";
                strLine = strLine + "Treatment Cycle Length: " + p_oRxPkgColl.Item(x).RxCycleLength + "\r\n";
                strLine = strLine + "KCP/KEY File: " + p_oRxPkgColl.Item(x).KcpFile.Trim() + "\r\n";
                strLine = strLine + "Treatment Schedule: \r\n";
                strLine = strLine + "Year   Rx        Harvest Method     Steep Slope      Description \r\n";
                strLine = strLine + "                                    Harvest Method\r\n";
                //year 00 row
                string strSimulationYear = "";
                if (p_oRxPkgColl.Item(x).SimulationYear1Rx.Trim().Length > 0)
                    strSimulationYear = p_oRxPkgColl.Item(x).SimulationYear1Rx;
                string strHarvestMethodLowSlope = "";
                string strHarvestMethodSteepSlope = "";
                string strDescription = "";
                if (p_oRxPkgColl.Item(x).SimulationYear1Rx.Trim().Length > 0)
                {
                    for (y = 0; y <= p_oRxColl.Count - 1; y++)
                    {
                        if (p_oRxColl.Item(y).RxId.Trim() ==
                            p_oRxPkgColl.Item(x).SimulationYear1Rx.Trim())
                        {
                            if (p_oRxColl.Item(y).HarvestMethodLowSlope.Trim().Length > 0)
                            {
                                strHarvestMethodLowSlope = p_oRxColl.Item(y).HarvestMethodLowSlope;
                            }
                            if (p_oRxColl.Item(y).HarvestMethodSteepSlope.Trim().Length > 0)
                            {
                                strHarvestMethodSteepSlope = p_oRxColl.Item(y).HarvestMethodSteepSlope;
                            }
                            if (p_oRxColl.Item(y).Description.Trim().Length > 0)
                            {
                                strDescription = "   " + p_oRxColl.Item(y).Description;
                            }
                            break;
                        }
                    }
                }

                strLine = strLine + String.Format("{0,2}{1,6}{2,22}{3,19}{4,0}",
                    " 00",
                    strSimulationYear,
                    strHarvestMethodLowSlope,
                    strHarvestMethodSteepSlope,
                    strDescription) + "\r\n";

                //year 10 row
                if (p_oRxPkgColl.Item(x).SimulationYear2Rx.Trim().Length > 0)
                {
                    strSimulationYear = p_oRxPkgColl.Item(x).SimulationYear2Rx;
                }
                else
                {
                    strSimulationYear = "";
                }
                // Reset variables to blanks
                strHarvestMethodLowSlope = "";
                strHarvestMethodSteepSlope = "";
                strDescription = "";
                if (p_oRxPkgColl.Item(x).SimulationYear2Rx.Trim().Length > 0)
                {
                    for (y = 0; y <= p_oRxColl.Count - 1; y++)
                    {
                        if (p_oRxColl.Item(y).RxId.Trim() ==
                            p_oRxPkgColl.Item(x).SimulationYear2Rx.Trim())
                        {
                            if (p_oRxColl.Item(y).HarvestMethodLowSlope.Trim().Length > 0)
                            {
                                strHarvestMethodLowSlope = p_oRxColl.Item(y).HarvestMethodLowSlope;
                            }
                            if (p_oRxColl.Item(y).HarvestMethodSteepSlope.Trim().Length > 0)
                            {
                                strHarvestMethodSteepSlope = p_oRxColl.Item(y).HarvestMethodSteepSlope;
                            }
                            if (p_oRxColl.Item(y).Description.Trim().Length > 0)
                            {
                                strDescription = "   " + p_oRxColl.Item(y).Description;
                            }
                            break;
                        }
                    }
                }
                
                strLine = strLine + String.Format("{0,2}{1,6}{2,22}{3,19}{4,0}",
                    Convert.ToString(p_oRxPkgColl.Item(x).RxCycleLength * 1).PadLeft(2, '0').PadLeft(3),
                    strSimulationYear,
                    strHarvestMethodLowSlope,
                    strHarvestMethodSteepSlope,
                    strDescription) + "\r\n";

                //year 20 row
                if (p_oRxPkgColl.Item(x).SimulationYear3Rx.Trim().Length > 0)
                {
                    strSimulationYear = p_oRxPkgColl.Item(x).SimulationYear3Rx;
                }
                else
                {
                    strSimulationYear = "";
                }
                // Reset variables to blanks
                strHarvestMethodLowSlope = "";
                strHarvestMethodSteepSlope = "";
                strDescription = "";
                if (p_oRxPkgColl.Item(x).SimulationYear3Rx.Trim().Length > 0)
                {
                    for (y = 0; y <= p_oRxColl.Count - 1; y++)
                    {
                        if (p_oRxColl.Item(y).RxId.Trim() ==
                            p_oRxPkgColl.Item(x).SimulationYear3Rx.Trim())
                        {
                            if (p_oRxColl.Item(y).HarvestMethodLowSlope.Trim().Length > 0)
                            {
                                strHarvestMethodLowSlope = p_oRxColl.Item(y).HarvestMethodLowSlope;
                            }
                            if (p_oRxColl.Item(y).HarvestMethodSteepSlope.Trim().Length > 0)
                            {
                                strHarvestMethodSteepSlope = p_oRxColl.Item(y).HarvestMethodSteepSlope;
                            }
                            if (p_oRxColl.Item(y).Description.Trim().Length > 0)
                            {
                                strDescription = "   " + p_oRxColl.Item(y).Description;
                            }
                            break;
                        }
                    }
                }

                strLine = strLine + String.Format("{0,2}{1,6}{2,22}{3,19}{4,0}",
                    Convert.ToString(p_oRxPkgColl.Item(x).RxCycleLength * 2).PadLeft(2, '0').PadLeft(3),
                    strSimulationYear,
                    strHarvestMethodLowSlope,
                    strHarvestMethodSteepSlope,
                    strDescription) + "\r\n";

                //year 30 row
                if (p_oRxPkgColl.Item(x).SimulationYear4Rx.Trim().Length > 0)
                {
                    strSimulationYear = p_oRxPkgColl.Item(x).SimulationYear4Rx;
                }
                else
                {
                    strSimulationYear = "";
                }
                // Reset variables to blanks
                strHarvestMethodLowSlope = "";
                strHarvestMethodSteepSlope = "";
                strDescription = "";
                if (p_oRxPkgColl.Item(x).SimulationYear4Rx.Trim().Length > 0)
                {
                    for (y = 0; y <= p_oRxColl.Count - 1; y++)
                    {
                        if (p_oRxColl.Item(y).RxId.Trim() ==
                            p_oRxPkgColl.Item(x).SimulationYear4Rx.Trim())
                        {
                            if (p_oRxColl.Item(y).HarvestMethodLowSlope.Trim().Length > 0)
                            {
                                strHarvestMethodLowSlope = p_oRxColl.Item(y).HarvestMethodLowSlope;
                            }
                            if (p_oRxColl.Item(y).HarvestMethodSteepSlope.Trim().Length > 0)
                            {
                                strHarvestMethodSteepSlope = p_oRxColl.Item(y).HarvestMethodSteepSlope;
                            }
                            if (p_oRxColl.Item(y).Description.Trim().Length > 0)
                            {
                                strDescription = "   " + p_oRxColl.Item(y).Description;
                            }
                            break;
                        }
                    }
                }

                strLine = strLine + String.Format("{0,2}{1,6}{2,22}{3,19}{4,0}",
                    Convert.ToString(p_oRxPkgColl.Item(x).RxCycleLength * 3).PadLeft(2, '0').PadLeft(3),
                    strSimulationYear,
                    strHarvestMethodLowSlope,
                    strHarvestMethodSteepSlope,
                    strDescription) + "\r\n";

                strLine = strLine + "Associated Harvest Cost Columns: \r\n";
                //see if any harvest cost columns in the package
                bool bHarvestColumnsExist = false;
                for (y = 0; y <= p_oRxColl.Count - 1; y++)
                {
                    if (p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection != null &&
                        p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection.Count > 0)
                    {

                        if (p_oRxColl.Item(y).RxId.Trim() ==
                                        p_oRxPkgColl.Item(x).SimulationYear1Rx.Trim() ||
                                        p_oRxColl.Item(y).RxId.Trim() ==
                                        p_oRxPkgColl.Item(x).SimulationYear2Rx.Trim() ||
                                        p_oRxColl.Item(y).RxId.Trim() ==
                                        p_oRxPkgColl.Item(x).SimulationYear3Rx.Trim() ||
                                        p_oRxColl.Item(y).RxId.Trim() ==
                                        p_oRxPkgColl.Item(x).SimulationYear4Rx.Trim())
                        {
                            bHarvestColumnsExist = true; break;
                        }


                    }
                    if (bHarvestColumnsExist == true) break;
                }
                if (bHarvestColumnsExist == false)
                {
                    strLine = strLine + "None Defined \r\n";
                }
                else
                {
                    strLine = strLine + " Rx   Simulation        Harvest Cost     Description \r\n";
                    strLine = strLine + "         Cycle             Column \r\n";

                    int intCycle;
                    for (intCycle = 1; intCycle <= 4; intCycle++)
                    {
                        bHarvestColumnsExist = false;
                        if (intCycle == 1)
                        {
                            for (y = 0; y <= p_oRxColl.Count - 1; y++)
                            {
                                if (p_oRxPkgColl.Item(x).SimulationYear1Rx.Trim() ==
                                    p_oRxColl.Item(y).RxId.Trim())
                                {
                                    if (p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection != null &&
                                    p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection.Count > 0)
                                    {
                                        bHarvestColumnsExist = true;
                                    }
                                    break;
                                }
                            }
                        }
                        else if (intCycle == 2)
                        {
                            for (y = 0; y <= p_oRxColl.Count - 1; y++)
                            {
                                if (p_oRxPkgColl.Item(x).SimulationYear2Rx.Trim() ==
                                    p_oRxColl.Item(y).RxId.Trim())
                                {
                                    if (p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection != null &&
                                    p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection.Count > 0)
                                    {
                                        bHarvestColumnsExist = true;
                                    }
                                    break;
                                }
                            }
                        }
                        else if (intCycle == 3)
                        {
                            for (y = 0; y <= p_oRxColl.Count - 1; y++)
                            {
                                if (p_oRxPkgColl.Item(x).SimulationYear3Rx.Trim() ==
                                    p_oRxColl.Item(y).RxId.Trim())
                                {
                                    if (p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection != null &&
                                    p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection.Count > 0)
                                    {
                                        bHarvestColumnsExist = true;
                                    }
                                    break;
                                }
                            }
                        }
                        else
                        {
                            for (y = 0; y <= p_oRxColl.Count - 1; y++)
                            {
                                if (p_oRxPkgColl.Item(x).SimulationYear4Rx.Trim() ==
                                    p_oRxColl.Item(y).RxId.Trim())
                                {
                                    if (p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection != null &&
                                        p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection.Count > 0)
                                    {
                                        bHarvestColumnsExist = true;
                                    }
                                    break;
                                }
                            }
                        }
                        if (bHarvestColumnsExist)
                        {
                            System.Collections.Generic.IList<string> lstHarvestCostColumns =
                                new System.Collections.Generic.List<string>();
                            System.Collections.Generic.IList<string> lstHarvestCostColumnDesc =
                                new System.Collections.Generic.List<string>();

                            //rx
                            string strRx = "";
                            if (p_oRxColl.Item(y).RxId.Trim().Trim().Length != 0)
                            {
                                strRx = p_oRxColl.Item(y).RxId.Trim();
                            }
                            //SimYear
                            string strCycle = "";
                            if (intCycle.ToString().Trim().Length > 0)
                            {
                                strCycle = intCycle.ToString().Trim();
                            }
                            for (z = 0; z <= p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection.Count - 1; z++)
                            {
                                if (p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection.Item(z).HarvestCostColumn.Trim().Length > 0)
                                {

                                    lstHarvestCostColumns.Add(p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection.Item(z).HarvestCostColumn.Trim());
                                    if (p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection.Item(z).Description.Trim().Length > 0)
                                    {
                                        lstHarvestCostColumnDesc.Add(p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection.Item(z).HarvestCostColumn.Trim() + "=" +
                                                                     p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection.Item(z).Description.Trim());
                      
                                    }
                                    else
                                    {
                                        lstHarvestCostColumnDesc.Add(p_oRxColl.Item(y).ReferenceHarvestCostColumnCollection.Item(z).HarvestCostColumn.Trim() + "=No Description Given");  
                                    }
                                }
                            }

                            //Harvest Cost Rows
                            string strDescrPadding = Convert.ToString(lstHarvestCostColumnDesc[0].Length + 7);
                            strLine = strLine + String.Format("{0,3}{1,9}{2,22}{3," + strDescrPadding + "}",
                                      strRx,
                                      strCycle,
                                      lstHarvestCostColumns[0],
                                      lstHarvestCostColumnDesc[0]) + "\r\n";
                            for (z = 0; z < lstHarvestCostColumns.Count; z++)
                            {
                                // We don't want to do anything unless z > 0 because we already handled the 0 row
                                if (z > 0)
                                {
                                    strDescrPadding = Convert.ToString(lstHarvestCostColumnDesc[z].Length + 7);
                                    strLine = strLine + String.Format("{0,34}{1," + strDescrPadding + "}",
                                        lstHarvestCostColumns[z],
                                        lstHarvestCostColumnDesc[z]) + "\r\n";
                                }
                                
                            }

                        }
                    }	
                }


                strLine = strLine + "\r\n\r\n";
            }

            strLine = strLine + "\r\n\r\nEOF";
            return strLine;
        }

        public System.Collections.Generic.IList<string> GetFvsOutTableNames(DataMgr oDataMgr, string strConnection)
        {
            string[] arrTableNames = null;
            System.Collections.Generic.List<string> lstTableNames = new System.Collections.Generic.List<string>();
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConnection))
            {
                oConn.Open();
                arrTableNames = oDataMgr.getTableNames(oConn);
            }
            foreach (var tName in arrTableNames)
            {
                if (ValidFVSTable(tName))
                {
                    lstTableNames.Add(tName.ToUpper());
                }
            }
            lstTableNames.Sort();
            return lstTableNames;
        }

        public void CreateFVSPrePostSeqNumTables(string strTargetDb, FVSPrePostSeqNumItem p_oItem, string p_strSourceTableName, string p_strSourceLinkedTableName, bool p_bAudit, bool p_bDebug, string p_strDebugFile)
        {
            DataMgr dataMgr = new DataMgr();
            if (p_strSourceTableName.Trim().ToUpper() == "FVS_CASES") return;
            int x;
            string[] strSQLArray = null;
            string strPrePostSeqNumMatrixTable = p_strSourceTableName + "_PREPOST_SEQNUM_MATRIX";
            string strAuditPrePostSeqNumCountsTable = "audit_" + p_strSourceTableName + "_prepost_seqnum_counts_table";
            string strAuditYearCountsTable = "audit_" + p_strSourceTableName + "_year_counts_table";
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dataMgr.GetConnectionString(strTargetDb)))
            {
                conn.Open();
                if (dataMgr.TableExist(conn, strPrePostSeqNumMatrixTable))
                    dataMgr.SqlNonQuery(conn, "DROP TABLE " + strPrePostSeqNumMatrixTable);

                if (dataMgr.TableExist(conn, strAuditPrePostSeqNumCountsTable))
                    dataMgr.SqlNonQuery(conn, "DROP TABLE " + strAuditPrePostSeqNumCountsTable);

                if (dataMgr.TableExist(conn, strAuditYearCountsTable))
                    dataMgr.SqlNonQuery(conn, "DROP TABLE " + strAuditYearCountsTable);

                if (dataMgr.TableExist(conn, "temp_rowcount"))
                    dataMgr.SqlNonQuery(conn, "DROP TABLE temp_rowcount");

                //
                //DEFAULT CONFIGURATIONS
                //
                if (p_strSourceTableName.Trim().ToUpper() == "FVS_SUMMARY" ||
                    p_strSourceTableName.Trim().ToUpper() == "FVS_CUTLIST" ||
                    p_strSourceTableName.Trim().ToUpper().IndexOf("FVS_POTFIRE", 0) >= 0)
                {
                    // Attach the FVSOut.db
                    dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                                        Tables.FVS.DefaultFVSOutDbFile + "' AS FVS";
                    dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);

                    //create the audit SeqNum table structure
                    frmMain.g_oTables.m_oFvs.CreateSQLiteFVSOutputPrePostSeqNumAuditGenericTable(dataMgr, conn, strPrePostSeqNumMatrixTable);
                    if (p_strSourceTableName.Trim().ToUpper() == "FVS_SUMMARY" ||
                    p_strSourceTableName.Trim().ToUpper().IndexOf("FVS_POTFIRE", 0) >= 0)
                    {
                        if (p_oItem.UseSummaryTableSeqNumYN == "Y")
                        {
                            //STANDID + YEAR = ONE RECORD
                            dataMgr.m_strSQL = Queries.FVS.SqliteFVSOutputTable_AuditPrePostGenericSQL("", "FVS_SUMMARY", false);
                        }
                        else
                        {
                            //STANDID + YEAR = ONE RECORD
                            dataMgr.m_strSQL = Queries.FVS.SqliteFVSOutputTable_AuditPrePostGenericSQL("", p_strSourceTableName, false);
                        }
                    }
                    else
                    {
                        //STANDID + YEAR = MULTIPLE RECORDS
                        dataMgr.m_strSQL = Queries.FVS.SqliteFVSOutputTable_AuditPrePostGenericSQL("", "FVS_SUMMARY", false);
                    }
                    dataMgr.m_strSQL = "INSERT INTO " + strPrePostSeqNumMatrixTable + " " +
                                      dataMgr.m_strSQL;

                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, dataMgr.m_strSQL + "\r\n\r\n");
                    dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);

                    //p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditUpdatePrePostGenericSQL(
                    //    p_oItem, strPrePostSeqNumMatrixTable);

                    //if (p_bDebug && frmMain.g_intDebugLevel > 2)
                    //    frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                    //p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);

                    //if (p_bAudit)
                    //{
                    //    if (p_strSourceTableName == "FVS_SUMMARY")
                    //    {
                    //        p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPrePostSeqNumCount
                    //       (p_oItem, strAuditPrePostSeqNumCountsTable, strPrePostSeqNumMatrixTable);

                    //        if (p_bDebug && frmMain.g_intDebugLevel > 2)
                    //            frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                    //        p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);

                    //        p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_PrePostGenericSQL(
                    //           strAuditYearCountsTable, "FVS_SUMMARY", false);

                    //        if (p_bDebug && frmMain.g_intDebugLevel > 2)
                    //            frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                    //        p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);

                    //    }
                    //    else
                    //    {

                    //        if (p_oItem.UseSummaryTableSeqNumYN == "N")
                    //        {
                    //            p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPrePostSeqNumCount
                    //              (p_oItem, strAuditPrePostSeqNumCountsTable, strPrePostSeqNumMatrixTable);

                    //            if (p_bDebug && frmMain.g_intDebugLevel > 2)
                    //                frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
                    //            p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
                    //        }

                    //    }
                    //}
                    if (p_strSourceTableName == "FVS_SUMMARY") return;
                }
            }


            //if (p_bAudit)
            //{
            //    //
            //    //stand + year rowcounts for all tables as compared with stand + year in the summary table
            //    //
            //    p_oAdo.AddColumn(p_oAdo.m_OleDbConnection, "audit_fvs_summary_year_counts_table", p_strSourceTableName, "INTEGER", "");
            //    string[] strSQL = Queries.FVS.FVSOutputTable_AuditFVSSummaryTableRowCountsSQL(
            //        "temp_rowcount",
            //        "audit_FVS_SUMMARY_year_counts_table",
            //        p_strSourceTableName,
            //        p_strSourceTableName);

            //    for (x = 0; x <= strSQL.Length - 1; x++)
            //    {
            //        if (p_bDebug && frmMain.g_intDebugLevel > 2)
            //            frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
            //        p_oAdo.SqlNonQuery(p_oConn, strSQL[x]);
            //    }
            //}

            //if (p_strSourceTableName.Trim().ToUpper() == "FVS_TREELIST" ||
            //    p_strSourceTableName.Trim().ToUpper() == "FVS_CUTLIST" ||
            //    p_strSourceTableName.Trim().ToUpper().IndexOf("FVS_POTFIRE", 0) >= 0) return;

            ////check if this table uses a default configuration of a different table
            //if (p_oItem.TableName.Trim().ToUpper() !=
            //    p_strSourceTableName.Trim().ToUpper() &&
            //    p_oItem.Type == "D") return;

            ////
            ////CUSTOM CONFIGURATIONS
            ////
            //if (p_oItem.TableName.Trim().ToUpper() == "FVS_STRCLASS")
            //{
            //    frmMain.g_oTables.m_oFvs.CreateFVSOutputPrePostSeqNumAuditStrClassTable(p_oAdo, p_oConn, strPrePostSeqNumMatrixTable);
            //}
            //else
            //{
            //    frmMain.g_oTables.m_oFvs.CreateFVSOutputPrePostSeqNumAuditGenericTable(p_oAdo, p_oConn, strPrePostSeqNumMatrixTable);
            //}

            //if (p_oItem.TableName.Trim().ToUpper() == "FVS_STRCLASS")
            //{
            //    if (p_oItem.UseSummaryTableSeqNumYN == "N")
            //    {
            //        p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditPrePostStrClassSQL("", p_strSourceTableName, false);
            //    }
            //    else
            //    {
            //        strSQLArray = Queries.FVS.FVSOutputTable_AuditPrePostFvsStrClassUsingFVSSummarySQL("", "FVS_SUMMARY", false);
            //    }
            //}
            //else
            //{
            //    if (p_oItem.UseSummaryTableSeqNumYN == "N")
            //        //STANDID + YEAR = ONE RECORD
            //        p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditPrePostGenericSQL("", p_strSourceTableName, false);
            //    else
            //        //STANDID + YEAR = MULTIPLE RECORDS
            //        p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditPrePostGenericSQL("", "FVS_SUMMARY", false);

            //}

            //if (p_oItem.TableName.Trim().ToUpper() == "FVS_STRCLASS")
            //{
            //    if (p_oItem.UseSummaryTableSeqNumYN == "Y")
            //    {
            //        for (x = 0; x <= strSQLArray.Length - 1; x++)
            //        {
            //            p_oAdo.m_strSQL = "INSERT INTO " + strPrePostSeqNumMatrixTable + " " +
            //                                 strSQLArray[x];
            //            if (p_bDebug && frmMain.g_intDebugLevel > 2)
            //                frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
            //            p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
            //        }
            //    }
            //    else
            //    {
            //        p_oAdo.m_strSQL = "INSERT INTO " + strPrePostSeqNumMatrixTable + " " +
            //                           p_oAdo.m_strSQL;
            //        if (p_bDebug && frmMain.g_intDebugLevel > 2)
            //            frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
            //        p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
            //    }
            //    p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditUpdatePrePostStrClassSQL(
            //       p_oItem, strPrePostSeqNumMatrixTable);

            //    if (p_bDebug && frmMain.g_intDebugLevel > 2)
            //        frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
            //    p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);

            //}
            //else
            //{
            //    p_oAdo.m_strSQL = "INSERT INTO " + strPrePostSeqNumMatrixTable + " " +
            //                            p_oAdo.m_strSQL;

            //    if (p_bDebug && frmMain.g_intDebugLevel > 2)
            //        frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
            //    p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);


            //    p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditUpdatePrePostGenericSQL(
            //        p_oItem, strPrePostSeqNumMatrixTable);

            //    if (p_bDebug && frmMain.g_intDebugLevel > 2)
            //        frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
            //    p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
            //}



            //if (p_oItem.UseSummaryTableSeqNumYN == "N")
            //{
            //    p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPrePostSeqNumCount
            //      (p_oItem, strAuditPrePostSeqNumCountsTable, strPrePostSeqNumMatrixTable);

            //    if (p_bDebug && frmMain.g_intDebugLevel > 2)
            //        frmMain.g_oUtils.WriteText(p_strDebugFile, p_oAdo.m_strSQL + "\r\n\r\n");
            //    p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
            //}
        }

    }
	/*********************************************************************************************************
	 **RX Item                          
	 *********************************************************************************************************/
	public class RxItem
	{
		private int _intIndex;
		public RxItemFvsCommandItem_Collection m_oFvsCommandItem_Collection1=null;
		public RxItemFvsCommandItem_Collection _FvsCommandItem_Collection1=null;
		public RxItemHarvestCostColumnItem_Collection _HarvestCostColumnItem_Collection1=null;
		public RxItemHarvestCostColumnItem_Collection m_oHarvestCostColumnItem_Collection1=null;

		[CategoryAttribute("General"),ReadOnly(true),DescriptionAttribute("RX Item Index")]
		public int Index
		{
			get {return _intIndex;}
			set {_intIndex = value;}
		}
		private string _strRxId="";
		[CategoryAttribute("General"),DescriptionAttribute("RX Indentifier")]
		public string RxId
		{
			get {return _strRxId;}
			set {_strRxId=value;}
		}
		private string _strDesc="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Description
		{
			get {return _strDesc;}
			set {_strDesc=value;}

		}
		private string _strCat="";
		[CategoryAttribute("General"),DescriptionAttribute("FVS Rx Category")]
		public string Category
		{
			get {return _strCat;}
			set {_strCat=value;}

		}
		private string _strSubCat="";
		[CategoryAttribute("General"),DescriptionAttribute("FVS Rx SubCategory")]
		public string SubCategory
		{
			get {return _strSubCat;}
			set {_strSubCat=value;}

		}

		private string _strFvsCycleList="";
		/// <summary>
		/// contains a comma delimited list of fvs cycles the rx is applied to in a package
		/// </summary>
		public string FvsCycleList
		{
			get {return _strFvsCycleList;}
			set {_strFvsCycleList=value;}
		}



		private string _strRxPackageMemberList="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string RxPackageMemberList
		{
			get {return _strRxPackageMemberList;}
			set {_strRxPackageMemberList=value;}

		}

		private string _strHarvestMethodLowSlope="";
		public string HarvestMethodLowSlope
		{
			get {return _strHarvestMethodLowSlope;}
			set {_strHarvestMethodLowSlope=value;}
		}

		private string _strHarvestMethodSteepSlope="";
		public string HarvestMethodSteepSlope
		{
			get {return _strHarvestMethodSteepSlope;}
			set {_strHarvestMethodSteepSlope=value;}
		}

		//private string _strHarvestCostColumnList="";
		//public string HarvestCostColumnList
		//{
		//	get {return _strHarvestCostColumnList;}
		//	set {_strHarvestCostColumnList=value;}
		//}

		public  RxItemHarvestCostColumnItem_Collection	ReferenceHarvestCostColumnCollection
		{
			get {return _HarvestCostColumnItem_Collection1;}
			set {_HarvestCostColumnItem_Collection1=value;}
		}
		public RxItemFvsCommandItem_Collection ReferenceFvsCommandsCollection
		{
			get {return _FvsCommandItem_Collection1;}
			set {_FvsCommandItem_Collection1=value;}
		}
		bool _bAdd = false;
		public bool Add
		{
			get {return _bAdd;}
			set {_bAdd=value;}
		}


		bool _bDelete=false;
		public bool Delete
		{
			get {return _bDelete;}
			set {_bDelete=value;}
		}
			
		//private string _strFvsCmd="";
		//[CategoryAttribute("Estimation Engine And Excel"), BrowsableAttribute(false), DescriptionAttribute("FVS Command")]
		//public string FVSCommand
		//{
		//	get {return _strFVSCmd;}
		//	set {_strFVSCmd=value;}
		//}
		public void CopyProperties(FIA_Biosum_Manager.RxItem p_oRxItemSource,FIA_Biosum_Manager.RxItem  p_oRxItemDestination)
		{
			int x;

			
			p_oRxItemDestination.Category="";
			p_oRxItemDestination.Description="";
			p_oRxItemDestination.Index=-1;
			p_oRxItemDestination.RxId="";
			p_oRxItemDestination.RxPackageMemberList="";
			p_oRxItemDestination.SubCategory="";
			p_oRxItemDestination.HarvestMethodLowSlope="";
			p_oRxItemDestination.HarvestMethodSteepSlope="";
			
			if (p_oRxItemDestination.m_oFvsCommandItem_Collection1 != null)
			{
				for (x=p_oRxItemDestination.m_oFvsCommandItem_Collection1.Count-1;x>=0;x--)
				{
					p_oRxItemDestination.m_oFvsCommandItem_Collection1.Remove(x);
				}

			}
			if (p_oRxItemDestination.ReferenceFvsCommandsCollection != null)
			{
				for (x=p_oRxItemDestination.ReferenceFvsCommandsCollection.Count-1;x>=0;x--)
				{
					p_oRxItemDestination.ReferenceFvsCommandsCollection.Remove(x);
				}

			}
			if (p_oRxItemDestination.m_oHarvestCostColumnItem_Collection1 != null)
			{
				for (x=p_oRxItemDestination.m_oHarvestCostColumnItem_Collection1.Count-1;x>=0;x--)
				{
					p_oRxItemDestination.m_oHarvestCostColumnItem_Collection1.Remove(x);
				}

			}
			if (p_oRxItemDestination.ReferenceHarvestCostColumnCollection != null)
			{
				for (x=p_oRxItemDestination.ReferenceHarvestCostColumnCollection.Count-1;x>=0;x--)
				{
					p_oRxItemDestination.ReferenceHarvestCostColumnCollection.Remove(x);
				}

			}

			p_oRxItemDestination.Index  = p_oRxItemSource.Index;
			p_oRxItemDestination.Description = p_oRxItemSource.Description;
			p_oRxItemDestination.RxId = p_oRxItemSource.RxId;
			p_oRxItemDestination.RxPackageMemberList = p_oRxItemSource.RxPackageMemberList;
			p_oRxItemDestination.Category = p_oRxItemSource.Category;
			p_oRxItemDestination.SubCategory=p_oRxItemSource.SubCategory;
			p_oRxItemDestination.FvsCycleList=p_oRxItemSource.FvsCycleList;
            p_oRxItemDestination.HarvestMethodLowSlope = p_oRxItemSource.HarvestMethodLowSlope;
			p_oRxItemDestination.HarvestMethodSteepSlope=p_oRxItemSource.HarvestMethodSteepSlope;
			p_oRxItemDestination.Delete = p_oRxItemSource.Delete;
			p_oRxItemDestination.Add = p_oRxItemSource.Add;

			//remove any existing destimation fvs command collection items 
			//since we are copying all the source to the destination
			if (p_oRxItemDestination.ReferenceFvsCommandsCollection!=null)
			{
				for (x=0;x<=p_oRxItemDestination.ReferenceFvsCommandsCollection.Count-1;x++)
				{
					if (p_oRxItemDestination.ReferenceFvsCommandsCollection.Item(x).RxId==
						p_oRxItemSource.RxId)
						p_oRxItemDestination.ReferenceFvsCommandsCollection.Remove(x);
				}
			}
			//
			//FVS COMMANDS
			//
			if (p_oRxItemSource.ReferenceFvsCommandsCollection != null)
			{
				
			    p_oRxItemDestination.m_oFvsCommandItem_Collection1=new RxItemFvsCommandItem_Collection();
				for (x=0;x<=p_oRxItemSource.ReferenceFvsCommandsCollection.Count-1;x++)
				{
					if (p_oRxItemSource.RxId == p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).RxId)
					{
						FIA_Biosum_Manager.RxItemFvsCommandItem oItem = new RxItemFvsCommandItem();
						oItem.Index = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Index;
						oItem.SaveIndex =  p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).SaveIndex;
						oItem.RxId = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).RxId;
						oItem.FVSCommand = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).FVSCommand;
						oItem.FVSCommandId = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).FVSCommandId;
						oItem.Other = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Other;
						oItem.OtherDescription = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).OtherDescription;
						oItem.Parameter1=p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter1;
						oItem.Parameter1Description = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter1Description;
						oItem.Parameter2 = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter2;
						oItem.Parameter2Description = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter2Description;
						oItem.Parameter3 = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter3;
						oItem.Parameter3Description = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter3Description;
						oItem.Parameter4 = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter4;
						oItem.Parameter4Description = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter4Description;
						oItem.Parameter5 = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter5;
						oItem.Parameter5Description = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter5Description;
						oItem.Parameter6 = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter6;
						oItem.Parameter6Description = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter6Description;
						oItem.Parameter7 = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter7;
						oItem.Parameter7Description = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Parameter7Description;
						oItem.Delete =  p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Delete;
						oItem.Add = p_oRxItemSource.ReferenceFvsCommandsCollection.Item(x).Add;
						p_oRxItemDestination.m_oFvsCommandItem_Collection1.Add(oItem);

					}
					p_oRxItemDestination.ReferenceFvsCommandsCollection = p_oRxItemDestination.m_oFvsCommandItem_Collection1;
				
				}
			}
			//
			//HARVEST COST COLUMNS
			//
			if (p_oRxItemSource.ReferenceHarvestCostColumnCollection != null)
			{
				
				p_oRxItemDestination.m_oHarvestCostColumnItem_Collection1=new RxItemHarvestCostColumnItem_Collection();
				for (x=0;x<=p_oRxItemSource.ReferenceHarvestCostColumnCollection.Count-1;x++)
				{
					if (p_oRxItemSource.RxId == p_oRxItemSource.ReferenceHarvestCostColumnCollection.Item(x).RxId)
					{
						FIA_Biosum_Manager.RxItemHarvestCostColumnItem oItem = new RxItemHarvestCostColumnItem();
						oItem.Index = p_oRxItemSource.ReferenceHarvestCostColumnCollection.Item(x).Index;
						oItem.SaveIndex =  p_oRxItemSource.ReferenceHarvestCostColumnCollection.Item(x).SaveIndex;
						oItem.RxId = p_oRxItemSource.ReferenceHarvestCostColumnCollection.Item(x).RxId;
						oItem.HarvestCostColumn = p_oRxItemSource.ReferenceHarvestCostColumnCollection.Item(x).HarvestCostColumn;
						oItem.Description = p_oRxItemSource.ReferenceHarvestCostColumnCollection.Item(x).Description;
						oItem.Delete =  p_oRxItemSource.ReferenceHarvestCostColumnCollection.Item(x).Delete;
						oItem.Add = p_oRxItemSource.ReferenceHarvestCostColumnCollection.Item(x).Add;
						p_oRxItemDestination.m_oHarvestCostColumnItem_Collection1.Add(oItem);

					}
					p_oRxItemDestination.ReferenceHarvestCostColumnCollection = p_oRxItemDestination.m_oHarvestCostColumnItem_Collection1;
				
				}
			}
		}

	}
	public class RxItem_Collection : System.Collections.CollectionBase
	{
		public RxItem_Collection()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public void Add(FIA_Biosum_Manager.RxItem m_PropertiesRxItem)
		{
			// vérify if object is not already in
			if (this.List.Contains(m_PropertiesRxItem))
				throw new InvalidOperationException();
 
			// adding it
			this.List.Add(m_PropertiesRxItem);
 
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
		public FIA_Biosum_Manager.RxItem Item(int Index)
		{
			// The appropriate item is retrieved from the List object and
			// explicitly cast to the Widget type, then returned to the 
			// caller.
			return (FIA_Biosum_Manager.RxItem) List[Index];
		}

	}
	/*********************************************************************************************************
	 **RX FVS Command Item                      
	 *********************************************************************************************************/
	public class RxItemFvsCommandItem
	{
		private int _intIndex;
		
		[CategoryAttribute("General"),ReadOnly(true),DescriptionAttribute("RX Item Index")]
		public int Index
		{
			get {return _intIndex;}
			set {_intIndex = value;}
		}
		private int _intSaveIndex;
		[CategoryAttribute("General"),ReadOnly(true),DescriptionAttribute("RX Item Save Index")]
		public int SaveIndex
		{
			get {return _intSaveIndex;}
			set {_intSaveIndex = value;}
		}

		private string _strRxId="";
		[CategoryAttribute("General"),DescriptionAttribute("RX Indentifier")]
		public string RxId
		{
			get {return _strRxId;}
			set {_strRxId=value;}
		}
		private string _strFVSCmd="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string FVSCommand
		{
			get {return _strFVSCmd;}
			set {_strFVSCmd=value;}
		}
		private byte _byteFVSCmdId=0;
		public byte FVSCommandId
		{
			get {return _byteFVSCmdId;}
			set {_byteFVSCmdId=value;}
		}
		private string _strP1="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter1
		{
			get {return _strP1;}
			set {_strP1=value;}
		}
		private string _strP1Desc="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter1Description
		{
			get {return _strP1Desc;}
			set {_strP1Desc=value;}
		}
		private string _strP2="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter2
		{
			get {return _strP2;}
			set {_strP2=value;}
		}
		private string _strP2Desc="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter2Description
		{
			get {return _strP2Desc;}
			set {_strP2Desc=value;}
		}
		private string _strP3="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter3
		{
			get {return _strP3;}
			set {_strP3=value;}
		}
		private string _strP3Desc="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter3Description
		{
			get {return _strP3Desc;}
			set {_strP3Desc=value;}
		}
		private string _strP4="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter4
		{
			get {return _strP4;}
			set {_strP4=value;}
		}
		private string _strP4Desc="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter4Description
		{
			get {return _strP4Desc;}
			set {_strP4Desc=value;}
		}
		private string _strP5="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter5
		{
			get {return _strP5;}
			set {_strP5=value;}
		}
		private string _strP5Desc="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter5Description
		{
			get {return _strP5Desc;}
			set {_strP5Desc=value;}
		}
		private string _strP6="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter6
		{
			get {return _strP6;}
			set {_strP6=value;}
		}
		private string _strP6Desc="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter6Description
		{
			get {return _strP6Desc;}
			set {_strP6Desc=value;}
		}
		private string _strP7="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter7
		{
			get {return _strP7;}
			set {_strP7=value;}
		}
		private string _strP7Desc="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Parameter7Description
		{
			get {return _strP7Desc;}
			set {_strP7Desc=value;}
		}			
		private string _strOther="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Other
		{
			get {return _strOther;}
			set {_strOther=value;}
		}
		private string _strOtherDesc="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string OtherDescription
		{
			get {return _strOtherDesc;}
			set {_strOtherDesc=value;}
		}
		bool _bDelete=false;
		public bool Delete
		{
			get {return _bDelete;}
			set {_bDelete=value;}
		}
		bool _bAdd=false;
		public bool Add
		{
			get {return _bAdd;}
			set {_bAdd=value;}
		}
		
		public void CopyProperties(RxItemFvsCommandItem p_oFvsCmdItemSource, RxItemFvsCommandItem p_oFvsCmdItemDestination)
		{
			
			
			p_oFvsCmdItemDestination.Index = p_oFvsCmdItemSource.Index;
			p_oFvsCmdItemDestination.SaveIndex = p_oFvsCmdItemSource.SaveIndex;
			p_oFvsCmdItemDestination.RxId = p_oFvsCmdItemSource.RxId;
			p_oFvsCmdItemDestination.FVSCommand = p_oFvsCmdItemSource.FVSCommand;
			p_oFvsCmdItemDestination.FVSCommandId = p_oFvsCmdItemSource.FVSCommandId;
			p_oFvsCmdItemDestination.Other = p_oFvsCmdItemSource.Other;
			p_oFvsCmdItemDestination.OtherDescription = p_oFvsCmdItemSource.OtherDescription;
			p_oFvsCmdItemDestination.Parameter1=p_oFvsCmdItemSource.Parameter1;
			p_oFvsCmdItemDestination.Parameter1Description = p_oFvsCmdItemSource.Parameter1Description;
			p_oFvsCmdItemDestination.Parameter2 = p_oFvsCmdItemSource.Parameter2;
			p_oFvsCmdItemDestination.Parameter2Description = p_oFvsCmdItemSource.Parameter2Description;
			p_oFvsCmdItemDestination.Parameter3 = p_oFvsCmdItemSource.Parameter3;
			p_oFvsCmdItemDestination.Parameter3Description = p_oFvsCmdItemSource.Parameter3Description;
			p_oFvsCmdItemDestination.Parameter4 = p_oFvsCmdItemSource.Parameter4;
			p_oFvsCmdItemDestination.Parameter4Description = p_oFvsCmdItemSource.Parameter4Description;
			p_oFvsCmdItemDestination.Parameter5 = p_oFvsCmdItemSource.Parameter5;
			p_oFvsCmdItemDestination.Parameter5Description = p_oFvsCmdItemSource.Parameter5Description;
			p_oFvsCmdItemDestination.Parameter6 = p_oFvsCmdItemSource.Parameter6;
			p_oFvsCmdItemDestination.Parameter6Description = p_oFvsCmdItemSource.Parameter6Description;
			p_oFvsCmdItemDestination.Parameter7 = p_oFvsCmdItemSource.Parameter7;
			p_oFvsCmdItemDestination.Parameter7Description = p_oFvsCmdItemSource.Parameter7Description;
			p_oFvsCmdItemDestination.Delete = p_oFvsCmdItemSource.Delete;
			p_oFvsCmdItemDestination.Add = p_oFvsCmdItemSource.Add;
					
				
				
			
		}
		//private string _strFvsCmd="";
		//[CategoryAttribute("Estimation Engine And Excel"), BrowsableAttribute(false), DescriptionAttribute("FVS Command")]
		//public string FVSCommand
		//{
		//	get {return _strFVSCmd;}
		//	set {_strFVSCmd=value;}
		//}
		//public void CopyProperties(FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem p_oVarSubItemSource,ref FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem  p_oVarSubItemDestination)
		//{
		//	p_oVarSubItemDestination.Index = p_oVarSubItemSource.Index;
		//	p_oVarSubItemDestination.Description = p_oVarSubItemSource.Description;
		//	p_oVarSubItemDestination.PropertyGridName=p_oVarSubItemSource.PropertyGridName;
		//	p_oVarSubItemDestination.SQLVariableSubstitutionString=p_oVarSubItemSource.SQLVariableSubstitutionString;
		//	p_oVarSubItemDestination.VariableName = p_oVarSubItemSource.VariableName;
		//}

	}
	public class RxItemFvsCommandItem_Collection : System.Collections.CollectionBase
	{
		public RxItemFvsCommandItem_Collection()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public void Add(FIA_Biosum_Manager.RxItemFvsCommandItem m_RxItemFvsCommandItem)
		{
			// vérify if object is not already in
			if (this.List.Contains(m_RxItemFvsCommandItem))
				throw new InvalidOperationException();
 
			// adding it
			this.List.Add(m_RxItemFvsCommandItem);
 
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
		public FIA_Biosum_Manager.RxItemFvsCommandItem Item(int Index)
		{
			// The appropriate item is retrieved from the List object and
			// explicitly cast to the Widget type, then returned to the 
			// caller.
			return (FIA_Biosum_Manager.RxItemFvsCommandItem) List[Index];
		}

	}
	/*********************************************************************************************************
	 **RX Harvest Cost Column Item               
	 *********************************************************************************************************/
	public class RxItemHarvestCostColumnItem
	{
		private int _intIndex;
		
		[CategoryAttribute("General"),ReadOnly(true),DescriptionAttribute("RX Harvest Cost Column Item Index")]
		public int Index
		{
			get {return _intIndex;}
			set {_intIndex = value;}
		}
		private int _intSaveIndex;
		[CategoryAttribute("General"),ReadOnly(true),DescriptionAttribute("RX Harvest Cost Column Item Save Index")]
		public int SaveIndex
		{
			get {return _intSaveIndex;}
			set {_intSaveIndex = value;}
		}

		private string _strRxId="";
		[CategoryAttribute("General"),DescriptionAttribute("RX Indentifier")]
		public string RxId
		{
			get {return _strRxId;}
			set {_strRxId=value;}
		}
		private string _strHarvestCostColumn="";
		[CategoryAttribute("General"),DescriptionAttribute("Harvest Cost Column")]
		public string HarvestCostColumn
		{
			get {return _strHarvestCostColumn;}
			set {_strHarvestCostColumn=value;}
		}
		private string _strDesc="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Description
		{
			get {return _strDesc;}
			set {_strDesc=value;}
		}
		bool _bDelete=false;
		public bool Delete
		{
			get {return _bDelete;}
			set {_bDelete=value;}
		}
		bool _bAdd=false;
		public bool Add
		{
			get {return _bAdd;}
			set {_bAdd=value;}
		}
		
		public void CopyProperties(RxItemHarvestCostColumnItem p_oRxHarvestCostColumnItemSource,RxItemHarvestCostColumnItem p_oRxHarvestCostColumnItemDestination)
		{
			
			
			p_oRxHarvestCostColumnItemDestination.Index = p_oRxHarvestCostColumnItemSource.Index;
			p_oRxHarvestCostColumnItemDestination.SaveIndex = p_oRxHarvestCostColumnItemSource.SaveIndex;
			p_oRxHarvestCostColumnItemDestination.RxId = p_oRxHarvestCostColumnItemSource.RxId;
			p_oRxHarvestCostColumnItemDestination.HarvestCostColumn = p_oRxHarvestCostColumnItemSource.HarvestCostColumn;
			p_oRxHarvestCostColumnItemDestination.Delete = p_oRxHarvestCostColumnItemSource.Delete;
			p_oRxHarvestCostColumnItemDestination.Add = p_oRxHarvestCostColumnItemSource.Add;
			
		}
		

	}
	public class RxItemHarvestCostColumnItem_Collection : System.Collections.CollectionBase
	{
		public RxItemHarvestCostColumnItem_Collection()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public void Add(FIA_Biosum_Manager.RxItemHarvestCostColumnItem m_RxItemHarvestCostColumnItem)
		{
			// vérify if object is not already in
			if (this.List.Contains(m_RxItemHarvestCostColumnItem))
				throw new InvalidOperationException();
 
			// adding it
			this.List.Add(m_RxItemHarvestCostColumnItem);
 
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
		public FIA_Biosum_Manager.RxItemHarvestCostColumnItem Item(int Index)
		{
			// The appropriate item is retrieved from the List object and
			// explicitly cast to the Widget type, then returned to the 
			// caller.
			return (FIA_Biosum_Manager.RxItemHarvestCostColumnItem) List[Index];
		}

	}
}

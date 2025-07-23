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

			this.m_oQueries = new Queries();
			m_oQueries.m_oFvs.LoadDatasource=true;
			m_oQueries.LoadDatasourcesNew(true);

			this.m_oLvAlternateColors.InitializeRowCollection();      
			this.lstRx.Clear();
			this.lstRx.Columns.Add(" ",2,HorizontalAlignment.Left);
			this.lstRx.Columns.Add("Rx", 60, HorizontalAlignment.Left);
			this.lstRx.Columns.Add("Description", 300, HorizontalAlignment.Left);
			
					

			this.m_intError=0;

			this.m_oRxItem_Collection.Clear();

            this.m_oRxTools.LoadAllRxItemsFromTableIntoRxCollection(m_oQueries, this.m_oRxItem_Collection);
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
			oItem.Add=true;
			this.m_oRxItem_Collection.Add(oItem);
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
			oItem.Add=true;
			this.m_oRxItem_Collection.Add(oItem);
			
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
			oItem.Add=true;
			//oItem.ReferenceFvsCommandsCollection = this.m_oRxItemFvsCommandItem_Collection;
			this.m_oRxItem_Collection.Add(oItem);
			this.lstRx.Items[0].Selected = true;
			this.lstRx.EndUpdate();
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			this.m_oLvAlternateColors.ListView();
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
			
			string strFields;
			string strValues;

            DataMgr oDataMgr = new DataMgr();
            string strSource = m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Rx);
            if (!string.IsNullOrEmpty(strSource))
            {
                string strConn = oDataMgr.GetConnectionString(strSource);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    //delete all records from rx table
                    oDataMgr.m_strSQL = "DELETE FROM " + this.m_oQueries.m_oFvs.m_strRxTable;
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    if (oDataMgr.m_intError == 0)
                    {
                        //delete all records from the rx harvest cost columns table
                        oDataMgr.m_strSQL = "DELETE FROM " + this.m_oQueries.m_oFvs.m_strRxHarvestCostColumnsTable;
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                        if (oDataMgr.m_intError == 0)
                        {
                            //insert all the records
                            for (x = 0; x <= this.m_oRxItem_Collection.Count - 1; x++)
                            {
                                if (m_oRxItem_Collection.Item(x).Delete == false)
                                {
                                    //insert the rx record
                                    strFields = "";
                                    strValues = "";

                                    strFields = "rx,description,harvestmethodLowSlope,harvestmethodsteepslope";
                                    strValues = "'" + m_oRxItem_Collection.Item(x).RxId.Trim() + "',";
                                    strValues = strValues + "'" + oDataMgr.FixString(m_oRxItem_Collection.Item(x).Description.Trim(), "'", "''") + "',";
                                    strValues = strValues + "'" + oDataMgr.FixString(m_oRxItem_Collection.Item(x).HarvestMethodLowSlope.Trim(), "'", "''") + "',";
                                    strValues = strValues + "'" + oDataMgr.FixString(m_oRxItem_Collection.Item(x).HarvestMethodSteepSlope.Trim(), "'", "''") + "'";
                                    oDataMgr.m_strSQL = Queries.GetInsertSQL(strFields, strValues, m_oQueries.m_oFvs.m_strRxTable);
                                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                                    //insert all the harvest cost columns for the rx
                                    if (this.m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection != null)
                                    {
                                        for (y = 0; y <= this.m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Count - 1; y++)
                                        {
                                            if (m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).Delete == false)
                                            {
                                                if (this.m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).RxId ==
                                                    this.m_oRxItem_Collection.Item(x).RxId)
                                                {
                                                    strValues = "";
                                                    strFields = "rx,ColumnName,description";
                                                    strValues = strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).RxId.Trim() + "',";
                                                    strValues = strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).HarvestCostColumn.Trim() + "',";
                                                    strValues = strValues + "'" + m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).Description.Trim() + "'";
                                                    oDataMgr.m_strSQL = Queries.GetInsertSQL(strFields, strValues, m_oQueries.m_oFvs.m_strRxHarvestCostColumnsTable);
                                                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                                }
                                            }
                                            else
                                            {
                                                //delete rx + column_name from the harvest cost column table
                                                oDataMgr.m_strSQL = "DELETE FROM " + this.m_oQueries.m_oFvs.m_strRxHarvestCostColumnsTable + " " +
                                                        "WHERE TRIM(RX)='" + m_oRxItem_Collection.Item(x).RxId.Trim() + "' AND " +
                                                        "TRIM(ColumnName)='" + m_oRxItem_Collection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).HarvestCostColumn.Trim() + "'";
                                                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    //delete all rx items from the harvest cost column table
                                    oDataMgr.m_strSQL = "DELETE FROM " + this.m_oQueries.m_oFvs.m_strRxHarvestCostColumnsTable + " " +
                                        "WHERE TRIM(RX)='" + m_oRxItem_Collection.Item(x).RxId.Trim() + "'";
                                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                }
                            }
                        }
                    }
                }
            }
		    if (this.m_intError==0 && oDataMgr.m_intError==0)
		    {
				this.btnSave.Enabled=false;
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

		//public void LoadAllRxItemsFromTableIntoRxCollection(string p_strDbFile,Queries p_oQueries,RxItem_Collection p_oRxItemCollection)
		//{
		//	ado_data_access oAdo = new ado_data_access();
		//	oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile,"",""));
		//	if (oAdo.m_intError==0)
		//	{ 
  //              this.LoadAllRxItemsFromTableIntoRxCollection(p_oQueries, p_oRxItemCollection);
		//	}
		//	m_intError=oAdo.m_intError;
		//	oAdo.CloseConnection(oAdo.m_OleDbConnection);
		//	oAdo=null;
		//}

        public void LoadAllRxItemsFromTableIntoRxCollection(Queries p_oQueries, RxItem_Collection p_oRxItemCollection)
        {

            int x;
            int intIndex = 0;
            DataMgr oDataMgr = new DataMgr();
            string strSource = p_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Rx);
            if (!string.IsNullOrEmpty(strSource))
            {
                string strConn = oDataMgr.GetConnectionString(strSource);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    oDataMgr.m_strSQL = "SELECT * FROM " + p_oQueries.m_oFvs.m_strRxTable + " ORDER BY RX";
                    oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
                    if (oDataMgr.m_DataReader.HasRows)
                    {
                        while (oDataMgr.m_DataReader.Read())
                        {

                            if (oDataMgr.m_DataReader["rx"] != System.DBNull.Value)
                            {

                                RxItem oRxItem = new RxItem();
                                oRxItem.m_oHarvestCostColumnItem_Collection1 = new RxItemHarvestCostColumnItem_Collection();
                                oRxItem.ReferenceHarvestCostColumnCollection = oRxItem.m_oHarvestCostColumnItem_Collection1;
                                oRxItem.Index = intIndex;
                                oRxItem.RxId = Convert.ToString(oDataMgr.m_DataReader["rx"]);

                                if (oDataMgr.m_DataReader["description"] != System.DBNull.Value)
                                {
                                    oRxItem.Description = Convert.ToString(oDataMgr.m_DataReader["description"]);
                                }
                                if (oDataMgr.m_DataReader["HarvestMethodLowSlope"] != System.DBNull.Value)
                                {
                                    oRxItem.HarvestMethodLowSlope = Convert.ToString(oDataMgr.m_DataReader["HarvestMethodLowSlope"]).Trim();
                                }
                                if (oDataMgr.m_DataReader["HarvestMethodSteepSlope"] != System.DBNull.Value)
                                {
                                    oRxItem.HarvestMethodSteepSlope = Convert.ToString(oDataMgr.m_DataReader["HarvestMethodSteepSlope"]).Trim();
                                }

                                p_oRxItemCollection.Add(oRxItem);
                                intIndex++;
                            }
                        }

                    }
                    oDataMgr.m_DataReader.Close();
                }
            }
            intIndex = 0;
            //
            //HARVEST COST COLUMNS
            //
            strSource = p_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.RxHarvestCostColumns);
            if (!string.IsNullOrEmpty(strSource))
            {
                string strConn = oDataMgr.GetConnectionString(strSource);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    oDataMgr.m_strSQL = "SELECT * FROM " + p_oQueries.m_oFvs.m_strRxHarvestCostColumnsTable;
                    oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
                    if (oDataMgr.m_DataReader.HasRows)
                    {
                        while (oDataMgr.m_DataReader.Read())
                        {
                            if (oDataMgr.m_DataReader["rx"] != System.DBNull.Value &&
                                oDataMgr.m_DataReader["ColumnName"] != System.DBNull.Value)
                            {
                                for (x = 0; x <= p_oRxItemCollection.Count - 1; x++)
                                {
                                    if (p_oRxItemCollection.Item(x).RxId ==
                                        oDataMgr.m_DataReader["rx"].ToString().Trim())
                                    {
                                        FIA_Biosum_Manager.RxItemHarvestCostColumnItem oItem = new RxItemHarvestCostColumnItem();
                                        oItem.Index = p_oRxItemCollection.Item(x).m_oHarvestCostColumnItem_Collection1.Count;
                                        oItem.SaveIndex = p_oRxItemCollection.Item(x).m_oHarvestCostColumnItem_Collection1.Count;
                                        oItem.RxId = p_oRxItemCollection.Item(x).RxId;
                                        oItem.HarvestCostColumn = oDataMgr.m_DataReader["ColumnName"].ToString().Trim();

                                        if (oDataMgr.m_DataReader["description"] != System.DBNull.Value)
                                        {
                                            oItem.Description = oDataMgr.m_DataReader["description"].ToString().Trim();
                                        }
                                        p_oRxItemCollection.Item(x).m_oHarvestCostColumnItem_Collection1.Add(oItem);
                                        intIndex++;
                                    }
                                }

                            }
                        }
                    }
                    oDataMgr.m_DataReader.Close();
                }
            }

            //
            //PACKAGE MEMBERSHIP
            //
            strSource = p_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.RxPackage);
            if (!string.IsNullOrEmpty(strSource))
            {
                for (x = 0; x <= p_oRxItemCollection.Count - 1; x++)
                {
                    p_oRxItemCollection.Item(x).RxPackageMemberList = "";
                }
                string strConn = oDataMgr.GetConnectionString(strSource);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    oDataMgr.m_strSQL = "SELECT * FROM " + p_oQueries.m_oFvs.m_strRxPackageTable;
                    oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
                    if (oDataMgr.m_DataReader.HasRows)
                    {

                        while (oDataMgr.m_DataReader.Read())
                        {
                            string strPackage = "";
                            string strCycle1Rx = "";
                            string strCycle2Rx = "";
                            string strCycle3Rx = "";
                            string strCycle4Rx = "";

                            if (oDataMgr.m_DataReader["rxpackage"] != System.DBNull.Value)
                            {
                                strPackage = oDataMgr.m_DataReader["rxpackage"].ToString().Trim();
                                if (strPackage.Trim().Length > 0)
                                {

                                    //sim year 1
                                    if (oDataMgr.m_DataReader["simyear1_rx"] != System.DBNull.Value)
                                    {
                                        strCycle1Rx = Convert.ToString(oDataMgr.m_DataReader["simyear1_rx"]).Trim();
                                    }
                                    //sim year 2
                                    if (oDataMgr.m_DataReader["simyear2_rx"] != System.DBNull.Value)
                                    {
                                        strCycle2Rx = Convert.ToString(oDataMgr.m_DataReader["simyear2_rx"]).Trim();
                                    }
                                    //sim year 3
                                    if (oDataMgr.m_DataReader["simyear3_rx"] != System.DBNull.Value)
                                    {
                                        strCycle3Rx = Convert.ToString(oDataMgr.m_DataReader["simyear3_rx"]).Trim();
                                    }
                                    //sim year 4
                                    if (oDataMgr.m_DataReader["simyear4_rx"] != System.DBNull.Value)
                                    {
                                        strCycle4Rx = Convert.ToString(oDataMgr.m_DataReader["simyear4_rx"]).Trim();
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
                    oDataMgr.m_DataReader.Close();
                }
            }

            for (x = 0; x <= p_oRxItemCollection.Count - 1; x++)
            {
                //remove the last comma from the end of the string
                if (p_oRxItemCollection.Item(x).RxPackageMemberList.Trim().Length > 0)
                    p_oRxItemCollection.Item(x).RxPackageMemberList = p_oRxItemCollection.Item(x).RxPackageMemberList.Substring(0, p_oRxItemCollection.Item(x).RxPackageMemberList.Length - 1);
            }
        }

        public void LoadAllRxPackageItems(RxPackageItem_Collection p_oRxPackageItemCollection)
        {
            Queries oQueries = new Queries();
            oQueries.m_oFvs.LoadDatasource = true;
			// pulls from master databases - keep as Access version for now
            oQueries.LoadDatasourcesNew(true);
            this.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(oQueries, p_oRxPackageItemCollection);
            oQueries = null;
        }

        public void LoadAllRxPackageItemsFromTableIntoRxPackageCollection(Queries p_oQueries,
                                                                          RxPackageItem_Collection p_oRxPackageItemCollection)
        {
            int x;
            int intIndex = 0;
            DataMgr oDataMgr = new DataMgr();
            string strSource = p_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.RxPackage);
            if (!string.IsNullOrEmpty(strSource))
            {
                string strConn = oDataMgr.GetConnectionString(strSource);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    oDataMgr.m_strSQL = "SELECT * FROM " + p_oQueries.m_oFvs.m_strRxPackageTable + " ORDER BY RXPACKAGE";
                    oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
                    if (oDataMgr.m_DataReader.HasRows)
                    {
                        while (oDataMgr.m_DataReader.Read())
                        {
                            if (oDataMgr.m_DataReader["rxpackage"] != System.DBNull.Value)
                            {
                                RxPackageItem oRxPackageItem = new RxPackageItem();
                                oRxPackageItem.Index = intIndex;
                                oRxPackageItem.RxPackageId = Convert.ToString(oDataMgr.m_DataReader["rxpackage"]);
                                if (oDataMgr.m_DataReader["description"] != System.DBNull.Value)
                                {
                                    oRxPackageItem.Description = Convert.ToString(oDataMgr.m_DataReader["description"]);
                                }
                                //treatment cycle length
                                if (oDataMgr.m_DataReader["rxcycle_length"] != System.DBNull.Value)
                                {
                                    oRxPackageItem.RxCycleLength = Convert.ToInt32(oDataMgr.m_DataReader["rxcycle_length"]);
                                }
                                //sim year 1
                                if (oDataMgr.m_DataReader["simyear1_rx"] != System.DBNull.Value)
                                {
                                    oRxPackageItem.SimulationYear1Rx = Convert.ToString(oDataMgr.m_DataReader["simyear1_rx"]);
                                }
                                if (oDataMgr.m_DataReader["simyear1_fvscycle"] != System.DBNull.Value)
                                {
                                    oRxPackageItem.SimulationYear1Fvs = Convert.ToString(oDataMgr.m_DataReader["simyear1_fvscycle"]);
                                }
                                //sim year 2
                                if (oDataMgr.m_DataReader["simyear2_rx"] != System.DBNull.Value)
                                {
                                    oRxPackageItem.SimulationYear2Rx = Convert.ToString(oDataMgr.m_DataReader["simyear2_rx"]);
                                }
                                if (oDataMgr.m_DataReader["simyear2_fvscycle"] != System.DBNull.Value)
                                {
                                    oRxPackageItem.SimulationYear2Fvs = Convert.ToString(oDataMgr.m_DataReader["simyear2_fvscycle"]);
                                }
                                //sim year 3
                                if (oDataMgr.m_DataReader["simyear3_rx"] != System.DBNull.Value)
                                {
                                    oRxPackageItem.SimulationYear3Rx = Convert.ToString(oDataMgr.m_DataReader["simyear3_rx"]);
                                }
                                if (oDataMgr.m_DataReader["simyear3_fvscycle"] != System.DBNull.Value)
                                {
                                    oRxPackageItem.SimulationYear3Fvs = Convert.ToString(oDataMgr.m_DataReader["simyear3_fvscycle"]);
                                }
                                //sim year 4
                                if (oDataMgr.m_DataReader["simyear4_rx"] != System.DBNull.Value)
                                {
                                    oRxPackageItem.SimulationYear4Rx = Convert.ToString(oDataMgr.m_DataReader["simyear4_rx"]);
                                }
                                if (oDataMgr.m_DataReader["simyear4_fvscycle"] != System.DBNull.Value)
                                {
                                    oRxPackageItem.SimulationYear4Fvs = Convert.ToString(oDataMgr.m_DataReader["simyear4_fvscycle"]);
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
                    oDataMgr.m_DataReader.Close();
                }
            }
        }

        public string GetRxPackageRunTitle(string strVariant, RxPackageItem p_rxPackageItem)
        {

            //string strFile = "FVSOUT_" + p_OleDbReader["fvs_variant"].ToString().Trim() + "_P" + p_OleDbReader["rxpackage"].ToString().Trim();
            string strTitle = $@"FVSOUT_{strVariant}_P{p_rxPackageItem.RxPackageId}";
            if (p_rxPackageItem.SimulationYear1Rx.Trim().Length == 0 ||
                p_rxPackageItem.SimulationYear1Rx.Trim() == "GP")
            {
                strTitle = strTitle + "-000";
            }
            else
            {
                strTitle = strTitle + "-" + p_rxPackageItem.SimulationYear1Rx;
            }
            if (p_rxPackageItem.SimulationYear2Rx.Trim().Length == 0 ||
                p_rxPackageItem.SimulationYear2Rx.Trim() == "GP")
            {
                strTitle = strTitle + "-000";
            }
            else
            {
                strTitle = strTitle + "-" + p_rxPackageItem.SimulationYear2Rx.Trim();
            }
            if (p_rxPackageItem.SimulationYear3Rx.Trim().Length == 0 ||
                p_rxPackageItem.SimulationYear3Rx.Trim() == "GP")
            {
                strTitle = strTitle + "-000";
            }
            else
            {
                strTitle = strTitle + "-" + p_rxPackageItem.SimulationYear3Rx.Trim();

            }
            if (p_rxPackageItem.SimulationYear4Rx.Trim().Length == 0 ||
                p_rxPackageItem.SimulationYear4Rx.Trim() == "GP")
            {
                strTitle = strTitle + "-000";
            }
            else
            {

                strTitle = strTitle + "-" + p_rxPackageItem.SimulationYear4Rx.Trim();
            }
            return strTitle;
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

		public string GetListOfFVSVariantsInPlotTable(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strPlotTable)
		{
			return p_oAdo.CreateCommaDelimitedList(p_oConn,"SELECT DISTINCT fvs_variant FROM " +  p_strPlotTable +  " WHERE fvs_variant IS NOT NULL","");
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
        public void CheckCutListDbExist()
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
            }
        }

        /// <summary>
        /// Load the Rx Harvest Methods into the processor scenario dropdown combo boxes for 
        /// low slope and steep slope
        /// </summary>
        /// <param name="p_oConn"></param>
        /// <param name="p_oQueries"></param>
        /// <param name="p_cmbHarvestMethod"></param>
        /// <param name="p_cmbHarvestMethodSteepSlope"></param>
        public void LoadRxHarvestMethods(string p_strDbFile, Queries p_oQueries, ComboBox p_cmbHarvestMethod, ComboBox p_cmbHarvestMethodSteepSlope,
            System.Collections.Generic.IDictionary<string, string> dictDescr, System.Collections.Generic.IDictionary<string, string> dictSteepDescr)
        {
            p_cmbHarvestMethod.Items.Clear();
            p_cmbHarvestMethodSteepSlope.Items.Clear();
            DataMgr dataMgr = new DataMgr();
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
                                    if (!dictSteepDescr.ContainsKey(dataMgr.m_DataReader["method"].ToString().Trim()))
                                    {
                                        dictSteepDescr.Add(dataMgr.m_DataReader["method"].ToString().Trim(),
                                            dataMgr.m_DataReader["description"].ToString().Trim());
                                    }
                                }
                                else
                                {
                                    p_cmbHarvestMethod.Items.Add(dataMgr.m_DataReader["method"].ToString().Trim());
                                    if (!dictDescr.ContainsKey(dataMgr.m_DataReader["method"].ToString().Trim()))
                                    {
                                        dictDescr.Add(dataMgr.m_DataReader["method"].ToString().Trim(),
                                            dataMgr.m_DataReader["description"].ToString().Trim());
                                    }
                                }
                            }
                        }
                    }
                    dataMgr.m_DataReader.Close();
                }
            }
        }

        public void LoadFVSOutputPrePostRxCycleSeqNum(string strDbConn, FVSPrePostSeqNumItem_Collection p_oCollection)
        {
            FVSPrePostSeqNumItem oItem = null;
            int x;

            p_oCollection.Clear();

            DataMgr oDataMgr = new DataMgr();
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strDbConn))
            {
                conn.Open();
                oDataMgr.m_strSQL = "SELECT * FROM " + Tables.FVS.DefaultFVSPrePostSeqNumTable + " ORDER BY PREPOST_SEQNUM_ID";
                oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
                if (oDataMgr.m_DataReader.HasRows)
                {
                    while (oDataMgr.m_DataReader.Read())
                    {
                        oItem = new FVSPrePostSeqNumItem();
                        oItem.PrePostSeqNumId = Convert.ToInt32(oDataMgr.m_DataReader["PREPOST_SEQNUM_ID"]);
                        oItem.TableName = Convert.ToString(oDataMgr.m_DataReader["TableName"]).Trim();
                        int intAssignedCount = 0;
                        //
                        //PRE
                        //
                        if (oDataMgr.m_DataReader["RXCYCLE1_PRE_SEQNUM"] != DBNull.Value)
                        {
                            oItem.RxCycle1PreSeqNum = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE1_PRE_SEQNUM"]).Trim();
                            intAssignedCount++;
                        }
                        else
                        {
                            oItem.RxCycle1PreSeqNum = "Not Used";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE2_PRE_SEQNUM"] != DBNull.Value)
                        {
                            oItem.RxCycle2PreSeqNum = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE2_PRE_SEQNUM"]).Trim();
                            intAssignedCount++;
                        }
                        else
                        {
                            oItem.RxCycle2PreSeqNum = "Not Used";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE3_PRE_SEQNUM"] != DBNull.Value)
                        {
                            oItem.RxCycle3PreSeqNum = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE3_PRE_SEQNUM"]).Trim();
                            intAssignedCount++;
                        }
                        else
                        {
                            oItem.RxCycle3PreSeqNum = "Not Used";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE4_PRE_SEQNUM"] != DBNull.Value)
                        {
                            oItem.RxCycle4PreSeqNum = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE4_PRE_SEQNUM"]).Trim();
                            intAssignedCount++;
                        }
                        else
                        {
                            oItem.RxCycle4PreSeqNum = "Not Used";
                        }
                        //
                        //POST
                        //
                        if (oDataMgr.m_DataReader["RXCYCLE1_POST_SEQNUM"] != DBNull.Value)
                        {
                            oItem.RxCycle1PostSeqNum = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE1_POST_SEQNUM"]).Trim();
                            intAssignedCount++;
                        }
                        else
                        {
                            oItem.RxCycle1PostSeqNum = "Not Used";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE2_POST_SEQNUM"] != DBNull.Value)
                        {
                            oItem.RxCycle2PostSeqNum = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE2_POST_SEQNUM"]).Trim();
                            intAssignedCount++;
                        }
                        else
                        {
                            oItem.RxCycle2PostSeqNum = "Not Used";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE3_POST_SEQNUM"] != DBNull.Value)
                        {
                            oItem.RxCycle3PostSeqNum = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE3_POST_SEQNUM"]).Trim();
                            intAssignedCount++;
                        }
                        else
                        {
                            oItem.RxCycle3PostSeqNum = "Not Used";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE4_POST_SEQNUM"] != DBNull.Value)
                        {
                            oItem.RxCycle4PostSeqNum = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE4_POST_SEQNUM"]).Trim();
                            intAssignedCount++;
                        }
                        else
                        {
                            oItem.RxCycle4PostSeqNum = "Not Used";
                        }
                        //
                        //BASEYEAR
                        //
                        if (oDataMgr.m_DataReader["RXCYCLE1_PRE_BASEYR_YN"] != DBNull.Value)
                        {
                            oItem.RxCycle1PreSeqNumBaseYearYN = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE1_PRE_BASEYR_YN"]).Trim();
                        }
                        else
                        {
                            oItem.RxCycle1PreSeqNumBaseYearYN = "N";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE2_PRE_BASEYR_YN"] != DBNull.Value)
                        {
                            oItem.RxCycle2PreSeqNumBaseYearYN = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE2_PRE_BASEYR_YN"]).Trim();
                        }
                        else
                        {
                            oItem.RxCycle2PreSeqNumBaseYearYN = "N";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE3_PRE_BASEYR_YN"] != DBNull.Value)
                        {
                            oItem.RxCycle3PreSeqNumBaseYearYN = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE3_PRE_BASEYR_YN"]).Trim();
                        }
                        else
                        {
                            oItem.RxCycle3PreSeqNumBaseYearYN = "N";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE4_PRE_BASEYR_YN"] != DBNull.Value)
                        {
                            oItem.RxCycle4PreSeqNumBaseYearYN = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE4_PRE_BASEYR_YN"]).Trim();
                        }
                        else
                        {
                            oItem.RxCycle4PreSeqNumBaseYearYN = "N";
                        }
                        //
                        //STRCLASS BEFORE REMOVAL
                        //
                        if (oDataMgr.m_DataReader["RXCYCLE1_PRE_BEFORECUT_YN"] != DBNull.Value)
                        {
                            oItem.RxCycle1PreStrClassBeforeTreeRemovalYN = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE1_PRE_BEFORECUT_YN"]).Trim();
                        }
                        else
                        {
                            oItem.RxCycle1PreStrClassBeforeTreeRemovalYN = "Y";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE2_PRE_BEFORECUT_YN"] != DBNull.Value)
                        {
                            oItem.RxCycle2PreStrClassBeforeTreeRemovalYN = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE2_PRE_BEFORECUT_YN"]).Trim();
                        }
                        else
                        {
                            oItem.RxCycle2PreStrClassBeforeTreeRemovalYN = "Y";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE3_PRE_BEFORECUT_YN"] != DBNull.Value)
                        {
                            oItem.RxCycle3PreStrClassBeforeTreeRemovalYN = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE3_PRE_BEFORECUT_YN"]).Trim();
                        }
                        else
                        {
                            oItem.RxCycle3PreStrClassBeforeTreeRemovalYN = "Y";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE4_PRE_BEFORECUT_YN"] != DBNull.Value)
                        {
                            oItem.RxCycle4PreStrClassBeforeTreeRemovalYN = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE4_PRE_BEFORECUT_YN"]).Trim();
                        }
                        else
                        {
                            oItem.RxCycle4PreStrClassBeforeTreeRemovalYN = "Y";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE1_POST_BEFORECUT_YN"] != DBNull.Value)
                        {
                            oItem.RxCycle1PostStrClassBeforeTreeRemovalYN = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE1_POST_BEFORECUT_YN"]).Trim();
                        }
                        else
                        {
                            oItem.RxCycle1PostStrClassBeforeTreeRemovalYN = "N";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE2_POST_BEFORECUT_YN"] != DBNull.Value)
                        {
                            oItem.RxCycle2PostStrClassBeforeTreeRemovalYN = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE2_POST_BEFORECUT_YN"]).Trim();
                        }
                        else
                        {
                            oItem.RxCycle2PostStrClassBeforeTreeRemovalYN = "N";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE3_POST_BEFORECUT_YN"] != DBNull.Value)
                        {
                            oItem.RxCycle3PostStrClassBeforeTreeRemovalYN = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE3_POST_BEFORECUT_YN"]).Trim();
                        }
                        else
                        {
                            oItem.RxCycle3PostStrClassBeforeTreeRemovalYN = "N";
                        }
                        if (oDataMgr.m_DataReader["RXCYCLE4_POST_BEFORECUT_YN"] != DBNull.Value)
                        {
                            oItem.RxCycle4PostStrClassBeforeTreeRemovalYN = Convert.ToString(oDataMgr.m_DataReader["RXCYCLE4_POST_BEFORECUT_YN"]).Trim();
                        }
                        else
                        {
                            oItem.RxCycle4PostStrClassBeforeTreeRemovalYN = "N";
                        }

                        if (oDataMgr.m_DataReader["USE_SUMMARY_TABLE_SEQNUM_YN"] != DBNull.Value)
                        {
                            oItem.UseSummaryTableSeqNumYN = Convert.ToString(oDataMgr.m_DataReader["USE_SUMMARY_TABLE_SEQNUM_YN"]).Trim();
                        }
                        else
                        {
                            oItem.UseSummaryTableSeqNumYN = "Y";
                        }

                        oItem.Type = oDataMgr.m_DataReader["TYPE"].ToString().Trim();
                        // Check for rxpackages in the fvs_output_prepost_seqnum_rxpackage_assignment for assigned package for "C"
                        // tables. If missing report 0 because FVS Out won't use assignment without rxPackages
                        if (oItem.Type.Equals("C"))
                        {
                            string strRxPackageSql = $@"SELECT COUNT(*) FROM {Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable} 
                            WHERE PREPOST_SEQNUM_ID = {oItem.PrePostSeqNumId}";
                            double rxPackageCount = oDataMgr.getSingleDoubleValueFromSQLQuery(conn, strRxPackageSql, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable);
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
                    oDataMgr.m_DataReader.Close();
                    //rx package assignments for custom definitions
                    for (x = 0; x <= p_oCollection.Count - 1; x++)
                    {
                        if (p_oCollection.Item(x).Type == "C")
                        {
                            if (p_oCollection.Item(x).m_FVSPrePostSeqNumRxPackageAssgnItem_Collection1 == null)
                                p_oCollection.Item(x).m_FVSPrePostSeqNumRxPackageAssgnItem_Collection1 = new FVSPrePostSeqNumRxPackageAssgnItem_Collection();
                            else
                                p_oCollection.Item(x).m_FVSPrePostSeqNumRxPackageAssgnItem_Collection1.Clear();

                            oDataMgr.m_strSQL = "SELECT * FROM " + Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable + " " +
                                              "WHERE PREPOST_SEQNUM_ID=" + p_oCollection.Item(x).PrePostSeqNumId;
                            oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
                            if (oDataMgr.m_DataReader.HasRows)
                            {
                                string strList = "";
                                while (oDataMgr.m_DataReader.Read())
                                {

                                    FVSPrePostSeqNumRxPackageAssgnItem oPackageItem = new FVSPrePostSeqNumRxPackageAssgnItem();
                                    oPackageItem.PrePostSeqNumId = p_oCollection.Item(x).PrePostSeqNumId;
                                    oPackageItem.RxPackageId = oDataMgr.m_DataReader["RXPACKAGE"].ToString().Trim();
                                    strList = strList + oPackageItem.RxPackageId + ",";
                                    p_oCollection.Item(x).m_FVSPrePostSeqNumRxPackageAssgnItem_Collection1.Add(oPackageItem);

                                }
                                if (strList.Trim().Length > 0) strList = strList.Substring(0, strList.Length - 1);
                                p_oCollection.Item(x).RxPackageList = strList;
                            }
                            oDataMgr.m_DataReader.Close();
                            p_oCollection.Item(x).ReferenceFVSPrePostSeqNumRxPackageAssgnCollection =
                                p_oCollection.Item(x).m_FVSPrePostSeqNumRxPackageAssgnItem_Collection1;
                        }
                    }
                }
            }
            
        }

		public void CreateTableLinksToFVSPrePostTables(string p_strDestinationDbFile)
        {
			string strFVSPrePostPathAndDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.FVS.DefaultFVSOutPrePostDbFile;
			string strFVSWeightedPathAndDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile;
			dao_data_access oDao = new dao_data_access();
			DataMgr oDataMgr = new DataMgr();

			if (!System.IO.File.Exists(p_strDestinationDbFile))
            {
				oDao.CreateMDB(p_strDestinationDbFile);
            }

			string[] strTableNames = new string[1];
			int intCount = 0;

			// attach pre/post tables
			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strFVSPrePostPathAndDbFile)))
			{
				conn.Open();
				strTableNames = oDataMgr.getTableNames(conn);
				intCount = strTableNames.Length;
				conn.Close();
			}
			if (oDataMgr.m_intError == 0)
			{
				if (intCount > 0)
				{
					ODBCMgr odbcmgr = new ODBCMgr();
					if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.FVSPrePostDsnName))
					{
						odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.FVSPrePostDsnName);
					}
					odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.FVSPrePostDsnName, strFVSPrePostPathAndDbFile);

					for (int x = 0; x <= intCount - 1; x++)
					{
						oDao.CreateSQLiteTableLink(p_strDestinationDbFile, strTableNames[x], strTableNames[x], ODBCMgr.DSN_KEYS.FVSPrePostDsnName, strFVSPrePostPathAndDbFile);
						if (oDao.m_intError != 0)
						{
							oDao.m_strError = strTableNames[x] + " !!Error Creating Table Link!!!";
							this.m_intError = oDao.m_intError;
							break;
						}
					}

					if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.FVSPrePostDsnName))
					{
						odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.FVSPrePostDsnName);
					}
				}
			}

			// attach pre/post weighted tables
			if (System.IO.File.Exists(strFVSWeightedPathAndDbFile))
            {
				using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strFVSWeightedPathAndDbFile)))
				{
					conn.Open();
					strTableNames = oDataMgr.getTableNames(conn);
					intCount = strTableNames.Length;
					conn.Close();
				}
			}
			if (oDataMgr.m_intError == 0)
			{
				if (intCount > 0)
				{
					ODBCMgr odbcmgr = new ODBCMgr();
					if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.PrePostFvsWeightedDsnName))
					{
						odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.PrePostFvsWeightedDsnName);
					}
					odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.PrePostFvsWeightedDsnName, strFVSWeightedPathAndDbFile);

					for (int x = 0; x <= intCount - 1; x++)
					{
						oDao.CreateSQLiteTableLink(p_strDestinationDbFile, strTableNames[x], strTableNames[x], ODBCMgr.DSN_KEYS.PrePostFvsWeightedDsnName, strFVSWeightedPathAndDbFile);
						if (oDao.m_intError != 0)
						{
							oDao.m_strError = strTableNames[x] + " !!Error Creating Table Link!!!";
							this.m_intError = oDao.m_intError;
							break;
						}
					}
				}
			}
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

        //@ToDo: Need to work through all the permutations of Queries.FVS.SqliteFVSOutputTable_AuditPrePostGenericSQL
        // with the tempTable approach
        public void CreateFVSPrePostSeqNumTables(string strTargetDb, FVSPrePostSeqNumItem p_oItem, string p_strSourceTableName, string p_strSourceLinkedTableName, 
            bool p_bAudit, bool p_bDebug, string p_strDebugFile, string p_strRunTitle, string p_strPotfireTableName, string p_strTempTable)
        {
            DataMgr dataMgr = new DataMgr();
            if (p_strSourceTableName.Trim().ToUpper() == "FVS_CASES") return;
            int x;
            string[] strSQLArray = null;
            string strPrePostSeqNumMatrixTable = p_strSourceTableName + "_PREPOST_SEQNUM_MATRIX";
            string strAuditPrePostSeqNumCountsTable = "audit_" + p_strSourceTableName + "_prepost_seqnum_counts_table";
            if (p_strSourceTableName.ToUpper().Equals("FVS_POTFIRE_TEMP"))
            {
                strAuditPrePostSeqNumCountsTable = "audit_" + p_strPotfireTableName + "_prepost_seqnum_counts_table";
                strPrePostSeqNumMatrixTable = p_strPotfireTableName + "_PREPOST_SEQNUM_MATRIX";
            }
            string strAuditYearCountsTable = "audit_" + p_strSourceTableName + "_year_counts_table";
            string strVariant = p_strRunTitle.Substring(7, 2);
            string strRxPackage = p_strRunTitle.Substring(11,3);

            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dataMgr.GetConnectionString(strTargetDb)))
            {
                conn.Open();
                // Attach the FVSOut.db
                dataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                                    Tables.FVS.DefaultFVSOutDbFile + "' AS FVS";
                dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);

                if (dataMgr.TableExist(conn, strPrePostSeqNumMatrixTable))
                {
                    // Delete current variant/package from table
                    dataMgr.m_strSQL = $@"DELETE FROM {strPrePostSeqNumMatrixTable} WHERE FVS_VARIANT = '{strVariant}' AND RXPACKAGE = '{strRxPackage}'";
                    dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                }

                if (dataMgr.TableExist(conn, strAuditPrePostSeqNumCountsTable))
                {
                    // Delete current variant/package from table
                    dataMgr.m_strSQL = $@"DELETE FROM {strAuditPrePostSeqNumCountsTable} WHERE FVS_VARIANT = '{strVariant}' AND RXPACKAGE = '{strRxPackage}'";
                    dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                }

                if (p_strSourceTableName == "FVS_SUMMARY")
                {
                    if (dataMgr.TableExist(conn, strAuditYearCountsTable))
                    {
                        // Delete current variant/package from table
                        dataMgr.m_strSQL = $@"DELETE FROM {strAuditYearCountsTable} WHERE FVS_VARIANT = '{strVariant}' AND RXPACKAGE = '{strRxPackage}'";
                        dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                    }
                    else
                    {
                        // Create the table
                        dataMgr.m_strSQL = Tables.FVS.Audit.Pre.CreateFVSPreAuditCountsTableSQL(strAuditYearCountsTable);
                        dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                    }
                }

                if (dataMgr.TableExist(conn, "temp_rowcount"))
                    dataMgr.SqlNonQuery(conn, "DROP TABLE temp_rowcount");

                //
                //DEFAULT CONFIGURATIONS
                //
                if (p_strSourceTableName.Trim().ToUpper() == "FVS_SUMMARY" ||
                    p_strSourceTableName.Trim().ToUpper() == "FVS_CUTLIST" ||
                    p_strSourceTableName.Trim().ToUpper().IndexOf("FVS_POTFIRE", 0) >= 0)
                {
                    if (!dataMgr.TableExist(conn, strPrePostSeqNumMatrixTable))
                    {
                        frmMain.g_oTables.m_oFvs.CreateSQLiteFVSOutputPrePostSeqNumAuditGenericTable(dataMgr, conn, strPrePostSeqNumMatrixTable);
                    }
                    if (p_strSourceTableName.Trim().ToUpper() == "FVS_SUMMARY" ||
                        p_strSourceTableName.Trim().ToUpper().IndexOf("FVS_POTFIRE", 0) >= 0)
                    {
                        //STANDID + YEAR = ONE RECORD
                        dataMgr.m_strSQL = Queries.FVS.SqliteFVSOutputTable_AuditPrePostGenericSQL("", p_strTempTable, false);
                    }
                    else
                    {
                        //STANDID + YEAR = MULTIPLE RECORDS
                        dataMgr.m_strSQL = Queries.FVS.SqliteFVSOutputTable_AuditPrePostGenericSQL("", p_strTempTable, false);
                    }
                    dataMgr.m_strSQL = "INSERT INTO " + strPrePostSeqNumMatrixTable + " " +
                                      dataMgr.m_strSQL;

                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                    dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");

                    dataMgr.m_strSQL = Queries.FVS.SqliteFVSOutputTable_AuditUpdatePrePostGenericSQL(
                        p_oItem, strPrePostSeqNumMatrixTable, strVariant, strRxPackage);

                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                    dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");

                    if (p_bAudit)
                    {
                        if (p_strSourceTableName == "FVS_SUMMARY")
                        {
                            if (! dataMgr.TableExist(conn, strAuditPrePostSeqNumCountsTable))
                            {
                                dataMgr.m_strSQL = Tables.FVS.Audit.Pre.CreateFVSPreYearCountsTableSQL(strAuditPrePostSeqNumCountsTable);
                                if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                                dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                                if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                            }
                            dataMgr.m_strSQL = Queries.FVS.SqliteFVSOutputTable_AuditSelectIntoPrePostSeqNumCountSqlite
                                (p_oItem, strAuditPrePostSeqNumCountsTable, strPrePostSeqNumMatrixTable, p_strRunTitle);

                            if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                            dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                            if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");

                            string[] arrDefaultColumns = { "SEQNUM", "STANDID", "YEAR", "FVS_VARIANT", "RXPACKAGE" };
                            string[] arrAllColumns = dataMgr.getFieldNamesArray(conn, "SELECT * FROM " + strAuditYearCountsTable);
                            System.Collections.Generic.IList<string> lstExtraColumns =
                                new System.Collections.Generic.List<string>();
                            for (int i = 0; i < arrAllColumns.Length; i++)
                            {
                                if (!arrDefaultColumns.Contains(arrAllColumns[i].ToUpper()))
                                {
                                    lstExtraColumns.Add(arrAllColumns[i]);
                                }
                            }

                            // Build temp table from FVS_SUMMARY with only the current variant/package
                            string tmpTableName = "tmpSummary";
                            if (dataMgr.TableExist(conn, tmpTableName))
                            {
                                dataMgr.SqlNonQuery(conn, $@"DROP TABLE {tmpTableName}");
                            }
                            dataMgr.m_strSQL = $@"CREATE TABLE {tmpTableName} AS select s.standid, year, '{p_strRunTitle.Substring(7, 2)}' as fvs_variant, '{p_strRunTitle.Substring(11, 3)}' as rxpackage from 
                                FVS_SUMMARY s, {Tables.FVS.DefaultFVSCasesTempTableName} c where s.CaseID = c.CaseID";
                            if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                            dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                            if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");

                            dataMgr.m_strSQL = Queries.FVS.FVSOutputTable_PrePostGenericSQLite(
                               strAuditYearCountsTable, tmpTableName, false, p_strRunTitle, lstExtraColumns);
                            if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                            dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                            if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");

                            // Delete FVS_SUMMARY temp table
                            if (dataMgr.TableExist(conn, tmpTableName))
                            {
                                dataMgr.SqlNonQuery(conn, $@"DROP TABLE {tmpTableName}");
                            }
                        }
                        else
                        {
                            if (p_oItem.UseSummaryTableSeqNumYN == "N")
                            {
                                if (!dataMgr.TableExist(conn, strAuditPrePostSeqNumCountsTable))
                                {
                                    dataMgr.m_strSQL = Tables.FVS.Audit.Pre.CreateFVSPreYearCountsTableSQL(strAuditPrePostSeqNumCountsTable);
                                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                        frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                                    dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                        frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                                }

                                dataMgr.m_strSQL = Queries.FVS.SqliteFVSOutputTable_AuditSelectIntoPrePostSeqNumCountSqlite
                                  (p_oItem, strAuditPrePostSeqNumCountsTable, strPrePostSeqNumMatrixTable, p_strRunTitle);
                                if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                                dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                                if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                            }
                        }
                    }
                    if (p_strSourceTableName == "FVS_SUMMARY") return;
                }

                if (p_bAudit)
                {
                    // Create extract of p_strSourceTableName to improve performance
                    const string TMPFVSOUT = "tmpFvsOut";
                    if (dataMgr.TableExist(conn, TMPFVSOUT))
                    {
                        dataMgr.SqlNonQuery(conn, $@"drop table {TMPFVSOUT}");
                    }
                    dataMgr.m_strSQL = $@"CREATE TABLE {TMPFVSOUT} AS SELECT * FROM {p_strSourceTableName} s, {Tables.FVS.DefaultFVSCasesTempTableName} c
                                        where s.CaseID = c.CaseID AND RunTitle = '{p_strRunTitle}'";
                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                    dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");

                    //
                    //stand + year rowcounts for all tables as compared with stand + year in the summary table
                    //
                    if (!dataMgr.ColumnExists(conn, "audit_fvs_summary_year_counts_table", p_strSourceTableName))
                    {
                        dataMgr.AddColumn(conn, "audit_fvs_summary_year_counts_table", p_strSourceTableName, "INTEGER", "");
                    }
                    string[] strSQL = Queries.FVS.FVSOutputTable_AuditFVSSummaryTableRowCountsSQL(
                        "temp_rowcount",
                        "audit_FVS_SUMMARY_year_counts_table",
                        TMPFVSOUT,
                        p_strSourceTableName, p_strRunTitle);

                    for (x = 0; x <= strSQL.Length - 1; x++)
                    {
                        if (p_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL[x] + "\r\n");
                        dataMgr.SqlNonQuery(conn, strSQL[x]);
                        if (p_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                    }

                    // Clean up tmp table
                    if (dataMgr.TableExist(conn, TMPFVSOUT))
                    {
                        if (p_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + $@"drop table {TMPFVSOUT}" + "\r\n");
                        dataMgr.SqlNonQuery(conn, $@"drop table {TMPFVSOUT}");
                        if (p_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");

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
                    if (!dataMgr.TableExist(conn, strPrePostSeqNumMatrixTable))
                    {
                        frmMain.g_oTables.m_oFvs.CreateSQLiteFVSOutputPrePostSeqNumAuditStrClassTable(dataMgr, conn, strPrePostSeqNumMatrixTable);
                    }
                }
                else
                {
                    if (!dataMgr.TableExist(conn, strPrePostSeqNumMatrixTable))
                    {
                        frmMain.g_oTables.m_oFvs.CreateSQLiteFVSOutputPrePostSeqNumAuditGenericTable(dataMgr, conn, strPrePostSeqNumMatrixTable);
                    }
                }

                if (p_oItem.TableName.Trim().ToUpper() == "FVS_STRCLASS")
                {
                    if (p_oItem.UseSummaryTableSeqNumYN == "N")
                    {
                        dataMgr.m_strSQL = Queries.FVS.SqliteFVSOutputTable_AuditPrePostStrClassSQL("", p_strTempTable, false);
                    }
                    else
                    {
                        strSQLArray = Queries.FVS.SqliteFVSOutputTable_AuditPrePostFvsStrClassUsingFVSSummarySQL("", p_strTempTable, false, p_strRunTitle);
                    }
                }
                else
                {
                    dataMgr.m_strSQL = Queries.FVS.SqliteFVSOutputTable_AuditPrePostGenericSQL("", p_strTempTable, false);
                }

                if (p_oItem.TableName.Trim().ToUpper() == "FVS_STRCLASS")
                {
                    if (p_oItem.UseSummaryTableSeqNumYN == "Y")
                    {
                        for (x = 0; x <= strSQLArray.Length - 1; x++)
                        {
                            dataMgr.m_strSQL = "INSERT INTO " + strPrePostSeqNumMatrixTable + " " +
                                                 strSQLArray[x];
                            if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                            dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                            if (p_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                        }
                    }
                    else
                    {
                        dataMgr.m_strSQL = "INSERT INTO " + strPrePostSeqNumMatrixTable + " " +
                                           dataMgr.m_strSQL;
                        if (p_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                        dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                        if (p_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                    }

                    dataMgr.m_strSQL = Queries.FVS.SqliteFVSOutputTable_AuditUpdatePrePostStrClassSQL(
                       p_oItem, strPrePostSeqNumMatrixTable, strVariant, strRxPackage);

                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                    dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                }
                else
                {
                    dataMgr.m_strSQL = "INSERT INTO " + strPrePostSeqNumMatrixTable + " " +
                                            dataMgr.m_strSQL;

                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                    dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");

                    dataMgr.m_strSQL = Queries.FVS.SqliteFVSOutputTable_AuditUpdatePrePostGenericSQL(p_oItem, strPrePostSeqNumMatrixTable,
                        strVariant, strRxPackage);
                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                    dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                }


                if (p_oItem.UseSummaryTableSeqNumYN == "N")
                {
                    if (!dataMgr.TableExist(conn, strAuditPrePostSeqNumCountsTable))
                    {
                        dataMgr.m_strSQL = Tables.FVS.Audit.Pre.CreateFVSPreYearCountsTableSQL(strAuditPrePostSeqNumCountsTable);
                        if (p_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                        dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                        if (p_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                    }
                    dataMgr.m_strSQL = Queries.FVS.SqliteFVSOutputTable_AuditSelectIntoPrePostSeqNumCountSqlite
                      (p_oItem, strAuditPrePostSeqNumCountsTable, strPrePostSeqNumMatrixTable, p_strRunTitle);

                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + dataMgr.m_strSQL + "\r\n");
                    dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
                    if (p_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(p_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                }
            }
        }

        public void CreateFvsOutDbIndexes(string strFvsOutDbPath)
        {
            string idxSummary = "IDX_Summary";  //FVS-created index on fvs_summary.CaseId
            string idxCases = "IDX_Cases";  //FVS-created index on fvs_cases.CaseId
            string idxCutlist = "IDX_CutList";  //FVS-created index on fvs_CutList.CaseId
            string idxSummaryStandId = "IDX_Summary_StandId";    //Biosum-created index on fvs_summary.StandId
            string idxCasesRunTitle = "IDX_Cases_RunTitle"; //BioSum-created index on fvs_summary.StandId
            string idxCutListComposite = "IDX_CutList_Composite"; //BioSum-created index on fvs_cutTree.CaseID,StandId,Year,TreeId
            // Note: We also set some indexes in AppendRuntitleToFVSOut for FVSOut_BioSum.db
            DataMgr oDataMgr = new DataMgr();
            if (System.IO.File.Exists(strFvsOutDbPath))
            {
                string dbConn = oDataMgr.GetConnectionString(strFvsOutDbPath);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dbConn))
                {
                    conn.Open();
                    if (oDataMgr.TableExist(conn, Tables.FVS.DefaultFVSCasesTableName))
                    {
                        if (!oDataMgr.IndexExist(conn, idxCasesRunTitle))
                        {
                            oDataMgr.AddIndex(conn, Tables.FVS.DefaultFVSCasesTableName, idxCasesRunTitle, "RunTitle");
                        }
                        if (strFvsOutDbPath.ToUpper().IndexOf("BIOSUM") > 0 &&
                            (!oDataMgr.IndexExist(conn, idxCases)))
                        {
                            // Replicate index that was on original FVSOut.db
                            oDataMgr.AddIndex(conn, Tables.FVS.DefaultFVSCasesTableName, idxCases, "CaseId");
                        }
                    }
                    if (oDataMgr.TableExist(conn, Tables.FVS.DefaultFVSSummaryTableName))
                    {
                        if (!oDataMgr.IndexExist(conn, idxSummaryStandId))
                        {
                            oDataMgr.AddIndex(conn, Tables.FVS.DefaultFVSSummaryTableName, idxSummaryStandId, "StandId");
                        }
                        if (strFvsOutDbPath.ToUpper().IndexOf("BIOSUM") > 0 &&
                            (!oDataMgr.IndexExist(conn, idxSummary)))
                        {
                            // Replicate index that was on original FVSOut.db
                            oDataMgr.AddIndex(conn, Tables.FVS.DefaultFVSSummaryTableName, idxSummary, "CaseId");
                        }
                    }
                    if (oDataMgr.TableExist(conn, Tables.FVS.DefaultFVSCutListTableName))
                    {
                        if (!oDataMgr.IndexExist(conn, idxCutListComposite))
                        {
                            oDataMgr.AddIndex(conn, Tables.FVS.DefaultFVSCutListTableName, idxCutListComposite, "CaseID,StandId,Year,TreeId");
                        }
                        if (strFvsOutDbPath.ToUpper().IndexOf("BIOSUM") > 0 &&
                            (!oDataMgr.IndexExist(conn, idxCutlist)))
                        {
                            // Replicate index that was on original FVSOut.db
                            oDataMgr.AddIndex(conn, Tables.FVS.DefaultFVSCutListTableName, idxCutlist, "CaseId");
                        }
                    }
                }
            }
        }

        public System.Collections.Generic.IDictionary<string, RxPackageItem_Collection> GetFvsVariantPackageDictionary(Queries oQueries)
        {
            System.Collections.Generic.IDictionary<string, RxPackageItem_Collection> dictReturn =
                new System.Collections.Generic.Dictionary<string, RxPackageItem_Collection>();

            DataMgr oDataMgr = new DataMgr();
            string strConn = oDataMgr.GetConnectionString(oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Plot));
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                conn.Open();
                oDataMgr.m_strSQL = Queries.FVS.GetFVSVariantRxPackageSQL(oQueries.m_oFIAPlot.m_strPlotTable, oQueries.m_oFvs.m_strRxPackageTable);
                oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
                string strVariant;
                string strRxPackage;
                int intIndex = 0;
                while (oDataMgr.m_DataReader.Read())
                {
                    strVariant = oDataMgr.m_DataReader["fvs_variant"].ToString().Trim();
                    strRxPackage = oDataMgr.m_DataReader["rxPackage"].ToString().Trim();
                    if (!string.IsNullOrEmpty(strRxPackage))
                    {
                        if (oDataMgr.m_DataReader["rxpackage"] != System.DBNull.Value)
                        {
                            RxPackageItem oRxPackageItem = new RxPackageItem();
                            oRxPackageItem.Index = intIndex;
                            oRxPackageItem.RxPackageId = Convert.ToString(oDataMgr.m_DataReader["rxpackage"]);
                            //treatment cycle length
                            if (oDataMgr.m_DataReader["rxcycle_length"] != System.DBNull.Value)
                            {
                                oRxPackageItem.RxCycleLength = Convert.ToInt32(oDataMgr.m_DataReader["rxcycle_length"]);
                            }
                            //sim year 1
                            if (oDataMgr.m_DataReader["simyear1_rx"] != System.DBNull.Value)
                            {
                                oRxPackageItem.SimulationYear1Rx = Convert.ToString(oDataMgr.m_DataReader["simyear1_rx"]);
                            }
                            //sim year 2
                            if (oDataMgr.m_DataReader["simyear2_rx"] != System.DBNull.Value)
                            {
                                oRxPackageItem.SimulationYear2Rx = Convert.ToString(oDataMgr.m_DataReader["simyear2_rx"]);
                            }                            //sim year 3
                            if (oDataMgr.m_DataReader["simyear3_rx"] != System.DBNull.Value)
                            {
                                oRxPackageItem.SimulationYear3Rx = Convert.ToString(oDataMgr.m_DataReader["simyear3_rx"]);
                            }
                            //sim year 4
                            if (oDataMgr.m_DataReader["simyear4_rx"] != System.DBNull.Value)
                            {
                                oRxPackageItem.SimulationYear4Rx = Convert.ToString(oDataMgr.m_DataReader["simyear4_rx"]);
                            }
                            if (oRxPackageItem.SimulationYear1Rx.Trim().Length == 0) oRxPackageItem.SimulationYear1Rx = "000";
                            if (oRxPackageItem.SimulationYear2Rx.Trim().Length == 0) oRxPackageItem.SimulationYear2Rx = "000";
                            if (oRxPackageItem.SimulationYear3Rx.Trim().Length == 0) oRxPackageItem.SimulationYear3Rx = "000";
                            if (oRxPackageItem.SimulationYear4Rx.Trim().Length == 0) oRxPackageItem.SimulationYear4Rx = "000";

                            RxPackageItem_Collection oRxPackageItemCollection = null;
                            if (dictReturn.ContainsKey(strVariant))
                            {
                                oRxPackageItemCollection = dictReturn[strVariant];
                                oRxPackageItemCollection.Add(oRxPackageItem);
                            }
                            else
                            {
                                oRxPackageItemCollection = new RxPackageItem_Collection();
                                oRxPackageItemCollection.Add(oRxPackageItem);
                                dictReturn.Add(strVariant, oRxPackageItemCollection);
                            }
                            intIndex++;
                        }
                    }
                    
                }
                oDataMgr.m_DataReader.Close();
            }
            return dictReturn;
        }
    }
	/*********************************************************************************************************
	 **RX Item                          
	 *********************************************************************************************************/
	public class RxItem
	{
		private int _intIndex;
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
			p_oRxItemDestination.Description="";
			p_oRxItemDestination.Index=-1;
			p_oRxItemDestination.RxId="";
			p_oRxItemDestination.RxPackageMemberList="";
			p_oRxItemDestination.HarvestMethodLowSlope="";
			p_oRxItemDestination.HarvestMethodSteepSlope="";			
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
			p_oRxItemDestination.FvsCycleList=p_oRxItemSource.FvsCycleList;
            p_oRxItemDestination.HarvestMethodLowSlope = p_oRxItemSource.HarvestMethodLowSlope;
			p_oRxItemDestination.HarvestMethodSteepSlope=p_oRxItemSource.HarvestMethodSteepSlope;
			p_oRxItemDestination.Delete = p_oRxItemSource.Delete;
			p_oRxItemDestination.Add = p_oRxItemSource.Add;

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

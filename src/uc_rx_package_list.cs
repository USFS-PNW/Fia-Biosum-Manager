using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_rx_package_list.
	/// </summary>
	public class uc_rx_package_list : System.Windows.Forms.UserControl
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

		private Queries m_oQueries; 
		string m_strTable="";		
		private string m_strConn;
		private System.Windows.Forms.Button btnDelete;
		private ListViewAlternateBackgroundColors m_oLvAlternateColors = new FIA_Biosum_Manager.ListViewAlternateBackgroundColors();
		const int COLUMN_NULL=0;
		const int COLUMN_RXPACKAGE=1;
		//const int COLUMN_RXCYCLELENGTH=2;
		const int COLUMN_SIMYR1_RX=2;
		//const int COLUMN_SIMYR1_FVS=4;
		const int COLUMN_SIMYR2_RX=3;
		//const int COLUMN_SIMYR2_FVS=6;
		const int COLUMN_SIMYR3_RX=4;
		//const int COLUMN_SIMYR3_FVS=8;
		const int COLUMN_SIMYR4_RX=5;
		//const int COLUMN_SIMYR4_FVS=10;
		const int COLUMN_DESC=6;

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultFvsXPSFile;

		private FIA_Biosum_Manager.RxPackageItem_Collection m_oRxPackageItem_Collection = new RxPackageItem_Collection();
		private System.Windows.Forms.Button btnProperties;

		//private RxPackageItemFvsCommandItem_Collection  _FvsCommandItem_Collection1;

		private FIA_Biosum_Manager.RxItem_Collection m_oRxItem_Collection = new RxItem_Collection();
		
		private RxTools m_oRxTools = new RxTools();
		private frmMain _frmMain=null;
		private frmDialog _frmDialog=null;

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_rx_package_list()
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

			this.m_oQueries = new Queries();
			m_oQueries.m_oFvs.LoadDatasource=true;
			m_oQueries.LoadDatasources(true);
			this.m_oLvAlternateColors.InitializeRowCollection(); 
     
			this.lstRx.Clear();
			this.lstRx.Columns.Add(" ",2,HorizontalAlignment.Left);
			this.lstRx.Columns.Add("RxPackage", 100, HorizontalAlignment.Left);

			this.lstRx.Columns.Add("SimCycle1_rx", 110, HorizontalAlignment.Left);

			this.lstRx.Columns.Add("SimCycle2_rx",110, HorizontalAlignment.Left);

			this.lstRx.Columns.Add("SimCycle3_rx", 110, HorizontalAlignment.Left);

			this.lstRx.Columns.Add("SimCycle4_rx", 110, HorizontalAlignment.Left);

			this.lstRx.Columns.Add("Description", 200, HorizontalAlignment.Left);
			
			this.m_intError=0;

			m_oRxItem_Collection.Clear();
			this.m_oRxPackageItem_Collection.Clear();			
            this.m_oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(m_oQueries, m_oRxPackageItem_Collection);
            this.m_oRxTools.LoadAllRxItemsFromTableIntoRxCollection(m_oQueries, m_oRxItem_Collection);
            
			this.lstRx.BeginUpdate();    	
			for (x=0;x<= m_oRxPackageItem_Collection.Count-1;x++)
			{
				
				lstRx.Items.Add("");
				lstRx.Items[x].UseItemStyleForSubItems=false;
				lstRx.Items[x].SubItems.Add(m_oRxPackageItem_Collection.Item(x).RxPackageId);
				lstRx.Items[x].SubItems.Add(m_oRxPackageItem_Collection.Item(x).SimulationYear1Rx);
				lstRx.Items[x].SubItems.Add(m_oRxPackageItem_Collection.Item(x).SimulationYear2Rx);
				lstRx.Items[x].SubItems.Add(m_oRxPackageItem_Collection.Item(x).SimulationYear3Rx);
				lstRx.Items[x].SubItems.Add(m_oRxPackageItem_Collection.Item(x).SimulationYear4Rx);
				lstRx.Items[x].SubItems.Add(m_oRxPackageItem_Collection.Item(x).Description);
				m_oLvAlternateColors.AddRow();
				m_oLvAlternateColors.AddColumns(x,lstRx.Columns.Count);

			}
			this.lstRx.EndUpdate();

			
			this.m_oLvAlternateColors.ListView();
			
		
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
            this.btnDefault.Location = new System.Drawing.Point(32, 392);
            this.btnDefault.Name = "btnDefault";
            this.btnDefault.Size = new System.Drawing.Size(114, 32);
            this.btnDefault.TabIndex = 2;
            this.btnDefault.Text = "Use Default Values";
            this.btnDefault.Visible = false;
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
            this.lblTitle.Text = "Treatment Package List";
            // 
            // uc_rx_package_list
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_rx_package_list";
            this.Size = new System.Drawing.Size(672, 480);
            this.Resize += new System.EventHandler(this.uc_rx_package_list_Resize);
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


			

			FIA_Biosum_Manager.frmRxPackageItem frmRxPackageItem1 = new frmRxPackageItem();
			frmRxPackageItem1.MaximizeBox = true;
			frmRxPackageItem1.BackColor = System.Drawing.SystemColors.Control;
			frmRxPackageItem1.Text = "FVS: Treatment Package Item (New)";



					

			frmRxPackageItem1.uc_rx_package_edit1.m_oResizeForm.ControlToResize = frmRxPackageItem1;
			
			
			frmRxPackageItem1.uc_rx_package_edit1.m_oResizeForm.ResizeControl();

			frmRxPackageItem1.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;


			
			RxPackageItem oRxPackageItem = new RxPackageItem();
		    frmRxPackageItem1.ReferenceRxPackageItem = oRxPackageItem;
            
			frmRxPackageItem1.UsedRxPackageList=this.m_oRxTools.GetUsedRxPackageList(m_oRxPackageItem_Collection);
			
			frmRxPackageItem1.m_strAction="new";
			frmRxPackageItem1.ReferenceRxItemCollection = this.m_oRxItem_Collection;
			frmRxPackageItem1.loadvalues();

            frmRxPackageItem1.ReferenceUserControlRxPackageList = this;
            frmRxPackageItem1.ParentControl = (frmDialog)ParentForm;
            frmRxPackageItem1.ParentControl.Enabled = false;
            frmRxPackageItem1.MinimizeMainForm = true;
            frmRxPackageItem1.Show();
		    

		}
		public void AddItem(RxPackageItem p_oRxPackageItem)
		{
			
			this.lstRx.Items.Add("");
			lstRx.Items[lstRx.Items.Count-1].UseItemStyleForSubItems=false;
			
			lstRx.Items[lstRx.Items.Count-1].SubItems.Add(p_oRxPackageItem.RxPackageId);
			lstRx.Items[lstRx.Items.Count-1].SubItems.Add(p_oRxPackageItem.SimulationYear1Rx);
			lstRx.Items[lstRx.Items.Count-1].SubItems.Add(p_oRxPackageItem.SimulationYear2Rx);
			lstRx.Items[lstRx.Items.Count-1].SubItems.Add(p_oRxPackageItem.SimulationYear3Rx);
			lstRx.Items[lstRx.Items.Count-1].SubItems.Add(p_oRxPackageItem.SimulationYear4Rx);
			lstRx.Items[lstRx.Items.Count-1].SubItems.Add(p_oRxPackageItem.Description);
						
			this.m_oLvAlternateColors.AddRow();
			this.m_oLvAlternateColors.AddColumns(lstRx.Items.Count-1,lstRx.Columns.Count);

			this.m_oRxPackageItem_Collection.Add(p_oRxPackageItem);

			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
		}
		public void UpdateItem(RxPackageItem p_oRxPackageItem, bool bEnableSave)
		{
			if (this.lstRx.SelectedItems.Count == 0) return;

			int x = this.lstRx.SelectedItems[0].Index;
			this.lstRx.BeginUpdate();
			
            if (p_oRxPackageItem.SimulationYear1Rx.Trim().Length > 0)
			    this.lstRx.Items[x].SubItems[COLUMN_SIMYR1_RX].Text = p_oRxPackageItem.SimulationYear1Rx;
            else
                this.lstRx.Items[x].SubItems[COLUMN_SIMYR1_RX].Text = "000";
            if (p_oRxPackageItem.SimulationYear2Rx.Trim().Length > 0)
                this.lstRx.Items[x].SubItems[COLUMN_SIMYR2_RX].Text = p_oRxPackageItem.SimulationYear2Rx;
            else
                this.lstRx.Items[x].SubItems[COLUMN_SIMYR2_RX].Text = "000";
            if (p_oRxPackageItem.SimulationYear3Rx.Trim().Length > 0)
                this.lstRx.Items[x].SubItems[COLUMN_SIMYR3_RX].Text = p_oRxPackageItem.SimulationYear3Rx;
            else
                this.lstRx.Items[x].SubItems[COLUMN_SIMYR3_RX].Text = "000";
            if (p_oRxPackageItem.SimulationYear4Rx.Trim().Length > 0)
                this.lstRx.Items[x].SubItems[COLUMN_SIMYR4_RX].Text = p_oRxPackageItem.SimulationYear4Rx;
            else
                this.lstRx.Items[x].SubItems[COLUMN_SIMYR4_RX].Text = "000";
			
			this.lstRx.Items[x].SubItems[COLUMN_DESC].Text = p_oRxPackageItem.Description;
			this.lstRx.EndUpdate();

            if (bEnableSave == true) this.btnSave.Enabled = true;
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
				this.m_oRxPackageItem_Collection.Item(x).Delete=true;
				this.m_oRxPackageItem_Collection.Item(x).Index=-1;
			}
		}

		private void btnEdit_Click(object sender, System.EventArgs e)
		{

			string strDesc="";
			string strRxList = "";
			int x;

			if (this.lstRx.SelectedItems.Count == 0)
				return;

			
			FIA_Biosum_Manager.frmRxPackageItem frmRxPackageItem1 = new frmRxPackageItem();
			frmRxPackageItem1.MaximizeBox = true;
			frmRxPackageItem1.BackColor = System.Drawing.SystemColors.Control;
			frmRxPackageItem1.Text = "FVS: Treatment Package Item (Edit)";
					

			frmRxPackageItem1.uc_rx_package_edit1.m_oResizeForm.ControlToResize = frmRxPackageItem1;
			
			
			frmRxPackageItem1.uc_rx_package_edit1.m_oResizeForm.ResizeControl();

			frmRxPackageItem1.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			

			//find the current rxid
			for (x=0;x<=this.m_oRxPackageItem_Collection.Count-1;x++)
			{
				if (this.m_oRxPackageItem_Collection.Item(x).RxPackageId.Trim() == 
					this.lstRx.SelectedItems[0].SubItems[COLUMN_RXPACKAGE].Text.Trim())
				{
					frmRxPackageItem1.ReferenceRxPackageItem = this.m_oRxPackageItem_Collection.Item(x);
					break;
					
				}
			}
			
			frmRxPackageItem1.ReferenceRxItemCollection = this.m_oRxItem_Collection;
			frmRxPackageItem1.m_strAction="edit";
			frmRxPackageItem1.loadvalues();

            frmRxPackageItem1.ReferenceUserControlRxPackageList = this;
            frmRxPackageItem1.ParentControl = (frmDialog)ParentForm;
            frmRxPackageItem1.ParentControl.Enabled = false;
            frmRxPackageItem1.MinimizeMainForm = true;
            frmRxPackageItem1.Show();


		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
		  this.savevalues();
		}
		private void val_data()
		{

		}
		public void savevalues()
		{
			int x;
			string strFields;
			string strValues;			

            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            string strSource = m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.RxPackage);
            if (!string.IsNullOrEmpty(strSource))
            {
                string strConn = oDataMgr.GetConnectionString(strSource);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    oDataMgr.m_strSQL = "DELETE FROM " + this.m_oQueries.m_oFvs.m_strRxPackageTable;
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    if (oDataMgr.m_intError == 0)
                    {
                        for (x = 0; x <= m_oRxPackageItem_Collection.Count - 1; x++)
                        {
                            if (m_oRxPackageItem_Collection.Item(x).Delete == false)
                            {
                                //
                                //SAVE PACKAGES
                                //
                                strFields = "";
                                strValues = "";

                                strFields = "rxpackage,description,rxcycle_length," +
                                    "simyear1_rx,simyear1_fvscycle," +
                                    "simyear2_rx,simyear2_fvscycle," +
                                    "simyear3_rx,simyear3_fvscycle," +
                                    "simyear4_rx,simyear4_fvscycle";
                                strValues = "'" + this.m_oRxPackageItem_Collection.Item(x).RxPackageId.Trim() + "',";
                                strValues = strValues + "'" + oDataMgr.FixString(m_oRxPackageItem_Collection.Item(x).Description.Trim(), "'", "''") + "',";
                                strValues = strValues + "'" + Convert.ToString(m_oRxPackageItem_Collection.Item(x).RxCycleLength).Trim() + "',";
                                strValues = strValues + "'" + m_oRxPackageItem_Collection.Item(x).SimulationYear1Rx.Trim() + "',";
                                strValues = strValues + "'" + m_oRxPackageItem_Collection.Item(x).SimulationYear1Fvs.Trim() + "',";
                                strValues = strValues + "'" + m_oRxPackageItem_Collection.Item(x).SimulationYear2Rx.Trim() + "',";
                                strValues = strValues + "'" + m_oRxPackageItem_Collection.Item(x).SimulationYear2Fvs.Trim() + "',";
                                strValues = strValues + "'" + m_oRxPackageItem_Collection.Item(x).SimulationYear3Rx.Trim() + "',";
                                strValues = strValues + "'" + m_oRxPackageItem_Collection.Item(x).SimulationYear3Fvs.Trim() + "',";
                                strValues = strValues + "'" + m_oRxPackageItem_Collection.Item(x).SimulationYear4Rx.Trim() + "',";
                                strValues = strValues + "'" + m_oRxPackageItem_Collection.Item(x).SimulationYear4Fvs.Trim() + "'";
                                oDataMgr.m_strSQL = Queries.GetInsertSQL(strFields, strValues, m_oQueries.m_oFvs.m_strRxPackageTable);
                                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                            }
                        }
                    }

                    if (this.m_intError == 0 && oDataMgr.m_intError == 0)
                    {
                        this.btnSave.Enabled = false;
                    }
                }
            }
		}


		private void btnDelete_Click(object sender, System.EventArgs e)
		{
			if (this.lstRx.SelectedItems.Count==0) return;
			int x;
			for (x=0;x<=m_oRxPackageItem_Collection.Count-1;x++)
			{
				if (m_oRxPackageItem_Collection.Item(x).RxPackageId.Trim()==
					lstRx.SelectedItems[0].SubItems[COLUMN_RXPACKAGE].Text.Trim())
				{
					m_oRxPackageItem_Collection.Item(x).Delete=true;
					m_oRxPackageItem_Collection.Item(x).Index = -1;
					
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

		private void uc_rx_package_list_Resize(object sender, System.EventArgs e)
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
		
	

		private void btnProperties_Click(object sender, System.EventArgs e)
		{
			if  (this.lstRx.Items.Count==0) return;
			frmMain.g_sbpInfo.Text = "Creating Rx Package Properties Report...Stand By";
			FIA_Biosum_Manager.RxItem_Collection oRxColl = new RxItem_Collection();

			for (int x=0;x<=m_oRxItem_Collection.Count-1;x++)
			{
				if (m_oRxItem_Collection.Item(x).Delete==false)
				{
					RxItem oItem = new RxItem();
					oItem.CopyProperties(m_oRxItem_Collection.Item(x),oItem);
					oRxColl.Add(oItem);

				}
			}

            FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
            frmTemp.Text = "FIA Biosum";
            frmTemp.AutoScroll = false;
            uc_textbox uc_textbox1 = new uc_textbox();
            frmTemp.Controls.Add(uc_textbox1);
            uc_textbox1.Dock = DockStyle.Fill;
            uc_textbox1.lblTitle.Text = "Package Properties";
            RxTools oRxTools = new RxTools();
            uc_textbox1.TextValue = oRxTools.PackageProperties(this.m_oRxPackageItem_Collection, oRxColl);
            frmTemp.Show();


			frmMain.g_sbpInfo.Text = "Ready";


		
		}

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "FVS", "RX_PACKAGE_TREATMENT_LIST" });
        }
		
		
		
	}
	/*********************************************************************************************************
	 **RX Package Item                          
	 *********************************************************************************************************/
	public class RxPackageItem
	{				
		private int _intIndex=-1;
		[CategoryAttribute("General"),ReadOnly(true),DescriptionAttribute("RX Package Item Index")]
		public int Index
		{
			get {return _intIndex;}
			set {_intIndex = value;}
		}
		private int _intSaveIndex=-1;
		[CategoryAttribute("General"),ReadOnly(true),DescriptionAttribute("RX Package Item Index")]
		public int SaveIndex
		{
			get {return _intSaveIndex;}
			set {_intSaveIndex = value;}
		}

		private string _strRxPackageId="";
		[CategoryAttribute("General"),DescriptionAttribute("RX Package Indentifier")]
		public string RxPackageId
		{
			get {return _strRxPackageId;}
			set {_strRxPackageId=value;}
		}
		private string _strDesc="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string Description
		{
			get {return _strDesc;}
			set {_strDesc=value;}

		}
		private int _intRxCycleLength=10;
		public int RxCycleLength
		{
			get {return _intRxCycleLength;}
			set {_intRxCycleLength=value;}
		}
		private string _strSimYr1Rx="";
		[CategoryAttribute("General"),DescriptionAttribute("FVS Simulation Year Rx")]
		public string SimulationYear1Rx
		{
			get {return _strSimYr1Rx;}
			set {_strSimYr1Rx=value;}

		}
		private string _strSimYr1Fvs="";
		[CategoryAttribute("General"),DescriptionAttribute("FVS Simulation Year")]
		public string SimulationYear1Fvs
		{
			get {return _strSimYr1Fvs;}
			set {_strSimYr1Fvs=value;}

		}
		private string _strSimYr2Rx="";
		[CategoryAttribute("General"),DescriptionAttribute("FVS Simulation Year Rx")]
		public string SimulationYear2Rx
		{
			get {return _strSimYr2Rx;}
			set {_strSimYr2Rx=value;}

		}
		private string _strSimYr2Fvs="";
		[CategoryAttribute("General"),DescriptionAttribute("FVS Simulation Year")]
		public string SimulationYear2Fvs
		{
			get {return _strSimYr2Fvs;}
			set {_strSimYr2Fvs=value;}
		}
		private string _strSimYr3Rx="";
		[CategoryAttribute("General"),DescriptionAttribute("FVS Simulation Year Rx")]
		public string SimulationYear3Rx
		{
			get {return _strSimYr3Rx;}
			set {_strSimYr3Rx=value;}

		}
		private string _strSimYr3Fvs="";
		[CategoryAttribute("General"),DescriptionAttribute("FVS Simulation Year")]
		public string SimulationYear3Fvs
		{
			get {return _strSimYr3Fvs;}
			set {_strSimYr3Fvs=value;}

		}
		private string _strSimYr4Rx="";
		[CategoryAttribute("General"),DescriptionAttribute("FVS Simulation Year Rx")]
		public string SimulationYear4Rx
		{
			get {return _strSimYr4Rx;}
			set {_strSimYr4Rx=value;}

		}
		private string _strSimYr4Fvs="";
		[CategoryAttribute("General"),DescriptionAttribute("FVS Simulation Year")]
		public string SimulationYear4Fvs
		{
			get {return _strSimYr4Fvs;}
			set {_strSimYr4Fvs=value;}

		}
		private string _strRxMemberList="";
		[CategoryAttribute("General"),DescriptionAttribute("Description")]
		public string RxMemberList
		{
			get {return _strRxMemberList;}
			set {_strRxMemberList=value;}

		}

        [CategoryAttribute("General"), DescriptionAttribute("FVS RunTitle")]
        public string RunTitleSuffix
        {
            // The key is the RunTitle: FVSOUT_SO_P001-001-999-999-999
            get { return $@"_P{RxPackageId}-{SimulationYear1Rx.Trim()}-{SimulationYear2Rx.Trim()}-{SimulationYear3Rx.Trim()}-{SimulationYear4Rx.Trim()}"; }
        }
		
		bool _bDelete=false;
		public bool Delete
		{
			get {return _bDelete;}
			set {_bDelete=value;}
		}
			
		public void CopyProperties(FIA_Biosum_Manager.RxPackageItem p_RxPackageItemSource,FIA_Biosum_Manager.RxPackageItem  p_RxPackageItemDestination)
		{
			p_RxPackageItemDestination.Description="";
			p_RxPackageItemDestination.Index=-1;
			p_RxPackageItemDestination.RxPackageId="";
			p_RxPackageItemDestination.SimulationYear1Rx="";
			p_RxPackageItemDestination.SimulationYear1Fvs="";
			p_RxPackageItemDestination.SimulationYear2Rx="";
			p_RxPackageItemDestination.SimulationYear2Fvs="";
			p_RxPackageItemDestination.SimulationYear3Rx="";
			p_RxPackageItemDestination.SimulationYear3Fvs="";
            p_RxPackageItemDestination.SimulationYear4Rx="";
			p_RxPackageItemDestination.SimulationYear4Fvs="";
			p_RxPackageItemDestination.RxCycleLength=-1;
			
		
			p_RxPackageItemDestination.Index  = p_RxPackageItemSource.Index;
			p_RxPackageItemDestination.Description = p_RxPackageItemSource.Description;
			p_RxPackageItemDestination.RxPackageId = p_RxPackageItemSource.RxPackageId;
			p_RxPackageItemDestination.RxCycleLength = p_RxPackageItemSource.RxCycleLength;
			p_RxPackageItemDestination.SimulationYear1Rx=p_RxPackageItemSource.SimulationYear1Rx;
			p_RxPackageItemDestination.SimulationYear1Fvs=p_RxPackageItemSource.SimulationYear1Fvs;
			p_RxPackageItemDestination.SimulationYear2Rx=p_RxPackageItemSource.SimulationYear2Rx;
			p_RxPackageItemDestination.SimulationYear2Fvs=p_RxPackageItemSource.SimulationYear2Fvs;
			p_RxPackageItemDestination.SimulationYear3Rx=p_RxPackageItemSource.SimulationYear3Rx;
			p_RxPackageItemDestination.SimulationYear3Fvs=p_RxPackageItemSource.SimulationYear3Fvs;
            p_RxPackageItemDestination.SimulationYear4Rx=p_RxPackageItemSource.SimulationYear4Rx;
			p_RxPackageItemDestination.SimulationYear4Fvs=p_RxPackageItemSource.SimulationYear4Fvs;
			p_RxPackageItemDestination.Delete=p_RxPackageItemSource.Delete;			
		}
	}
	public class RxPackageItem_Collection : System.Collections.CollectionBase
	{
		public RxPackageItem_Collection()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public void Add(FIA_Biosum_Manager.RxPackageItem m_PropertiesRxPackageItem)
		{
			// vérify if object is not already in
			if (this.List.Contains(m_PropertiesRxPackageItem))
				throw new InvalidOperationException();
 
			// adding it
			this.List.Add(m_PropertiesRxPackageItem);
 
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
		public FIA_Biosum_Manager.RxPackageItem Item(int Index)
		{
			// The appropriate item is retrieved from the List object and
			// explicitly cast to the Widget type, then returned to the 
			// caller.
			return (FIA_Biosum_Manager.RxPackageItem) List[Index];
		}
	}
		
}


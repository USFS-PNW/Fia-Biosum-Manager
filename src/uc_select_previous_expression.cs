using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_previous_expressions.
	/// </summary>
	public class uc_previous_expressions : System.Windows.Forms.UserControl
	{
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.ListView listView1;
		private System.Windows.Forms.Button btnCancel;
		public System.Windows.Forms.Label lblSQL;
		private System.Windows.Forms.Button btnSelect;
		private System.Windows.Forms.Button btnDelete;
		private System.Windows.Forms.Button btnRecall;
        public int m_intDisplayColumn;
		public int m_intSelectColumn;
		public int m_intFullHt=464;
		public int m_intFullWd=592;
		private string m_strTable;
		private string m_strCurrentConnection;
		private string[] m_strFieldTypeAString_YN;
		private int m_intCurrSelect=0;
		public FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvAlternateColors = new ListViewAlternateBackgroundColors();
		
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_previous_expressions()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			
			m_oLvAlternateColors.ReferenceAlternateBackgroundColor=frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvAlternateColors.ReferenceListView=this.listView1;
			this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvAlternateColors.CustomFullRowSelect=true;
            if (frmMain.g_oGridViewFont != null) this.listView1.Font = frmMain.g_oGridViewFont;
			if (frmMain.g_oGridViewFont != null) this.lblSQL.Font = frmMain.g_oGridViewFont;

			this.groupBox1.Left = 10;
			this.groupBox1.Width = this.lblTitle.Width - 20;
			this.listView1.Left = 10;
			this.listView1.Width = this.groupBox1.Width - 20;
			this.lblSQL.Left = this.listView1.Left;
			this.lblSQL.Width = this.listView1.Width;
			this.btnSelect.Top = this.Height - this.btnSelect.Height - 10;
			this.btnDelete.Top = this.btnSelect.Top;
			this.btnCancel.Top = this.btnSelect.Top;
			this.btnRecall.Top = this.btnSelect.Top;
			this.btnSelect.Left = (int)(this.Width * .50) -
				                   this.btnDelete.Width - this.btnSelect.Width;
			this.btnDelete.Left = this.btnSelect.Left + this.btnSelect.Width;
			this.btnRecall.Left = this.btnDelete.Left + this.btnDelete.Width;
			this.btnCancel.Left = this.btnRecall.Left + this.btnRecall.Width;
			this.groupBox1.Height = this.btnSelect.Top - this.groupBox1.Top - 5;
			this.lblSQL.Top  = this.groupBox1.Height - this.lblSQL.Height - 5;
			this.listView1.Height = this.lblSQL.Top - this.listView1.Top - 5;

			// TODO: Add any initialization after the InitializeComponent call

		}
		public void loadvalues(string strConn, string strSelectSQL,string strDisplayColumn,string strSelectColumn,string strTable)
		{
            	    
            this.m_strCurrentConnection = strConn;
		    this.m_strTable = strTable;

            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                if (dataMgr.m_intError !=0)
                {
                    dataMgr = null;
                    return;
                }
                loadvalues(dataMgr, oConn, strSelectSQL, strDisplayColumn, strSelectColumn, strTable);
            }
            dataMgr = null;
        }
        public void loadvalues(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strSelectSQL, string p_strDisplayColumn, string p_strSelectColumn, string p_strTable)
        {
            ((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.OK;
            p_oDataMgr.SqlQueryReader(p_oConn, p_strSelectSQL);
            if (p_oDataMgr.m_intError != 0)
            {
               return;
            }
            if (p_oDataMgr.m_DataReader.HasRows)
            {
                this.listView1.Clear();
                this.listView1.Columns.Add(" ", 10, HorizontalAlignment.Center);
                this.m_strFieldTypeAString_YN = new string[p_oDataMgr.m_DataReader.FieldCount];
                for (int x = 0; x <= p_oDataMgr.m_DataReader.FieldCount - 1; x++)
                {
                    this.m_strFieldTypeAString_YN[x] = p_oDataMgr.getIsTheFieldAStringDataType(p_oDataMgr.m_DataReader.GetFieldType(x).FullName.ToString());
                    this.listView1.Columns.Add(p_oDataMgr.m_DataReader.GetName(x).ToString(), 100, HorizontalAlignment.Left);
                    if (p_oDataMgr.m_DataReader.GetName(x).ToString().Trim().ToUpper() ==
                         p_strDisplayColumn.Trim().ToUpper()) this.m_intDisplayColumn = x;
                    if (p_oDataMgr.m_DataReader.GetName(x).ToString().Trim().ToUpper() ==
                        p_strSelectColumn.Trim().ToUpper()) this.m_intSelectColumn = x;
                }

                while (p_oDataMgr.m_DataReader.Read())
                {
                    if (p_oDataMgr.m_DataReader[0] != System.DBNull.Value)
                    {
                        if (p_oDataMgr.m_DataReader[0].ToString().Trim().Length > 0)
                        {
                            System.Windows.Forms.ListViewItem entryListItem =
                                listView1.Items.Add("",0);
                            this.m_oLvAlternateColors.AddRow();
                            this.m_oLvAlternateColors.AddColumns(entryListItem.Index, listView1.Columns.Count);
                            entryListItem.UseItemStyleForSubItems = false;
                            this.m_oLvAlternateColors.ListViewSubItem(entryListItem.Index, 0, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                            for (int x = 0; x <= p_oDataMgr.m_DataReader.FieldCount - 1; x++)
                            {
                               
                                    this.listView1.Items[entryListItem.Index].SubItems.Add(p_oDataMgr.m_DataReader[x].ToString().Trim());
                                    this.m_oLvAlternateColors.ListViewSubItem(entryListItem.Index, x+1, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                                

                            }
                            this.listView1.Items[0].Selected = true;

                        }
                    }

                }
            }
            else
            {
                MessageBox.Show("No previous data to choose from", "FIA Biosum",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                ((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.Cancel;
            }
            p_oDataMgr.m_DataReader.Close();
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
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblSQL = new System.Windows.Forms.Label();
            this.listView1 = new System.Windows.Forms.ListView();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSelect = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnRecall = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(0, 0);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(592, 32);
            this.lblTitle.TabIndex = 26;
            this.lblTitle.Text = "Previous Expressions";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblSQL);
            this.groupBox1.Controls.Add(this.listView1);
            this.groupBox1.Location = new System.Drawing.Point(16, 40);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(560, 336);
            this.groupBox1.TabIndex = 27;
            this.groupBox1.TabStop = false;
            // 
            // lblSQL
            // 
            this.lblSQL.BackColor = System.Drawing.Color.White;
            this.lblSQL.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblSQL.ForeColor = System.Drawing.Color.Black;
            this.lblSQL.Location = new System.Drawing.Point(19, 176);
            this.lblSQL.Name = "lblSQL";
            this.lblSQL.Size = new System.Drawing.Size(520, 144);
            this.lblSQL.TabIndex = 30;
            // 
            // listView1
            // 
            this.listView1.GridLines = true;
            this.listView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(14, 24);
            this.listView1.MultiSelect = false;
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(536, 136);
            this.listView1.TabIndex = 29;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.SelectedIndexChanged += new System.EventHandler(this.listView1_SelectedIndexChanged);
            this.listView1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.listView1_MouseUp);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(364, 384);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(72, 32);
            this.btnCancel.TabIndex = 28;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSelect
            // 
            this.btnSelect.Location = new System.Drawing.Point(148, 384);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(72, 32);
            this.btnSelect.TabIndex = 29;
            this.btnSelect.Text = "Select";
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(220, 384);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(72, 32);
            this.btnDelete.TabIndex = 30;
            this.btnDelete.Text = "Delete";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnRecall
            // 
            this.btnRecall.Location = new System.Drawing.Point(292, 384);
            this.btnRecall.Name = "btnRecall";
            this.btnRecall.Size = new System.Drawing.Size(72, 32);
            this.btnRecall.TabIndex = 31;
            this.btnRecall.Text = "Recall";
            this.btnRecall.Click += new System.EventHandler(this.btnRecall_Click);
            // 
            // uc_previous_expressions
            // 
            this.Controls.Add(this.btnRecall);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnSelect);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lblTitle);
            this.Name = "uc_previous_expressions";
            this.Size = new System.Drawing.Size(592, 464);
            this.Resize += new System.EventHandler(this.uc_previous_expressions_Resize);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void uc_previous_expressions_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.groupBox1.Left = 10;
				this.groupBox1.Width = this.lblTitle.Width - 20;
				this.listView1.Left = 10;
				this.listView1.Width = this.groupBox1.Width - 20;
				this.lblSQL.Left = this.listView1.Left;
				this.lblSQL.Width = this.listView1.Width;
				this.btnSelect.Top = this.Height - this.btnSelect.Height - 10;
				this.btnDelete.Top = this.btnSelect.Top;
				this.btnCancel.Top = this.btnSelect.Top;
				this.btnRecall.Top = this.btnSelect.Top;
				this.btnSelect.Left = (int)(this.Width  * .50) -
					this.btnDelete.Width - this.btnSelect.Width;
				this.btnDelete.Left = this.btnSelect.Left + this.btnSelect.Width;
				this.btnRecall.Left = this.btnDelete.Left + this.btnDelete.Width;
				this.btnCancel.Left = this.btnRecall.Left + this.btnRecall.Width;
				this.groupBox1.Height = this.btnSelect.Top - this.groupBox1.Top - 5;
				this.lblSQL.Top  = this.groupBox1.Height - this.lblSQL.Height - 5;
				this.listView1.Height = this.lblSQL.Top - this.listView1.Top - 5;
			}
			catch
			{
			}

		}

		private void listView1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.listView1.SelectedItems.Count == 0) return; 
			
			if (listView1.SelectedItems.Count > 0)
			{
				m_intCurrSelect = listView1.SelectedItems[0].Index;
				this.m_oLvAlternateColors.DelegateListViewItem(listView1.SelectedItems[0]);
			}
			this.lblSQL.Text = this.listView1.SelectedItems[0].SubItems[this.m_intDisplayColumn + 1].Text;


		}

		private void btnDelete_Click(object sender, System.EventArgs e)
		{
			if (this.listView1.SelectedItems.Count==0) return;

			this.listView1.SelectedItems[0].Text ="*";
			this.listView1.Focus();
		}

		private void btnRecall_Click(object sender, System.EventArgs e)
		{
			if (this.listView1.SelectedItems.Count==0) return;

			this.listView1.SelectedItems[0].Text =" ";
			this.listView1.Focus();

		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{

			((frmDialog)this.ParentForm).Close();
		}

		public void DeleteRecords()
        {
			int row = 0;
			int col = 0;
			bool lprompt = false;
			DialogResult result = DialogResult.None;
			SQLite.ADO.DataMgr p_dataMgr = new SQLite.ADO.DataMgr();

			//check to see if there are any records to delete
			for (row = 0; row <= this.listView1.Items.Count - 1; row++)
            {
				if (this.listView1.Items[row].Text == "*")
                {
					if (lprompt == false)
					{
						result = MessageBox.Show("Permenently delete those items marked for deletion? Y/N", "SQL Expressions", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

					}
					if (result == DialogResult.Yes || lprompt == true)
                    {
						p_dataMgr.OpenConnection(this.m_strCurrentConnection);
						if (p_dataMgr.m_intError != 0)
						{

							p_dataMgr = null;
							return;
						}
						lprompt = true;
					}
					System.Text.StringBuilder sb = new System.Text.StringBuilder();
					string strSQL = "DELETE * FROM " + this.m_strTable +
							  " WHERE ";

					sb.Append(strSQL);
					for (col = 1; col <= this.listView1.Columns.Count - 1; col++)
					{
						if (this.m_strFieldTypeAString_YN[col - 1] == "Y")
						{

							if (this.listView1.Items[row].SubItems[col].Text.Trim().Length == 0)
							{
								strSQL = "(" + this.listView1.Columns[col].Text.Trim() + " IS NULL OR LENGTH(TRIM(" + this.listView1.Columns[col].Text.Trim() + ")) = 0)";
								sb.Append(strSQL);
							}
							else
							{
								strSQL = "TRIM(UPPER(" + this.listView1.Columns[col].Text.Trim() + ")) = ";
								sb.Append(strSQL);
								sb.Append("'");
								if (this.listView1.Items[row].SubItems[col].Text.IndexOf("'", 0) > 0)
								{
									strSQL = p_dataMgr.FixString(this.listView1.Items[row].SubItems[col].Text, "'", "''");
									sb.Append(strSQL);
								}
								else
								{
									strSQL = this.listView1.Items[row].SubItems[col].Text.Trim().ToUpper();
									sb.Append(strSQL);
								}
								sb.Append("'");

							}
						}
						else
						{
							if (this.listView1.Items[row].SubItems[col].Text.Trim().Length == 0)
							{
								strSQL = this.listView1.Columns[col].Text.Trim() + " IS NULL";
								sb.Append(strSQL);
							}
							else
							{
								strSQL = this.listView1.Columns[col].Text.Trim() + " = ";
								sb.Append(strSQL);
								strSQL = this.listView1.Items[row].SubItems[col].Text.Trim();
								sb.Append(strSQL);
							}
						}
						if (col < this.listView1.Columns.Count - 1)
						{
							strSQL = " AND ";
							sb.Append(strSQL);
						}
					}
					if (lprompt == true)
						p_dataMgr.SqlNonQuery(p_dataMgr.m_Connection, sb.ToString());
					sb = null;
				}

			}
			if (lprompt == true)
			{
				p_dataMgr.m_Connection.Close();
			}
			p_dataMgr = null;
		}

		private void btnSelect_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.OK;
		}

		private void listView1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = this.listView1.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.listView1.Items[this.listView1.TopItem.Index + (int)dblRow-1].Selected=true;
				}
			}
			catch 
			{
			}
		}
        public bool EnableDeleteButton
        {
            set { this.btnDelete.Enabled = value; }
        }
        public bool EnableRecallButton
        {
            set { this.btnRecall.Enabled = value; }
        }
        public bool ShowDeleteButton
        {
            set { this.btnDelete.Visible = value; }
        }
        public bool ShowRecallButton
        {
            set { this.btnRecall.Visible = value; }
        }

		public void LoadRxItemCollection(FIA_Biosum_Manager.RxItem_Collection p_oRxItemCollection)
		{
			listView1.Clear();
			listView1.FullRowSelect=false;
			m_oLvAlternateColors.InitializeRowCollection();      
			listView1.Columns.Add(" ",2,HorizontalAlignment.Left);
			listView1.Columns.Add("Rx", 60, HorizontalAlignment.Left);
			listView1.Columns.Add("Description", 500, HorizontalAlignment.Left);
			if (p_oRxItemCollection != null)
			{
				for (int x=0;x<=p_oRxItemCollection.Count-1;x++)
				{
					listView1.Items.Add(" ");
					listView1.Items[listView1.Items.Count-1].UseItemStyleForSubItems=false;
					listView1.Items[listView1.Items.Count-1].SubItems.Add(p_oRxItemCollection.Item(x).RxId);
					listView1.Items[listView1.Items.Count-1].SubItems.Add(p_oRxItemCollection.Item(x).Description);

					m_oLvAlternateColors.AddRow();
					m_oLvAlternateColors.AddColumns(x,listView1.Columns.Count);
					
				}
			}
			m_oLvAlternateColors.ListView();
			if (this.listView1.Items.Count > 0) this.listView1.Items[0].Selected=true;
			lblTitle.Text = "Treatment Prescriptions";
			m_intSelectColumn =0;
			m_intDisplayColumn=1;
			this.btnDelete.Hide();
			this.btnRecall.Hide();
			Show();

			((frmDialog)this.ParentForm).WindowState= System.Windows.Forms.FormWindowState.Maximized;
		}
        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_oRxPackageItemCollection">Collection of all the RxPackage Items defined in the project</param>
        /// <param name="p_strRxPackageArray">Contains RxPackage items that will be checked</param>
        public void LoadRxPackageItemCollection(FIA_Biosum_Manager.RxPackageItem_Collection p_oRxPackageItemCollection,string[] p_strRxPackageArray)
        {
            listView1.CheckBoxes=true;
            listView1.Clear();
            listView1.FullRowSelect = false;
            m_oLvAlternateColors.InitializeRowCollection();
            listView1.Columns.Add(" ", 30, HorizontalAlignment.Left);
            listView1.Columns.Add("RxPackage", 100, HorizontalAlignment.Left);
            listView1.Columns.Add("RxCycleLength", 110, HorizontalAlignment.Left);
            listView1.Columns.Add("Cycle1Rx", 80, HorizontalAlignment.Left);
            listView1.Columns.Add("Cycle2Rx", 80, HorizontalAlignment.Left);
            listView1.Columns.Add("Cycle3Rx", 80, HorizontalAlignment.Left);
            listView1.Columns.Add("Cycle4Rx", 80, HorizontalAlignment.Left);
            listView1.Columns.Add("Description", 500, HorizontalAlignment.Left);
            if (p_oRxPackageItemCollection != null)
            {
                for (int x = 0; x <= p_oRxPackageItemCollection.Count - 1; x++)
                {
                    listView1.Items.Add(" ");
                    listView1.Items[listView1.Items.Count - 1].UseItemStyleForSubItems = false;
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(p_oRxPackageItemCollection.Item(x).RxPackageId);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(p_oRxPackageItemCollection.Item(x).RxCycleLength.ToString().Trim());
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(p_oRxPackageItemCollection.Item(x).SimulationYear1Rx);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(p_oRxPackageItemCollection.Item(x).SimulationYear2Rx);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(p_oRxPackageItemCollection.Item(x).SimulationYear3Rx);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(p_oRxPackageItemCollection.Item(x).SimulationYear4Rx);
                    listView1.Items[listView1.Items.Count - 1].SubItems.Add(p_oRxPackageItemCollection.Item(x).Description);
                    if (p_strRxPackageArray != null)
                    {
                        for (int y = 0; y <= p_strRxPackageArray.Length - 1; y++)
                        {
                            if (p_strRxPackageArray[y] != null)
                            {
                                if (p_oRxPackageItemCollection.Item(x).RxPackageId.Trim()==
                                    p_strRxPackageArray[y].Trim())
                                {
                                    listView1.Items[listView1.Items.Count-1].Checked=true;
                                }
                            }
                        }
                    }
                    m_oLvAlternateColors.AddRow();
                    m_oLvAlternateColors.AddColumns(x, listView1.Columns.Count);

                }
            }
            m_oLvAlternateColors.ListView();
            if (this.listView1.Items.Count > 0) this.listView1.Items[0].Selected = true;
            lblTitle.Text = "Treatment Prescription Packages";
            m_intSelectColumn = 0;
            m_intDisplayColumn = 6;
            this.btnDelete.Hide();
            this.btnRecall.Hide();
            btnSelect.Text = "OK";
            Show();

            ((frmDialog)this.ParentForm).WindowState = System.Windows.Forms.FormWindowState.Maximized;
        }
	}
}

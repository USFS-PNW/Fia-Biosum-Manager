using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using SQLite.ADO;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_datasource_edit.
	/// </summary>
	public class uc_datasource_edit : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.GroupBox groupBox2;
		public System.Windows.Forms.Label lblTableType;
		private System.Windows.Forms.GroupBox groupBox3;
		public System.Windows.Forms.Label lblDBFile;
		private System.Windows.Forms.Button btnFile;
		private System.Windows.Forms.Button btnMove;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnCancel;
		public System.Windows.Forms.Label lblTable;
		private System.Windows.Forms.GroupBox groupBox4;
		public System.Windows.Forms.Label lblNewDBFile;
		public System.Windows.Forms.Label lblNewTable;
		private string strAction;
		private System.Windows.Forms.Button btnCommitChange;
		private System.Windows.Forms.Button btnCopy;
		public string m_strProjectFile="";
		public string m_strProjectDirectory="";
		public string m_strScenarioId="";
		public string m_strScenarioFile="";
		private System.Windows.Forms.Label lblProgress;
		private bool m_bOverwriteTable;
		private System.Windows.Forms.Button btnTableName;
		private string m_strDataSourceDBFile;
		private string m_strDataSourceTable;
		private System.Windows.Forms.ProgressBar progressBar1;
		private System.Windows.Forms.Button btnHelp;
		private System.Windows.Forms.Button btnCopyToSameDbFile;

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultDatabaseXPSFile;

		
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_datasource_edit()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
            this.strAction="";
            this.m_oEnv = new env();
			// TODO: Add any initialization after the InitializeComponent call

		}
		public uc_datasource_edit(string p_strProjectDBFile)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.strAction="";
			this.m_strDataSourceDBFile=p_strProjectDBFile.Trim();
			this.m_strDataSourceTable = "datasource";
            this.m_oEnv = new env();
			// TODO: Add any initialization after the InitializeComponent call

		}
		public uc_datasource_edit(string p_strScenarioDBFile, string p_strScenarioId)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.strAction="";
			this.m_strDataSourceDBFile = p_strScenarioDBFile.Trim();
			this.m_strDataSourceTable = "scenario_datasource";
			this.m_strScenarioId=p_strScenarioId;
            this.m_oEnv = new env();
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
		
		#endregion
		private void InitializeComponent()
		{
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnHelp = new System.Windows.Forms.Button();
			this.progressBar1 = new System.Windows.Forms.ProgressBar();
			this.lblProgress = new System.Windows.Forms.Label();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.lblNewTable = new System.Windows.Forms.Label();
			this.lblNewDBFile = new System.Windows.Forms.Label();
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnCommitChange = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.lblTable = new System.Windows.Forms.Label();
			this.btnFile = new System.Windows.Forms.Button();
			this.lblDBFile = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.lblTableType = new System.Windows.Forms.Label();
			this.lblTitle = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			this.groupBox4.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btnHelp);
			this.groupBox1.Controls.Add(this.progressBar1);
			this.groupBox1.Controls.Add(this.lblProgress);
			this.groupBox1.Controls.Add(this.groupBox4);
			this.groupBox1.Controls.Add(this.btnCancel);
			this.groupBox1.Controls.Add(this.btnCommitChange);
			this.groupBox1.Controls.Add(this.btnClose);
			this.groupBox1.Controls.Add(this.groupBox3);
			this.groupBox1.Controls.Add(this.groupBox2);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(696, 464);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// btnHelp
			// 
			this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
			this.btnHelp.Location = new System.Drawing.Point(8, 416);
			this.btnHelp.Name = "btnHelp";
			this.btnHelp.Size = new System.Drawing.Size(96, 32);
			this.btnHelp.TabIndex = 35;
			this.btnHelp.Text = "Help";
			this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
			// 
			// progressBar1
			// 
			this.progressBar1.Location = new System.Drawing.Point(224, 424);
			this.progressBar1.Name = "progressBar1";
			this.progressBar1.Size = new System.Drawing.Size(240, 8);
			this.progressBar1.TabIndex = 34;
			this.progressBar1.Visible = false;
			// 
			// lblProgress
			// 
			this.lblProgress.Location = new System.Drawing.Point(208, 440);
			this.lblProgress.Name = "lblProgress";
			this.lblProgress.Size = new System.Drawing.Size(360, 16);
			this.lblProgress.TabIndex = 33;
			this.lblProgress.Text = "lblProgress";
			this.lblProgress.Visible = false;
			// 
			// groupBox4
			// 
			this.groupBox4.Controls.Add(this.lblNewTable);
			this.groupBox4.Controls.Add(this.lblNewDBFile);
			this.groupBox4.Location = new System.Drawing.Point(8, 240);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(680, 80);
			this.groupBox4.TabIndex = 32;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "New Database File And Table";
			// 
			// lblNewTable
			// 
			this.lblNewTable.BackColor = System.Drawing.Color.White;
			this.lblNewTable.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblNewTable.Location = new System.Drawing.Point(16, 48);
			this.lblNewTable.Name = "lblNewTable";
			this.lblNewTable.Size = new System.Drawing.Size(256, 16);
			this.lblNewTable.TabIndex = 1;
			// 
			// lblNewMDBFile
			// 
			this.lblNewDBFile.BackColor = System.Drawing.Color.White;
			this.lblNewDBFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblNewDBFile.Location = new System.Drawing.Point(16, 24);
			this.lblNewDBFile.Name = "lblNewMDBFile";
			this.lblNewDBFile.Size = new System.Drawing.Size(648, 16);
			this.lblNewDBFile.TabIndex = 0;
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(423, 376);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(90, 40);
			this.btnCancel.TabIndex = 31;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// btnCommitChange
			// 
			this.btnCommitChange.Location = new System.Drawing.Point(264, 376);
			this.btnCommitChange.Name = "btnCommitChange";
			this.btnCommitChange.Size = new System.Drawing.Size(130, 40);
			this.btnCommitChange.TabIndex = 30;
			this.btnCommitChange.Text = "Commit Change";
			this.btnCommitChange.Click += new System.EventHandler(this.btnCommitChange_Click);
			// 
			// btnClose
			// 
			this.btnClose.Location = new System.Drawing.Point(592, 416);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(96, 32);
			this.btnClose.TabIndex = 29;
			this.btnClose.Text = "Close";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.lblTable);
			this.groupBox3.Controls.Add(this.btnFile);
			this.groupBox3.Controls.Add(this.lblDBFile);
			this.groupBox3.Location = new System.Drawing.Point(8, 104);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(680, 120);
			this.groupBox3.TabIndex = 27;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Current Database File And Table";
			// 
			// lblTable
			// 
			this.lblTable.BackColor = System.Drawing.Color.White;
			this.lblTable.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblTable.Location = new System.Drawing.Point(16, 48);
			this.lblTable.Name = "lblTable";
			this.lblTable.Size = new System.Drawing.Size(256, 16);
			this.lblTable.TabIndex = 28;
			// 
			// btnFile
			// 
			this.btnFile.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnFile.Location = new System.Drawing.Point(16, 80);
			this.btnFile.Name = "btnFile";
			this.btnFile.Size = new System.Drawing.Size(280, 24);
			this.btnFile.TabIndex = 8;
			this.btnFile.Text = "Get A SQLite Db File And Table";
			this.btnFile.Click += new System.EventHandler(this.btnFile_Click);
			// 
			// lblMDBFile
			// 
			this.lblDBFile.BackColor = System.Drawing.Color.White;
			this.lblDBFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblDBFile.Location = new System.Drawing.Point(16, 24);
			this.lblDBFile.Name = "lblMDBFile";
			this.lblDBFile.Size = new System.Drawing.Size(648, 16);
			this.lblDBFile.TabIndex = 0;
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.lblTableType);
			this.groupBox2.Location = new System.Drawing.Point(8, 56);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(680, 40);
			this.groupBox2.TabIndex = 26;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Table Type";
			// 
			// lblTableType
			// 
			this.lblTableType.BackColor = System.Drawing.Color.White;
			this.lblTableType.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblTableType.Location = new System.Drawing.Point(16, 16);
			this.lblTableType.Name = "lblTableType";
			this.lblTableType.Size = new System.Drawing.Size(648, 16);
			this.lblTableType.TabIndex = 0;
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 18);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(690, 32);
			this.lblTitle.TabIndex = 25;
			this.lblTitle.Text = "Data Source Edit";
			// 
			// uc_datasource_edit
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_datasource_edit";
			this.Size = new System.Drawing.Size(696, 464);
			this.Resize += new System.EventHandler(this.uc_datasource_edit_Resize);
			this.groupBox1.ResumeLayout(false);
			this.groupBox4.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		private void btnFile_Click(object sender, System.EventArgs e)
        {
			string strFullPath;
			string strLargestString = "";

			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "Get Database File Containing " + this.lblTableType.Text + " Table";
			OpenFileDialog1.Filter = "SQLite Database File (*.DB) | *.db";
			strFullPath = this.lblDBFile.Text.Trim();
			OpenFileDialog1.InitialDirectory = strFullPath;

			DialogResult result = OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK)
            {
				strFullPath = OpenFileDialog1.FileName.Trim();
				if (strFullPath.Length > 0)
                {
					DataMgr dataMgr = new DataMgr();
					string[] arrTables;
					
					using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dataMgr.GetConnectionString(strFullPath)))
                    {
						conn.Open();

						arrTables = dataMgr.getTableNames(conn);
                    }

					if (dataMgr.m_intError == 0)
                    {
						frmDialog frmTemp = new frmDialog();
						frmTemp.Text = "Select " + this.lblTableType.Text + " Table";

						frmTemp.uc_select_list_item1.lblMsg.Text = "Table contents of " + strFullPath;
						frmTemp.uc_select_list_item1.lblMsg.Visible = true;
						strLargestString = frmTemp.uc_select_list_item1.lblMsg.Text;
						frmTemp.uc_project1.Visible = false;
						frmTemp.uc_select_list_item1.listBox1.Items.Clear();

						for (int x = 0; x <= arrTables.Length - 1; x++)
                        {
							frmTemp.uc_select_list_item1.listBox1.Items.Add(arrTables[x]);
							if (arrTables[x].Trim().Length > strLargestString.Trim().Length)
                            {
								strLargestString = arrTables[x];
							}
						}

						frmTemp.uc_select_list_item1.Initialize_Width(strLargestString);
						frmTemp.uc_select_list_item1.Visible = true;
						result = frmTemp.ShowDialog(this);

						if (result == DialogResult.OK)
						{
							//validation of the table chose will be done here
							this.lblNewDBFile.Text = strFullPath;
							this.lblNewTable.Text = frmTemp.uc_select_list_item1.listBox1.Text;
						}

						frmTemp.Close();
						frmTemp = null;
					}
				}
			}
		}

		private void uc_datasource_edit_Resize(object sender, System.EventArgs e)
		{
             this.resize_uc_datasource_edit();
		}
		public void resize_uc_datasource_edit()
		{
			try
			{
				this.btnClose.Top = this.groupBox1.Height - (int)(this.btnClose.Height) - 5;
				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
				this.groupBox2.Left = 2;
				this.groupBox3.Left = 2;
				this.groupBox4.Left = 2;
				this.groupBox2.Width = this.Width - 4;
				this.groupBox3.Width = this.groupBox2.Width;
				this.groupBox4.Width = this.groupBox2.Width;
				this.lblDBFile.Width = this.groupBox2.Width - (this.lblDBFile.Left * 2);
				this.lblNewDBFile.Width = this.groupBox2.Width - (this.lblDBFile.Left * 2);
				this.lblTableType.Width = this.lblDBFile.Width;
				this.btnCancel.Top = this.groupBox4.Top + this.groupBox4.Height  + 5;
				this.btnCommitChange.Top = this.btnCancel.Top;
				this.btnCancel.Left = (int) (this.Width * .50) ; //+ (int) (this.btnSave.Width / 2);
				this.btnCommitChange.Left = this.btnCancel.Left - this.btnCancel.Width -50;
				this.btnHelp.Top = this.btnClose.Top;
				this.progressBar1.Left = (int)(this.groupBox1.Width * .50) - (int)(this.progressBar1.Width * .50);
				this.progressBar1.Top = this.btnCommitChange.Top + this.btnCommitChange.Height + 10;
				this.lblProgress.Left = this.progressBar1.Left;
				this.lblProgress.Top = this.progressBar1.Top + this.progressBar1.Height + 2;
			}
			catch
			{
			}
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
		    this.ParentForm.Close();
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
		    this.ParentForm.Close();
		}

		private void btnCommitChange_Click(object sender, System.EventArgs e)
        {
			string strConn = "";
			string strDir = "";
			string strFile = "";

			this.progressBar1.Left = (int)(this.groupBox1.Width * .50) - (int)(this.progressBar1.Width * .50);
			this.progressBar1.Top = this.btnCommitChange.Top + this.btnCommitChange.Height + 10;
			this.lblProgress.Left = this.progressBar1.Left;
			this.lblProgress.Top = this.progressBar1.Top + this.progressBar1.Height + 2;

			DataMgr dataMgr = new DataMgr();
			utils p_utils = new utils();

			strFile = p_utils.getFileName(this.lblNewDBFile.Text.Trim());
			strDir = p_utils.getDirectory(this.lblNewDBFile.Text.Trim());

			strConn = dataMgr.GetConnectionString(this.m_strDataSourceDBFile);
			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
				conn.Open();

				dataMgr.m_strSQL = "UPDATE " + this.m_strDataSourceTable +
					" SET path = '" + strDir + "', " +
						"file = '" + strFile + "', " +
						"table_name = '" + this.lblNewTable.Text.Trim() + "' " +
						"WHERE table_type = '" + this.lblTableType.Text.Trim() + "'";

				if (this.m_strScenarioId.Trim().Length > 0)
                {
					dataMgr.m_strSQL += "AND scenario_id = '" + this.m_strScenarioId + "'";
                }

				dataMgr.SqlNonQuery(conn, dataMgr.m_strSQL);
			}
			if (this.lblNewDBFile.Text.Length == 0 ||
				this.lblNewTable.Text.Length == 0)
			{
				MessageBox.Show("Not Valid: New DB File and/or Table are blank");
			}
			else
			{
				if (dataMgr.m_intError == 0)
				{	
					this.ParentForm.DialogResult = System.Windows.Forms.DialogResult.OK;
				}
			}
			dataMgr = null;
			p_utils = null;
		}
		
		public System.Windows.Forms.ProgressBar _ProgressBar
		{
			get 
			{
				return this.progressBar1;
			}
		}
		public System.Windows.Forms.Label _lblProgress
		{
			get
			{
				return this.lblProgress;
			}
		}
		public string strProjectDirectory
		{
			get 
			{
				return this.m_strProjectDirectory;
			}
			set
			{
				this.m_strProjectDirectory=value;
			}
		}
		public string strScenarioId
		{
			get
			{
				return this.m_strScenarioId;
			}
			set
			{
				this.m_strScenarioId = value;
			}
		}

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "DATABASE", "EDIT_DATA_SOURCE" });
        }


	}
}

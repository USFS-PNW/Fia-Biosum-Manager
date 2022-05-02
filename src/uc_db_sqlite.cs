using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_db.
	/// </summary>
	public class uc_db_sqlite : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.TreeView treeView1;
		public int m_intError=0;
		private string m_strProjDir;
        private SQLite.ADO.DataMgr m_DataMgr;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnHelp;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.ListBox lstTables;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.CheckedListBox lstFields;
		private System.ComponentModel.IContainer components;
        private string m_strCurDBFile = "";
        private string m_strCurConn="";
		private string m_strCurTable="";
		private System.Windows.Forms.Button btnBrowse;
		private System.Windows.Forms.Button btnCheckAll;
		private System.Windows.Forms.Button btnClearAll;
		string substringDirectory;
		private System.Windows.Forms.GroupBox grpMdb;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label lblFileSize;
		private System.Windows.Forms.Button btnOpen;
		FIA_Biosum_Manager.frmGridView m_frmGridView;

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultDatabaseXPSFile;
		

		public uc_db_sqlite(string p_strProjDir)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_strProjDir = p_strProjDir.Trim();
            this.m_DataMgr = new SQLite.ADO.DataMgr();
            this.m_oEnv = new env();

			// TODO: Add any initialization after the InitializeComponent call

		}
		public void loadvalues()
		{
			//TreeNode node1;
			//TreeNode node2;
            treeView1.Nodes.Add(this.m_strProjDir);
			treeView1.Nodes[0].ImageIndex=0;
			treeView1.Nodes[0].SelectedImageIndex=1;
			//string strNode="";

			this.PopulateTreeView(this.m_strProjDir,this.treeView1.Nodes[0]);
		}
		public void PopulateTreeView(string directoryValue, TreeNode parentNode )
		{
			string strNode;
			TreeNode node2;
			string[] directoryArray = 
				System.IO.Directory.GetDirectories( directoryValue );

			try
			{
				if ( directoryArray.Length != 0 )
				{
					foreach ( string directory in directoryArray )
					{
						substringDirectory = directory.Substring(
							directory.LastIndexOf( '\\' ) + 1,
							directory.Length - directory.LastIndexOf( '\\' ) - 1 );

						TreeNode myNode = new TreeNode( substringDirectory );
						myNode.ImageIndex=0;
						myNode.SelectedImageIndex=1;

						parentNode.Nodes.Add( myNode );
						string[] dbFiles = System.IO.Directory.GetFiles(myNode.FullPath,"*.db");
						string[] db3Files = System.IO.Directory.GetFiles(myNode.FullPath,"*.db3");
						string[] allFiles = new string[dbFiles.Length + db3Files.Length];
						
						if (dbFiles.Length > 0)
						{
							dbFiles.CopyTo(allFiles,0);
							if (db3Files.Length > 0)
                                db3Files.CopyTo(allFiles,dbFiles.Length);
						}
						else
						{
							if (db3Files.Length > 0)
                                db3Files.CopyTo(allFiles,0);
						}

						

						if (allFiles.Length > 0 && allFiles[0] != null && allFiles[0].Trim().Length > 0)
						{
					
							for (int y=0;y<=allFiles.Length -1;y++)
							{
								strNode = this.getSubDir(allFiles[y]);
                                // This section is intended to exclude some database files; Do we want to do this?
                                // If so, add additional clauses to the != statement
								if (strNode.Trim().ToUpper() != "PROJECT.MDB")
								{
									node2 = new TreeNode(strNode);
									node2.ImageIndex=0;
									node2.SelectedImageIndex=1;
									myNode.Nodes.Add(node2);
								}

							}
						}


						PopulateTreeView( directory, myNode );
					}
				}
			} 
			catch ( UnauthorizedAccessException ) 
			{
				parentNode.Nodes.Add( "Access denied" );
			} // end catch
		}

		private string getSubDir(string strFullPath)
		{
			string str="";
			for (int x=strFullPath.Trim().Length -1; x >= 0;x--)
			{
				if (strFullPath.Substring(x,1) != "\\")
				{
					str = strFullPath.Substring(x,1) + str;
				}
				else break;
			}
			return str;
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_db_sqlite));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.grpMdb = new System.Windows.Forms.GroupBox();
            this.btnOpen = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.lblFileSize = new System.Windows.Forms.Label();
            this.btnClearAll = new System.Windows.Forms.Button();
            this.btnCheckAll = new System.Windows.Forms.Button();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.lstFields = new System.Windows.Forms.CheckedListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lstTables = new System.Windows.Forms.ListBox();
            this.btnHelp = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.groupBox1.SuspendLayout();
            this.grpMdb.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.grpMdb);
            this.groupBox1.Controls.Add(this.btnClearAll);
            this.groupBox1.Controls.Add(this.btnCheckAll);
            this.groupBox1.Controls.Add(this.btnBrowse);
            this.groupBox1.Controls.Add(this.lstFields);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.lstTables);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.treeView1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(744, 608);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // grpMdb
            // 
            this.grpMdb.Controls.Add(this.btnOpen);
            this.grpMdb.Controls.Add(this.label4);
            this.grpMdb.Controls.Add(this.lblFileSize);
            this.grpMdb.Location = new System.Drawing.Point(8, 440);
            this.grpMdb.Name = "grpMdb";
            this.grpMdb.Size = new System.Drawing.Size(232, 88);
            this.grpMdb.TabIndex = 51;
            this.grpMdb.TabStop = false;
            // 
            // btnOpen
            // 
            this.btnOpen.Enabled = false;
            this.btnOpen.Location = new System.Drawing.Point(0, 58);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(182, 24);
            this.btnOpen.TabIndex = 53;
            this.btnOpen.Text = "Open With SQLite Editor";
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(8, 16);
            this.label4.Name = "label4";
            this.label4.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.label4.Size = new System.Drawing.Size(64, 16);
            this.label4.TabIndex = 52;
            this.label4.Text = "File Size";
            // 
            // lblFileSize
            // 
            this.lblFileSize.BackColor = System.Drawing.Color.White;
            this.lblFileSize.Location = new System.Drawing.Point(88, 16);
            this.lblFileSize.Name = "lblFileSize";
            this.lblFileSize.Size = new System.Drawing.Size(136, 24);
            this.lblFileSize.TabIndex = 52;
            // 
            // btnClearAll
            // 
            this.btnClearAll.Location = new System.Drawing.Point(568, 488);
            this.btnClearAll.Name = "btnClearAll";
            this.btnClearAll.Size = new System.Drawing.Size(64, 24);
            this.btnClearAll.TabIndex = 50;
            this.btnClearAll.Text = "Clear All";
            this.btnClearAll.Click += new System.EventHandler(this.btnClearAll_Click);
            // 
            // btnCheckAll
            // 
            this.btnCheckAll.Location = new System.Drawing.Point(504, 488);
            this.btnCheckAll.Name = "btnCheckAll";
            this.btnCheckAll.Size = new System.Drawing.Size(64, 24);
            this.btnCheckAll.TabIndex = 49;
            this.btnCheckAll.Text = "Check All";
            this.btnCheckAll.Click += new System.EventHandler(this.btnCheckAll_Click);
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(632, 488);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(64, 24);
            this.btnBrowse.TabIndex = 48;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // lstFields
            // 
            this.lstFields.CheckOnClick = true;
            this.lstFields.Location = new System.Drawing.Point(493, 40);
            this.lstFields.Name = "lstFields";
            this.lstFields.Size = new System.Drawing.Size(237, 378);
            this.lstFields.TabIndex = 45;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(496, 16);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(136, 16);
            this.label3.TabIndex = 44;
            this.label3.Text = "Columns";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(248, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(136, 16);
            this.label2.TabIndex = 42;
            this.label2.Text = "Tables";
            // 
            // lstTables
            // 
            this.lstTables.ItemHeight = 16;
            this.lstTables.Location = new System.Drawing.Point(253, 40);
            this.lstTables.Name = "lstTables";
            this.lstTables.Size = new System.Drawing.Size(237, 388);
            this.lstTables.TabIndex = 41;
            this.lstTables.SelectedIndexChanged += new System.EventHandler(this.lstTables_SelectedIndexChanged);
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(8, 567);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(96, 32);
            this.btnHelp.TabIndex = 40;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(632, 568);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 39;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(8, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(152, 16);
            this.label1.TabIndex = 5;
            this.label1.Text = "SQLite Database Files";
            // 
            // treeView1
            // 
            this.treeView1.ImageIndex = 0;
            this.treeView1.ImageList = this.imageList1;
            this.treeView1.Location = new System.Drawing.Point(10, 40);
            this.treeView1.Name = "treeView1";
            this.treeView1.SelectedImageIndex = 0;
            this.treeView1.Size = new System.Drawing.Size(237, 392);
            this.treeView1.TabIndex = 4;
            this.treeView1.AfterCollapse += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterCollapse);
            this.treeView1.AfterExpand += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterExpand);
            this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "");
            this.imageList1.Images.SetKeyName(1, "");
            // 
            // uc_db_sqlite
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_db_sqlite";
            this.Size = new System.Drawing.Size(744, 608);
            this.Resize += new System.EventHandler(this.uc_db_sqlite_Resize);
            this.groupBox1.ResumeLayout(false);
            this.grpMdb.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void treeView1_AfterExpand(object sender, System.Windows.Forms.TreeViewEventArgs e)
		{
//			e.Node.ImageIndex = 1;
		}

		private void treeView1_AfterCollapse(object sender, System.Windows.Forms.TreeViewEventArgs e)
		{
			
		}

		private void treeView1_AfterSelect(object sender, System.Windows.Forms.TreeViewEventArgs e)
		{
			try
			{
				if (e.Node.Text.ToUpper().IndexOf(".DB",0) > 0 || e.Node.Text.ToUpper().IndexOf(".DB3",0) > 0)
				{
					this.lstFields.Items.Clear();
					this.lstTables.Items.Clear();
                    this.m_DataMgr.LoadTablesIntoListBox(e.Node.FullPath, this.lstTables);
					this.m_strCurDBFile = e.Node.FullPath.Trim();
					this.btnOpen.Enabled=true;
					System.IO.FileInfo fi = new System.IO.FileInfo(this.m_strCurDBFile);
					lblFileSize.Text = fi.Length.ToString();
					fi = null;
				}
				else
				{
					this.lstTables.Items.Clear();
					this.lstFields.Items.Clear();
					this.m_strCurTable="";
                    this.m_strCurDBFile = "";
					this.lblFileSize.Text="";
					this.btnOpen.Enabled=false;
				}
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_db:treeView1_AfterSelect() \n" + 
					"Err Msg - " + err.Message,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
			}
		}

		private void lstTables_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		       this.m_strCurTable = this.lstTables.SelectedItems[0].ToString().Trim();
			   this.lstFields.Items.Clear();
			   this.m_DataMgr.LoadFieldsIntoCheckedListBox(this.m_strCurDBFile,this.m_strCurTable,this.lstFields);
		}

		private void uc_db_sqlite_Resize(object sender, System.EventArgs e)
		{
            if (this.ParentForm.WindowState == System.Windows.Forms.FormWindowState.Minimized) return;
			try
			{
				this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
				this.btnHelp.Top = this.btnClose.Top;
				this.treeView1.Left = 5;
				this.grpMdb.Top = this.btnHelp.Top - (int)(this.grpMdb.Height * 1.5);
				this.treeView1.Top = this.label1.Top + (int)(this.label1.Height * 1.5);
				this.treeView1.Height = this.grpMdb.Top - this.treeView1.Top - 2;
				this.lstTables.Top = this.treeView1.Top;
				this.lstTables.Height = this.treeView1.Height;
				this.lstFields.Top = this.treeView1.Top;
				this.lstFields.Height = this.treeView1.Height;
				this.btnBrowse.Top = this.grpMdb.Top;
				this.btnCheckAll.Top = this.grpMdb.Top;
				this.btnClearAll.Top = this.grpMdb.Top;

				int intWidth = (int)(this.groupBox1.Width * .33);
				this.treeView1.Width = intWidth;
				this.lstTables.Left = this.treeView1.Left + intWidth + 5;
				this.lstTables.Width = intWidth;
				this.lstFields.Left = this.lstTables.Left + intWidth + 5;
				this.lstFields.Width = intWidth;
				if (this.lstFields.Left + this.lstFields.Width >= this.groupBox1.Width)
				{
					for (int x=0;;x++)
					{
						this.lstFields.Width = this.lstFields.Width - x;
	  				    if (this.lstFields.Left + this.lstFields.Width < this.groupBox1.Width - 5) break;
						

					}
				}
				this.grpMdb.Left = this.treeView1.Left;
				this.btnClearAll.Left = this.lstFields.Left + intWidth - (int)(intWidth * .50) - (int)(this.btnClearAll.Width * .50);
				this.btnCheckAll.Left = this.btnClearAll.Left - this.btnClearAll.Width;
				this.btnBrowse.Left = this.btnClearAll.Left + this.btnClearAll.Width;
				this.label1.Left = this.treeView1.Left;
				this.label2.Left = this.lstTables.Left;
				this.label3.Left = this.lstFields.Left;


                
			}
			catch
			{
			}
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.ParentForm.Close();		
		}

		private void btnCheckAll_Click(object sender, System.EventArgs e)
		{
			if (this.lstFields.CheckedItems.Count == this.lstFields.Items.Count ) return;
			for (int x = 0; x<= this.lstFields.Items.Count-1;x++)
			{
				this.lstFields.SetItemChecked(x,true);
			}
		}

		private void btnClearAll_Click(object sender, System.EventArgs e)
		{
			if (this.lstFields.CheckedItems.Count == 0 ) return;
			for (int x = 0; x<= this.lstFields.Items.Count-1;x++)
			{
				this.lstFields.SetItemChecked(x,false);
			}
		}

		private void btnBrowse_Click(object sender, System.EventArgs e)
		{
			if (this.lstFields.Items.Count > 0)
			{
				if (this.lstFields.CheckedItems.Count==0)
				{
					MessageBox.Show("!!Select The Columns To Browse!!","FIA Biosum",
						               System.Windows.Forms.MessageBoxButtons.OK,
						               System.Windows.Forms.MessageBoxIcon.Exclamation);

                    return;
				}

                this.m_strCurConn = this.m_DataMgr.GetConnectionString(this.m_strCurDBFile);
                this.m_DataMgr.m_strSQL = "";
				for (int x=0;x<=this.lstFields.CheckedItems.Count-1;x++)
				{
						if (this.m_DataMgr.m_strSQL.Trim().Length ==0)
						{
							this.m_DataMgr.m_strSQL = this.lstFields.CheckedItems[x].ToString();
						}
						else
						{
							this.m_DataMgr.m_strSQL += "," + this.lstFields.CheckedItems[x].ToString();
						}
				}
				this.m_DataMgr.m_strSQL = "SELECT " + this.m_DataMgr.m_strSQL + " FROM " + this.m_strCurTable;

				this.m_frmGridView = new frmGridView();
				this.m_frmGridView.Text = "Database: Browse (" + this.m_strCurTable.Trim() + ")";
                this.m_frmGridView.UsingSQLite = true;

				this.m_frmGridView.LoadDataSet(this.m_strCurConn,this.m_DataMgr.m_strSQL,this.m_strCurTable);
				this.m_frmGridView.Show();
				this.m_frmGridView.Focus();

			}
		}

		private void btnOpen_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process proc = new System.Diagnostics.Process();
			proc.StartInfo.UseShellExecute = true;
			try
			{
                proc.StartInfo.FileName = this.m_strCurDBFile;
			}
			catch
			{
			}
			try
			{
				proc.Start();
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_db:btnOpen_Click() \n" + 
					"Err Msg - " + err.Message,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
			}
		}

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "DATABASE", "SQLITE_DATA_SOURCES" });

        }
	}
}

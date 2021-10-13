using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

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
		public System.Windows.Forms.Label lblMDBFile;
		private System.Windows.Forms.Button btnFile;
		private System.Windows.Forms.Button btnMove;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnCancel;
		public System.Windows.Forms.Label lblTable;
		private System.Windows.Forms.GroupBox groupBox4;
		public System.Windows.Forms.Label lblNewMDBFile;
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
		private string m_strDataSourceMDBFile;
		private string m_strDataSourceTable;
        private bool m_bUsingSqlite;
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
		public uc_datasource_edit(string p_strProjectMDBFile)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.strAction="";
			this.m_strDataSourceMDBFile=p_strProjectMDBFile.Trim();
			this.m_strDataSourceTable = "datasource";
            this.m_oEnv = new env();
			// TODO: Add any initialization after the InitializeComponent call

		}
		public uc_datasource_edit(string p_strScenarioMDBFile, string p_strScenarioId, bool p_bUsingSqlite)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.strAction="";
			this.m_strDataSourceMDBFile = p_strScenarioMDBFile.Trim();
			this.m_strDataSourceTable = "scenario_datasource";
			this.m_strScenarioId=p_strScenarioId;
            this.m_bUsingSqlite = p_bUsingSqlite;
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
		private void InitializeComponent()
		{
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnHelp = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblProgress = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.btnTableName = new System.Windows.Forms.Button();
            this.lblNewTable = new System.Windows.Forms.Label();
            this.lblNewMDBFile = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnCommitChange = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnCopyToSameDbFile = new System.Windows.Forms.Button();
            this.btnCopy = new System.Windows.Forms.Button();
            this.lblTable = new System.Windows.Forms.Label();
            this.btnMove = new System.Windows.Forms.Button();
            this.btnFile = new System.Windows.Forms.Button();
            this.lblMDBFile = new System.Windows.Forms.Label();
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
            this.groupBox4.Controls.Add(this.btnTableName);
            this.groupBox4.Controls.Add(this.lblNewTable);
            this.groupBox4.Controls.Add(this.lblNewMDBFile);
            this.groupBox4.Location = new System.Drawing.Point(8, 280);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(680, 80);
            this.groupBox4.TabIndex = 32;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "New Database File And Table";
            // 
            // btnTableName
            // 
            this.btnTableName.Location = new System.Drawing.Point(290, 46);
            this.btnTableName.Name = "btnTableName";
            this.btnTableName.Size = new System.Drawing.Size(170, 30);
            this.btnTableName.TabIndex = 2;
            this.btnTableName.Text = "Change Table Name";
            this.btnTableName.Click += new System.EventHandler(this.btnTableName_Click);
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
            this.lblNewMDBFile.BackColor = System.Drawing.Color.White;
            this.lblNewMDBFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblNewMDBFile.Location = new System.Drawing.Point(16, 24);
            this.lblNewMDBFile.Name = "lblNewMDBFile";
            this.lblNewMDBFile.Size = new System.Drawing.Size(648, 16);
            this.lblNewMDBFile.TabIndex = 0;
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
            this.groupBox3.Controls.Add(this.btnCopyToSameDbFile);
            this.groupBox3.Controls.Add(this.btnCopy);
            this.groupBox3.Controls.Add(this.lblTable);
            this.groupBox3.Controls.Add(this.btnMove);
            this.groupBox3.Controls.Add(this.btnFile);
            this.groupBox3.Controls.Add(this.lblMDBFile);
            this.groupBox3.Location = new System.Drawing.Point(8, 104);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(680, 160);
            this.groupBox3.TabIndex = 27;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Current Database File And Table";
            // 
            // btnCopyToSameDbFile
            // 
            this.btnCopyToSameDbFile.Location = new System.Drawing.Point(304, 80);
            this.btnCopyToSameDbFile.Name = "btnCopyToSameDbFile";
            this.btnCopyToSameDbFile.Size = new System.Drawing.Size(360, 24);
            this.btnCopyToSameDbFile.TabIndex = 30;
            this.btnCopyToSameDbFile.Text = "Copy Table To Same Db File And Assign A New Table Name";
            this.btnCopyToSameDbFile.Click += new System.EventHandler(this.btnCopyToSameDbFile_Click);
            // 
            // btnCopy
            // 
            this.btnCopy.Location = new System.Drawing.Point(16, 128);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(280, 24);
            this.btnCopy.TabIndex = 29;
            this.btnCopy.Text = "Copy Table To A Different MS Access Db File";
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
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
            // btnMove
            // 
            this.btnMove.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnMove.Location = new System.Drawing.Point(16, 104);
            this.btnMove.Name = "btnMove";
            this.btnMove.Size = new System.Drawing.Size(280, 24);
            this.btnMove.TabIndex = 27;
            this.btnMove.Text = "Move Table To A Different MS Access Db  File";
            this.btnMove.Click += new System.EventHandler(this.btnMove_Click);
            // 
            // btnFile
            // 
            this.btnFile.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnFile.Location = new System.Drawing.Point(16, 80);
            this.btnFile.Name = "btnFile";
            this.btnFile.Size = new System.Drawing.Size(280, 24);
            this.btnFile.TabIndex = 8;
            this.btnFile.Text = "Get An MS Access Db File And Table";
            this.btnFile.Click += new System.EventHandler(this.btnFile_Click);
            // 
            // lblMDBFile
            // 
            this.lblMDBFile.BackColor = System.Drawing.Color.White;
            this.lblMDBFile.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblMDBFile.Location = new System.Drawing.Point(16, 24);
            this.lblMDBFile.Name = "lblMDBFile";
            this.lblMDBFile.Size = new System.Drawing.Size(648, 16);
            this.lblMDBFile.TabIndex = 0;
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
		#endregion

		private void btnFile_Click(object sender, System.EventArgs e)
		{
			string strFullPath;
			string strDir="";
			string strFile="";
			//string strTemp="";
           // string strAction="";
			string strLargestString="";
			

			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "Get Database File Containing " + this.lblTableType.Text + " Table";
			//OpenFileDialog1.Filter = "Access Database File (*.MDB) |*.mdb";
			OpenFileDialog1.Filter = "MS Access Database File (*.MDB,*.MDE,*.ACCDB) |*.mdb;*.mde;*.accdb";
			strFullPath = this.lblMDBFile.Text.Trim() ;
			OpenFileDialog1.InitialDirectory = strFullPath;
			//strFullPath += this.lblTable.Text;
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK)
			{
				strFullPath = OpenFileDialog1.FileName.Trim();
				if (strFullPath.Length > 0) 
				{
					utils p_utils = new utils();
					strFile = p_utils.getFileName(strFullPath);
					strDir = p_utils.getDirectory(strFullPath);
					p_utils = null;
					dao_data_access tempDao = new dao_data_access();
					tempDao.OpenDb(strFullPath);
					if (tempDao.m_intError == 0) 
					{
						frmDialog frmTemp = new frmDialog();
						frmTemp.Text = "Select " + this.lblTableType.Text + " Table";

						frmTemp.uc_select_list_item1.lblMsg.Text= "Table contents of " + strFullPath;
						frmTemp.uc_select_list_item1.lblMsg.Visible = true;
						strLargestString = frmTemp.uc_select_list_item1.lblMsg.Text;
						//frmTemp.uc_select_list_item1.Dock = System.Windows.Forms.DockStyle.Fill;
						
						frmTemp.uc_project1.Visible=false;
						frmTemp.uc_select_list_item1.listBox1.Items.Clear();
						for (int x=0; x <= tempDao.m_DaoDatabase.TableDefs.Count - 1; x++)
						{
							if (tempDao.m_DaoDatabase.TableDefs[x].Name.IndexOf("MSys",0) < 0) 		
							{
								frmTemp.uc_select_list_item1.listBox1.Items.Add(tempDao.m_DaoDatabase.TableDefs[x].Name);
								if (tempDao.m_DaoDatabase.TableDefs[x].Name.Trim().Length > 
									strLargestString.Trim().Length)
									strLargestString = tempDao.m_DaoDatabase.TableDefs[x].Name;
							}

						}
                        
						
						tempDao.m_DaoDatabase.Close();
						tempDao.m_DaoDatabase = null;
						
                        frmTemp.uc_select_list_item1.Initialize_Width(strLargestString);
						//frmTemp.uc_select_list_item1.groupBox2.Left = 10 ;
						//frmTemp.uc_select_list_item1.groupBox2.Width =  frmTemp.uc_select_list_item1.groupBox1.Width - 20;
						frmTemp.uc_select_list_item1.Visible=true;
						result = frmTemp.ShowDialog(this);
                        
						
						if (result == DialogResult.OK) 
						{
							
							//validation of the table chose will be done here
							this.lblNewMDBFile.Text = strFullPath;
							this.lblNewTable.Text = frmTemp.uc_select_list_item1.listBox1.Text;
							this.strAction="new";
						}
					
						frmTemp.Close();
						frmTemp = null;
					}
					tempDao = null;
				
				

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
				this.lblMDBFile.Width = this.groupBox2.Width - (this.lblMDBFile.Left * 2);
				this.lblNewMDBFile.Width = this.groupBox2.Width - (this.lblMDBFile.Left * 2);
				this.lblTableType.Width = this.lblMDBFile.Width;
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

		private void btnMove_Click(object sender, System.EventArgs e)
		{
		   this.strAction="move";
           this.PerformAction();

		}
		private void btnCommitChange_Click(object sender, System.EventArgs e)
		{
			string strSQL="";
			string strConn = "";
			//string strFullPath ="";
			string strDir = "";
			string strFile = "";

			this.progressBar1.Left = (int)(this.groupBox1.Width * .50) - (int)(this.progressBar1.Width * .50);
			this.progressBar1.Top = this.btnCommitChange.Top + this.btnCommitChange.Height + 10;
			this.lblProgress.Left = this.progressBar1.Left;
			this.lblProgress.Top = this.progressBar1.Top + this.progressBar1.Height + 2;

			dao_data_access tempDao = new dao_data_access();
			ado_data_access tempAdo = new ado_data_access();
            SQLite.ADO.DataMgr tempDataMgr = new SQLite.ADO.DataMgr();
			utils p_utils = new utils();
			switch (strAction)
			{
				case "copy":
                    frmMain.g_oFrmMain.ActivateStandByAnimation(
                       frmMain.g_oFrmMain.WindowState,
                       frmMain.g_oFrmMain.Left,
                       frmMain.g_oFrmMain.Height,
                       frmMain.g_oFrmMain.Width,
                       frmMain.g_oFrmMain.Top);
					//check if the db file exists
					if (System.IO.File.Exists(this.lblNewMDBFile.Text) == false) 
					{
						tempDao.CreateMDB(this.lblNewMDBFile.Text.Trim());
						if (tempDao.m_intError < 0) break;
					}
					//create a table link from the source db file to the destination db file
					tempDao.CreateTableLink(this.lblNewMDBFile.Text.Trim(),
											"temporary_table_link",
											this.lblMDBFile.Text.Trim(),
											this.lblTable.Text.Trim(),true);
					//close the dao workspace
					tempDao.m_DaoWorkspace.Close();
					if (tempDao.m_intError==0)
					{
						//open an ADO connection to the destination DB file
						tempAdo.OpenConnection(tempAdo.getMDBConnString(this.lblNewMDBFile.Text.Trim(),"",""));
						if (tempAdo.m_intError==0)
						{
							if (tempAdo.TableExist(tempAdo.m_OleDbConnection,this.lblNewTable.Text.Trim()))
							{
								//delete the old destination table
								tempAdo.m_strSQL = "DROP TABLE " + this.lblNewTable.Text;
								tempAdo.SqlNonQuery(tempAdo.m_OleDbConnection,tempAdo.m_strSQL);
							}
							//create the new destination table
							tempAdo.m_strSQL = "SELECT * " + 
								"INTO " + this.lblNewTable.Text.Trim() + " " + 
								"FROM temporary_table_link";
							tempAdo.SqlNonQuery(tempAdo.m_OleDbConnection,tempAdo.m_strSQL);

					
							if (tempAdo.m_intError==0)
							{
								//close connection
								tempAdo.CloseConnection(tempAdo.m_OleDbConnection);

								//update the datasource info
								strFile = p_utils.getFileName(this.lblNewMDBFile.Text);
								strDir = p_utils.getDirectory(this.lblNewMDBFile.Text);

                                if (this.m_strScenarioId.Trim().Length > 0)
                                {

                                    strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
                                        "file = '" + strFile + "', " +
                                        "table_name = '" + this.lblNewTable.Text.Trim() + "'" +
                                        " WHERE scenario_id = '" + this.m_strScenarioId + "' AND " +
                                        "table_type = '" + this.lblTableType.Text.Trim() + "';";
                                }
                                else
                                {
                                    strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
                                        "file = '" + strFile + "', " +
                                        "table_name = '" + this.lblNewTable.Text.Trim() + "'" +
                                        " WHERE table_type = '" + this.lblTableType.Text.Trim() + "';";
                                }

                                if (! m_bUsingSqlite)
                                {
                                    strConn = tempAdo.getMDBConnString(this.m_strDataSourceMDBFile, "admin", "");
                                    tempAdo.SqlNonQuery(strConn, strSQL);
                                }
                                else
                                {
                                    strConn = tempDataMgr.GetConnectionString(this.m_strDataSourceMDBFile);
                                    using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
                                    {
                                        con.Open();
                                        tempDataMgr.SqlNonQuery(con, strSQL);
                                    }
                                }
								
							}
							else tempAdo.CloseConnection(tempAdo.m_OleDbConnection);
						}
					}
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();
					break;
				case "move":
                    frmMain.g_oFrmMain.ActivateStandByAnimation(
                      frmMain.g_oFrmMain.WindowState,
                      frmMain.g_oFrmMain.Left,
                      frmMain.g_oFrmMain.Height,
                      frmMain.g_oFrmMain.Width,
                      frmMain.g_oFrmMain.Top);
					//check if the db file exists
					if (System.IO.File.Exists(this.lblNewMDBFile.Text) == false) 
					{
						tempDao.CreateMDB(this.lblNewMDBFile.Text.Trim());
						if (tempDao.m_intError < 0) break;
					}
					//create a table link from the source db file to the destination db file
					tempDao.CreateTableLink(this.lblNewMDBFile.Text.Trim(),
						"temporary_table_link",
						this.lblMDBFile.Text.Trim(),
						this.lblTable.Text.Trim(),true);
					//close the dao workspace
					tempDao.m_DaoWorkspace.Close();
					if (tempDao.m_intError==0)
					{
						//open an ADO connection to the destination DB file
						tempAdo.OpenConnection(tempAdo.getMDBConnString(this.lblNewMDBFile.Text.Trim(),"",""));
						if (tempAdo.m_intError==0)
						{
							if (tempAdo.TableExist(tempAdo.m_OleDbConnection,this.lblNewTable.Text.Trim()))
							{
								//delete the old destination table
								tempAdo.m_strSQL = "DROP TABLE " + this.lblNewTable.Text.Trim();
								tempAdo.SqlNonQuery(tempAdo.m_OleDbConnection,tempAdo.m_strSQL);
							}
							//create the new destination table
							tempAdo.m_strSQL = "SELECT * " + 
								"INTO " + this.lblNewTable.Text.Trim() + " " + 
								"FROM temporary_table_link";
							tempAdo.SqlNonQuery(tempAdo.m_OleDbConnection,tempAdo.m_strSQL);

							//drop the temporary link from the destination DB file
							tempAdo.SqlNonQuery(tempAdo.m_OleDbConnection,"DROP TABLE temporary_table_link");
							tempAdo.CloseConnection(tempAdo.m_OleDbConnection);
					
							//open connection to the source DB file
							tempAdo.OpenConnection(tempAdo.getMDBConnString(this.lblMDBFile.Text.Trim(),"",""));
							if (tempAdo.TableExist(tempAdo.m_OleDbConnection,this.lblTable.Text))
							{
								//delete the source table
								tempAdo.m_strSQL = "DROP TABLE " + this.lblTable.Text.Trim();
								tempAdo.SqlNonQuery(tempAdo.m_OleDbConnection,tempAdo.m_strSQL);
							}
					
							if (tempAdo.m_intError==0)
							{
						
								strFile = p_utils.getFileName(this.lblNewMDBFile.Text.Trim());
								strDir = p_utils.getDirectory(this.lblNewMDBFile.Text.Trim());
                                if (this.m_strScenarioId.Trim().Length > 0)
                                {
                                    strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
                                        "file = '" + strFile + "', " +
                                        "table_name = '" + this.lblNewTable.Text.Trim() + "'" +
                                        " WHERE scenario_id = '" + this.m_strScenarioId + "' AND " +
                                        "table_type = '" + this.lblTableType.Text.Trim() + "';";
                                }
                                else
                                {
                                    strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
                                        "file = '" + strFile + "', " +
                                        "table_name = '" + this.lblNewTable.Text.Trim() + "'" +
                                        " WHERE table_type = '" + this.lblTableType.Text.Trim() + "';";
                                }

                                if (!m_bUsingSqlite)
                                {
                                    strConn = tempAdo.getMDBConnString(this.m_strDataSourceMDBFile.Trim(), "admin", "");
                                    tempAdo.SqlNonQuery(strConn, strSQL);
                                }
                                else
                                {
                                    strConn = tempDataMgr.GetConnectionString(this.m_strDataSourceMDBFile);
                                    using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
                                    {
                                        con.Open();
                                        tempDataMgr.SqlNonQuery(con, strSQL);
                                    }
                                }
                            }
						}
					}
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();
					break;
				case "new":
					
					strFile = p_utils.getFileName(this.lblNewMDBFile.Text.Trim());
					strDir = p_utils.getDirectory(this.lblNewMDBFile.Text.Trim());
											
					strConn = tempAdo.getMDBConnString(this.m_strDataSourceMDBFile.Trim(),"admin","");	

					if (this.m_strScenarioId.Trim().Length > 0)
					{

						strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
							"file = '" +  strFile + "', " +
							"table_name = '" + this.lblNewTable.Text.Trim() + "'" + 
							" WHERE scenario_id = '" + 	this.m_strScenarioId + "' AND " +
							"table_type = '" + this.lblTableType.Text.Trim() + "';";
					}
					else
					{
						strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
							"file = '" +  strFile + "', " +
							"table_name = '" + this.lblNewTable.Text.Trim() + "'" + 
							" WHERE table_type = '" + this.lblTableType.Text.Trim() + "';";
					}

                    if (!m_bUsingSqlite)
                    {
                        tempAdo.SqlNonQuery(strConn, strSQL);
                    }
                    else
                    {
                        strConn = tempDataMgr.GetConnectionString(this.m_strDataSourceMDBFile);
                        using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
                        {
                            con.Open();
                            tempDataMgr.SqlNonQuery(con, strSQL);
                        }
                    }

					break;
				case "copytosamedbfile":
                    frmMain.g_oFrmMain.ActivateStandByAnimation(
                     frmMain.g_oFrmMain.WindowState,
                     frmMain.g_oFrmMain.Left,
                     frmMain.g_oFrmMain.Height,
                     frmMain.g_oFrmMain.Width,
                     frmMain.g_oFrmMain.Top);
					//open connection to the source DB file
					tempAdo.OpenConnection(tempAdo.getMDBConnString(this.lblMDBFile.Text.Trim(),"",""));
					if (tempAdo.m_intError==0)
					{
						if (tempAdo.TableExist(tempAdo.m_OleDbConnection,this.lblNewTable.Text.Trim()))
						{
							DialogResult result = MessageBox.Show("Overwrite " + this.lblNewTable.Text + "?(Y/N)","FIA_Biosum",System.Windows.Forms.MessageBoxButtons.YesNo,System.Windows.Forms.MessageBoxIcon.Question);
							if (result == DialogResult.Yes)
							{
								//delete the source table
								tempAdo.m_strSQL = "DROP TABLE " + this.lblNewTable.Text;
								tempAdo.SqlNonQuery(tempAdo.m_OleDbConnection,tempAdo.m_strSQL);

								//copy table to new table
								tempAdo.m_strSQL = "SELECT * INTO " + this.lblNewTable.Text + " " + 
									"FROM " + this.lblTable.Text;
								tempAdo.SqlNonQuery(tempAdo.m_OleDbConnection,tempAdo.m_strSQL);

							}						
						}
						else
						{
							//copy table to new table
							tempAdo.m_strSQL = "SELECT * INTO " + this.lblNewTable.Text + " " + 
								"FROM " + this.lblTable.Text;
							tempAdo.SqlNonQuery(tempAdo.m_OleDbConnection,tempAdo.m_strSQL);
						}

						strFile = p_utils.getFileName(this.lblNewMDBFile.Text.Trim());
						strDir = p_utils.getDirectory(this.lblNewMDBFile.Text.Trim());

                        if (this.m_strScenarioId.Trim().Length > 0)
                        {

                            strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
                                "file = '" + strFile + "', " +
                                "table_name = '" + this.lblNewTable.Text.Trim() + "'" +
                                " WHERE scenario_id = '" + this.m_strScenarioId + "' AND " +
                                "table_type = '" + this.lblTableType.Text.Trim() + "';";
                        }
                        else
                        {
                            strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
                                "file = '" + strFile + "', " +
                                "table_name = '" + this.lblNewTable.Text.Trim() + "'" +
                                " WHERE table_type = '" + this.lblTableType.Text.Trim() + "';";
                        }

                        if (!m_bUsingSqlite)
                        {
                            strConn = tempAdo.getMDBConnString(this.m_strDataSourceMDBFile.Trim(), "admin", "");
                            tempAdo.SqlNonQuery(strConn, strSQL);
                        }
                        else
                        {
                            strConn = tempDataMgr.GetConnectionString(this.m_strDataSourceMDBFile);
                            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
                            {
                                con.Open();
                                tempDataMgr.SqlNonQuery(con, strSQL);
                            }
                        }

					}
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();

					break;
				default:
					break;
			}
			if (this.lblNewMDBFile.Text.Length == 0 || 
				this.lblNewTable.Text.Length == 0) 
			{
				MessageBox.Show("Not Valid: New MDB File and/or Table are blank");
			}
			else 
			{
				if (tempDao.m_intError == 0 && tempAdo.m_intError == 0) 
				{
					//((frmScenario)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.OK;	
					this.ParentForm.DialogResult = System.Windows.Forms.DialogResult.OK;
				}
			}
			tempDao = null;
			tempAdo = null;
            tempDataMgr = null;
			p_utils = null;			
		}

		private void btnCommitChange_Click_Old(object sender, System.EventArgs e)
		{
            string strSQL="";
			string strConn = "";
			//string strFullPath ="";
			string strDir = "";
			string strFile = "";

			this.progressBar1.Left = (int)(this.groupBox1.Width * .50) - (int)(this.progressBar1.Width * .50);
			this.progressBar1.Top = this.btnCommitChange.Top + this.btnCommitChange.Height + 10;
			this.lblProgress.Left = this.progressBar1.Left;
			this.lblProgress.Top = this.progressBar1.Top + this.progressBar1.Height + 2;

			dao_data_access tempDao = new dao_data_access();
			ado_data_access tempAdo = new ado_data_access();
			utils p_utils = new utils();
			switch (strAction)
			{
				case "copy":
					if (System.IO.File.Exists(this.lblNewMDBFile.Text) == false) 
					{
						tempDao.CreateMDB(this.lblNewMDBFile.Text);
						if (tempDao.m_intError < 0) break;
					}
					else
					{
						if (this.m_bOverwriteTable==true)
						{
							tempDao.DeleteTableFromMDB(this.lblNewMDBFile.Text.Trim(),this.lblNewTable.Text.Trim());
						}
					}
					tempDao.MoveTableToMDB(this,false);
					if (tempDao.m_intError == 0) 
					{
						
						strFile = p_utils.getFileName(this.lblNewMDBFile.Text);
						strDir = p_utils.getDirectory(this.lblNewMDBFile.Text);
						
						
						strConn = tempAdo.getMDBConnString(this.m_strDataSourceMDBFile,"admin","");
						//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + 
						//	this.m_strDataSourceMDBFile + ";User Id=admin;Password=;";

						if (this.m_strScenarioId.Trim().Length > 0)
						{

							strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
								"file = '" +  strFile + "', " +
								"table_name = '" + this.lblNewTable.Text.Trim() + "'" + 
								" WHERE scenario_id = '" + 	this.m_strScenarioId + "' AND " +
								"table_type = '" + this.lblTableType.Text.Trim() + "';";
						}
						else
						{
							strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
								"file = '" +  strFile + "', " +
								"table_name = '" + this.lblNewTable.Text.Trim() + "'" + 
								" WHERE table_type = '" + this.lblTableType.Text.Trim() + "';";
						}
						
						tempAdo.SqlNonQuery(strConn, strSQL);
					}

					break;
				case "move":
					if (System.IO.File.Exists(this.lblNewMDBFile.Text) == false) 
					{
						tempDao.CreateMDB(this.lblNewMDBFile.Text);
						if (tempDao.m_intError < 0) break;
					}
					else
					{
						if (this.m_bOverwriteTable==true)
						{
							tempDao.DeleteTableFromMDB(this.lblNewMDBFile.Text.Trim(),this.lblNewTable.Text.Trim());
						}
					}
                    //dao_data_access tempDao = new dao_data_access();
					tempDao.MoveTableToMDB(this,true);
					if (tempDao.m_intError == 0) 
					{
						
						strFile = p_utils.getFileName(this.lblNewMDBFile.Text);
						strDir = p_utils.getDirectory(this.lblNewMDBFile.Text);
						
						
						strConn = tempAdo.getMDBConnString(this.m_strDataSourceMDBFile,"admin","");
						//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + 
						//	this.m_strDataSourceMDBFile + ";User Id=admin;Password=;";

						if (this.m_strScenarioId.Trim().Length > 0)
						{

							strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
								"file = '" +  strFile + "', " +
								"table_name = '" + this.lblNewTable.Text.Trim() + "'" + 
								" WHERE scenario_id = '" + 	this.m_strScenarioId + "' AND " +
								"table_type = '" + this.lblTableType.Text.Trim() + "';";
						}
						else
						{
							strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
								"file = '" +  strFile + "', " +
								"table_name = '" + this.lblNewTable.Text.Trim() + "'" + 
								" WHERE table_type = '" + this.lblTableType.Text.Trim() + "';";
						}
						
						tempAdo.SqlNonQuery(strConn, strSQL);
					}
					break;
				case "new":
					
					strFile = p_utils.getFileName(this.lblNewMDBFile.Text);
					strDir = p_utils.getDirectory(this.lblNewMDBFile.Text);
					
						
					strConn = tempAdo.getMDBConnString(this.m_strDataSourceMDBFile,"admin","");	
					//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + 
					//	this.m_strDataSourceMDBFile + ";User Id=admin;Password=;";

					if (this.m_strScenarioId.Trim().Length > 0)
					{

						strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
							"file = '" +  strFile + "', " +
							"table_name = '" + this.lblNewTable.Text.Trim() + "'" + 
							" WHERE scenario_id = '" + 	this.m_strScenarioId + "' AND " +
							"table_type = '" + this.lblTableType.Text.Trim() + "';";
					}
					else
					{
						strSQL = "UPDATE " + this.m_strDataSourceTable + " SET path = '" + strDir + "', " +
							"file = '" +  strFile + "', " +
							"table_name = '" + this.lblNewTable.Text.Trim() + "'" + 
							" WHERE table_type = '" + this.lblTableType.Text.Trim() + "';";
					}
						
					tempAdo.SqlNonQuery(strConn, strSQL);

					break;

                default:
					break;
			}
			if (this.lblNewMDBFile.Text.Length == 0 || 
				this.lblNewTable.Text.Length == 0) 
			{
				MessageBox.Show("Not Valid: New MDB File and/or Table are blank");
			}
			else 
			{
				if (tempDao.m_intError == 0 && tempAdo.m_intError == 0) 
				{
					//((frmScenario)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.OK;	
					this.ParentForm.DialogResult = System.Windows.Forms.DialogResult.OK;
				}
			}
		    tempDao = null;
			tempAdo = null;
			p_utils = null;

			
		}

		private void btnCopy_Click(object sender, System.EventArgs e)
		{
			//copy will not remove the table copied from the Source MDB File
			strAction="copy";
			this.PerformAction();
			
		}
		private void PerformAction()
		{
			string strMDBFileDestination;    //move the table to this MDB file
			string strMDBFileSource;
			
			string strDir="";
			string strFile="";
			string strTableDestination = ""; //name to give moved table
			string strTableSource = "";
			//string strTemp="";
			//int x;
			int y;
			
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "Get Destination Database To " + this.strAction + " " + this.lblTableType.Text + " Table";
			//OpenFileDialog1.Filter = "Access Database File (*.MDB) |*.mdb";
			OpenFileDialog1.Filter = "MS Access Database File (*.MDB,*.MDE,*.ACCDB) |*.mdb;*.mde;*.accdb";
			strMDBFileDestination = this.lblMDBFile.Text.Trim() ;
			OpenFileDialog1.InitialDirectory = strMDBFileDestination;
			OpenFileDialog1.CheckFileExists = false;
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK)
			{
				strMDBFileDestination = OpenFileDialog1.FileName.Trim();
				if (strMDBFileDestination.Length > 0) 
				{
					utils p_utils = new utils();
					strFile = p_utils.getFileName(strMDBFileDestination);
					strDir = p_utils.getDirectory(strMDBFileDestination);
					p_utils = null;
					
					dao_data_access tempDao = new dao_data_access();
					
					strTableSource = this.lblTable.Text.Trim();
					strTableDestination = strTableSource;
					strMDBFileSource = this.lblMDBFile.Text.ToString().Trim();
					//see if the mdb file exists
					if (System.IO.File.Exists(strMDBFileDestination) == true) 
					{
						//open the mdb file
						tempDao.OpenDb(strMDBFileDestination);
						if (tempDao.m_intError == 0) 
						{
							
							//see if the table name already exists within the mdb table
							if (tempDao.TableExists(tempDao.m_DaoDatabase,strTableDestination) == true )
							{
								FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
								frmTemp.MaximizeBox = false;
								frmTemp.BackColor = System.Drawing.SystemColors.Control;
								frmTemp.Text = "Table: " + strTableDestination + " Exists";

								FIA_Biosum_Manager.uc_table_exists_dialog p_uc = new uc_table_exists_dialog();
								frmTemp.Controls.Add(p_uc);
                                p_uc.strOkButtonText = "OK";
           
			

								frmTemp.Height=0;
								frmTemp.Width=0;
								if (p_uc.Top + p_uc.Height > frmTemp.ClientSize.Height + 2)
								{
									for (y=1;;y++)
									{
										frmTemp.Height = y;
										if (p_uc.Top + 
											p_uc.Height < 
											frmTemp.ClientSize.Height)
										{
											break;
										}
									}

								}
								if (p_uc.Left + p_uc.Width > frmTemp.ClientSize.Width + 2)
								{
									for (y=1;;y++)
									{
										frmTemp.Width = y;
										if (p_uc.Left + 
											p_uc.Width < 
											frmTemp.ClientSize.Width)
										{
											break;
										}
									}

								}
								frmTemp.Left = 0;
								frmTemp.Top = 0;

								p_uc.listBox1.Items.Clear();
								p_uc.strMDBFileLabel = "Table contents of " + strMDBFileDestination.Trim();
								tempDao.LoadTablesIntoListBox(tempDao.m_DaoDatabase,p_uc.listBox1);
								p_uc.strTable=strTableDestination;
								frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
								frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
								for (;;)
								{
									this.m_bOverwriteTable=false;
									strTableDestination=p_uc.strTable;
									result = frmTemp.ShowDialog();
									if (result==System.Windows.Forms.DialogResult.OK)
									{
										
										if (p_uc.rdoOverwrite.Checked==true)
										{
											if (strMDBFileDestination.Trim().ToUpper()==
												strMDBFileSource.Trim().ToUpper())
												
											{
												MessageBox.Show("!!Cannot have source file and table equal to the destination file and table!!","Datasource",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
											}
											else
											{
												this.m_bOverwriteTable = true;
												this.lblNewMDBFile.Text = strMDBFileDestination;
												this.lblNewTable.Text = strTableDestination;
												break;
											}
										}
										else if (tempDao.TableExists(tempDao.m_DaoDatabase,p_uc.strTable) == false)
										{
											this.lblNewMDBFile.Text = strMDBFileDestination;
											this.lblNewTable.Text = p_uc.strTable;
											break;
										}
										else
										{
											frmTemp.Text = "Table:" + p_uc.strTable + " Exists";
										}
									}
									else
									{
										strTableDestination="";
										break;
									}
								}
								frmTemp.Dispose();
								
							}
							else 
							{
							    
								this.lblNewMDBFile.Text = strMDBFileDestination;
								this.lblNewTable.Text = strTableDestination;
							}
							tempDao.m_DaoDatabase.Close();
						}
						
					}
					else 
					{
						this.lblNewMDBFile.Text = strMDBFileDestination;
						this.lblNewTable.Text = this.lblTable.Text;
					}  
					tempDao = null;
				}
				else 
				{
					this.strAction="";
				}

			}
			else
			{
				this.strAction="";
			}

		}

		private void btnTableName_Click(object sender, System.EventArgs e)
		{
			int intTop;
			int intLeft;
			frmDialog frmTemp = new frmDialog();
			System.Windows.Forms.TextBox p_txtString = new TextBox();
		                    
								

			//define form properties
			frmTemp.Width = 200;
			frmTemp.Height = 200;
			frmTemp.WindowState = System.Windows.Forms.FormWindowState.Normal;
			frmTemp.MaximizeBox = false;
			frmTemp.MinimizeBox = false;
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.Text = "Table Name";

			//define numeric text class properties
			p_txtString.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			p_txtString.Name = "txtTable";
			p_txtString.TabIndex = 0;
			p_txtString.Tag = "";
			p_txtString.Visible = true;
			p_txtString.Enabled = true;
			frmTemp.Controls.Add(p_txtString);  //add the text box to the form
			p_txtString.Height = 100;

			p_txtString.Top = (int)(frmTemp.ClientSize.Height * .50) - (int)(p_txtString.Height * .50);
			p_txtString.Width = frmTemp.ClientSize.Width - 20;
			p_txtString.Left = 10;

			p_txtString.ReadOnly=false;
			p_txtString.Text = this.lblNewTable.Text ;

			intLeft = p_txtString.Left;
			intTop = p_txtString.Top;
			//frmTemp.m_txtNumeric = p_txtDefault;
			frmTemp.txtBox = p_txtString;
					//-----------label---------------//
			System.Windows.Forms.Label lblTemp = new Label();
			lblTemp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			lblTemp.Location = new System.Drawing.Point(168, 53);
			lblTemp.Name = "lblTemp";
			lblTemp.Size = new System.Drawing.Size(128, 15);
			lblTemp.TabIndex = 1;
			lblTemp.Text = "New Value";
			frmTemp.Controls.Add(lblTemp);
			lblTemp.Top = intTop - lblTemp.Height - 10;
			lblTemp.Left = intLeft;
			lblTemp.Visible=true;


                    //----------------OK button-----------------//
			System.Windows.Forms.Button btnTempOK = new Button();
			btnTempOK.Location = new System.Drawing.Point(392, 328);
			btnTempOK.Name = "btnOK";
			btnTempOK.Size = new System.Drawing.Size(80, 32);
			btnTempOK.TabIndex = 2;
			btnTempOK.Text = "OK";
			btnTempOK.Click += new System.EventHandler(frmTemp.btnOK_Click);
			frmTemp.Controls.Add(btnTempOK);
			btnTempOK.Top = intTop + btnTempOK.Height + 10;
			btnTempOK.Left = (int)(frmTemp.Width * .50)  - btnTempOK.Width;
			btnTempOK.Visible=true;

					//----------------Cancel button-----------------//
			System.Windows.Forms.Button btnTempCancel = new Button();
			btnTempCancel.Location = new System.Drawing.Point(392, 328);
			btnTempCancel.Name = "btnCancel";
			btnTempCancel.Size = new System.Drawing.Size(80, 32);
			btnTempCancel.TabIndex = 2;
			btnTempCancel.Text = "CANCEL";
			btnTempCancel.Click += new System.EventHandler(frmTemp.btnCancel_Click);
			frmTemp.Controls.Add(btnTempCancel);
			btnTempCancel.Top = intTop + btnTempCancel.Height + 10;
			btnTempCancel.Left = (int)(frmTemp.Width * .50) ;
			btnTempCancel.Visible=true;
                    //frmTemp.Top = (int)(this.ParentForm.ClientSize.Height * .50) + (int)(frmTemp.Height * .50) + this.ParentForm.Top;
					//frmTemp.Left = (int)(this.ParentForm.ClientSize.Width * .50) + (int)(frmTemp.Width * .50) + this.ParentForm.Left;
			frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			DialogResult result = frmTemp.ShowDialog();
					//MessageBox.Show(dlgResult.ToString());
			if (result.ToString().Trim().ToUpper() == "OK" )
			{
				string strTable = p_txtString.Text.Trim();
				frmTemp.Dispose();
				if (System.IO.File.Exists(this.lblNewMDBFile.Text.Trim()) == true) 
				{
					dao_data_access p_dao = new dao_data_access();
					//open the mdb file
					p_dao.OpenDb(this.lblNewMDBFile.Text.Trim());
					if (p_dao.m_intError == 0) 
					{
							
						//see if the table name already exists within the mdb table
						if (p_dao.TableExists(p_dao.m_DaoDatabase,strTable) == true )
						{
							int y;
							FIA_Biosum_Manager.frmDialog frmDlg = new frmDialog();
							frmDlg.MaximizeBox = false;
							frmDlg.BackColor = System.Drawing.SystemColors.Control;
							frmDlg.Text = "Table: " + strTable + " Exists";
							FIA_Biosum_Manager.uc_table_exists_dialog p_uc = new uc_table_exists_dialog();
							frmDlg.Controls.Add(p_uc);
							p_uc.strOkButtonText = "OK";
           
			

							frmDlg.Height=0;
							frmDlg.Width=0;
							if (p_uc.Top + p_uc.Height > frmDlg.ClientSize.Height + 2)
							{
								for (y=1;;y++)
								{
									frmDlg.Height = y;
									if (p_uc.Top + 
										p_uc.Height < 
										frmDlg.ClientSize.Height)
									{
										break;
									}
								}
							}
							if (p_uc.Left + p_uc.Width > frmDlg.ClientSize.Width + 2)
							{
								for (y=1;;y++)
								{
									frmDlg.Width = y;
									if (p_uc.Left + 
										p_uc.Width < 
										frmDlg.ClientSize.Width)
									{
										break;
									}
								}
							}
							frmDlg.Left = 0;
							frmDlg.Top = 0;

							p_uc.listBox1.Items.Clear();
							p_uc.strMDBFileLabel = "Table contents of " + this.lblNewMDBFile.Text.Trim();
							p_dao.LoadTablesIntoListBox(p_dao.m_DaoDatabase,p_uc.listBox1);
							p_uc.strTable=strTable;
							frmDlg.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
							frmDlg.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
							for (;;)
							{
								this.m_bOverwriteTable=false;
								strTable=p_uc.strTable;
								result = frmDlg.ShowDialog();
								if (result==System.Windows.Forms.DialogResult.OK)
								{
									
									if (p_uc.rdoOverwrite.Checked==true)
									{
										if (this.lblNewMDBFile.Text.Trim().ToUpper()==
											this.lblMDBFile.Text.Trim().ToUpper())
												
										{
											MessageBox.Show("!!Cannot have source file and table equal to the destination file and table!!","Datasource",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
										}
										else
										{
											this.m_bOverwriteTable = true;
											this.lblNewTable.Text = strTable;
											break;
										}
									}
									else if (p_dao.TableExists(p_dao.m_DaoDatabase,p_uc.strTable) == false)
									{
										this.lblNewTable.Text = p_uc.strTable;
										break;
									}
									else
									{
										frmTemp.Text = "Table:" + p_uc.strTable + " Exists";
									}
								}
								else
								{
									break;
								}
							}
							frmDlg.Dispose();
							frmDlg=null;
						}
						else 
						{
							    
							this.lblNewTable.Text = strTable;
						}
						p_dao.m_DaoDatabase.Close();
						
					}
					
					p_dao=null;
				}
				else
				{
					this.lblNewTable.Text = strTable;
				}
			}
			else
			{
				frmTemp.Dispose();
			}
			frmTemp=null;
		}

		private void btnCopyToSameDbFile_Click(object sender, System.EventArgs e)
		{
			//open connection to the source DB file
			string strTableName = this.lblTable.Text.Trim();
			FIA_Biosum_Manager.ado_data_access oAdo = new ado_data_access();
			oAdo.OpenConnection(oAdo.getMDBConnString(this.lblMDBFile.Text.Trim(),"",""));
		
			for (int x=1;x<=100-1;x++)
			{
				if (oAdo.TableExist(oAdo.m_OleDbConnection,strTableName + x.ToString().Trim())==false)
				{
					strTableName = strTableName + x.ToString().Trim();
					break;
				}
			}
			this.lblNewMDBFile.Text = this.lblMDBFile.Text.Trim();
			this.lblNewTable.Text = strTableName;
		    this.strAction="copytosamedbfile";

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

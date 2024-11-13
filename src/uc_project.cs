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
	/// Summary description for uc_project.
	/// </summary>
	public class uc_project : System.Windows.Forms.UserControl
	{

		//oledb
		private System.Data.DataSet dataSet1;
		private System.Data.OleDb.OleDbCommand oleDbSelectCommand1;
		private System.Data.OleDb.OleDbCommand oleDbInsertCommand1;
		private System.Data.OleDb.OleDbCommand oleDbUpdateCommand1;
		private System.Data.OleDb.OleDbCommand oleDbDeleteCommand1;
		private System.Data.OleDb.OleDbDataAdapter oleDbDataAdapter1;
		private System.Data.OleDb.OleDbCommand oleDbCommand1;
		private System.Data.OleDb.OleDbConnection oleDbConnection1;
		//private System.Data.OleDb.OleDbException oleException;

		//project variables
		public bool boolProjectOpen = false;
        public string m_strDebugFile = "";
		//new project variables
		public string m_strNewProjectFile = "";
		public string m_strNewProjectDirectory = "";
		public string m_strNewProjectId="";
        public string m_strNewName="";
		public string m_strNewDate="";
		public string m_strNewCompany="";
		public string m_strNewDescription="";
        public string m_strNewShared="";
		public string m_strNewRootDirectory="";
		public string m_strNewProjectVersion="";
        


		//current open project
		public string m_strProjectId="";
		public string m_strProjectFile="";
		public string m_strProjectVersion="";
		
		public string m_strDBProjectDirectory="";
		public string m_strDBProjectDirectoryDrive="";
		public string m_strDBProjectRootDirectory="";

		public string m_strProjectDirectory="";
		public string m_strProjectDirectoryDrive = "";
		public string m_strProjectRootDirectory="";
		


		//shared user directory for keeping notes and document links
		public string m_strDBSharedDirectory="";
		public string m_strDBSharedRootDirectory="";
		public string m_strDBSharedDirectoryDrive="";
		public string m_strSharedDirectory="";
		public string m_strSharedRootDirectory="";
		public string m_strSharedDirectoryDrive="";

		//personal user directory for keeping notes and document links
		public string m_strDBPersonalDirectory="";
		public string m_strDBPersonalRootDirectory="";
		public string m_strDBPersonalDirectoryDrive="";

		public string m_strPersonalDirectoryDrive="";
        public string m_strPersonalRootDirectory="";
		public string m_strPersonalDirectory="";

		public int m_intError;
		public string m_strError;
		public string m_strAction;
        public int m_intFullHt = 0;
        public int m_intFullWh = 0;

        private FIA_Biosum_Manager.frmDialog m_frmDialog1;
		private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultDatabaseXPSFile;
        private string m_pdfFile = @"Help\DATABASE_Help.pdf";

        public FIA_Biosum_Manager.ResizeFormUsingVisibleScrollBars m_oResizeForm = new ResizeFormUsingVisibleScrollBars();
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox grpboxDescription;
		public System.Windows.Forms.TextBox txtDescription;
		private System.Windows.Forms.GroupBox grpboxCompany;
		public System.Windows.Forms.TextBox txtCompany;
		private System.Windows.Forms.GroupBox grpboxProjectId;
		public System.Windows.Forms.TextBox txtProjectId;
		private System.Windows.Forms.GroupBox grpboxCreated;
		public System.Windows.Forms.TextBox txtDate;
		private System.Windows.Forms.Label lblDate;
		public System.Windows.Forms.TextBox txtName;
		private System.Windows.Forms.Label lblName;
		private System.Windows.Forms.GroupBox grpboxProjectFiles;
		private System.Windows.Forms.Button btnPersonalHelp;
		private System.Windows.Forms.Button btnPersonalDirectory;
		private System.Windows.Forms.Label lblLocal;
		public System.Windows.Forms.TextBox txtPersonal;
		private System.Windows.Forms.Button btnSharedHelp;
		private System.Windows.Forms.Button btnSharedDirectory;
		private System.Windows.Forms.Label lblShared;
		private System.Windows.Forms.Button btnRootDirectoryHelp;
		private System.Windows.Forms.Button btnRootDirectory;
		private System.Windows.Forms.Label lblRootDirectory;
		public System.Windows.Forms.TextBox txtRootDirectory;
		public System.Windows.Forms.TextBox txtShared;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Button btnEdit;
		public System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button btnHelp;
		

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_project()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

            this.m_oEnv = new env();
			// TODO: Add any initialization after the InitializeComponent call
			this.dataSet1 = new System.Data.DataSet();
			this.oleDbSelectCommand1 = new System.Data.OleDb.OleDbCommand();
			this.oleDbInsertCommand1 = new System.Data.OleDb.OleDbCommand();
			this.oleDbUpdateCommand1 = new System.Data.OleDb.OleDbCommand();
			this.oleDbDeleteCommand1 = new System.Data.OleDb.OleDbCommand();
			this.oleDbDataAdapter1 = new System.Data.OleDb.OleDbDataAdapter();
			this.oleDbCommand1 = new System.Data.OleDb.OleDbCommand();
			this.oleDbConnection1 = new System.Data.OleDb.OleDbConnection();
			this.m_intError = 0;
			this.m_strError="";
            // Set the control height to accomodate the lowest button
            this.m_intFullHt = this.btnClose.Location.Y + this.btnClose.Size.Height + 10;
            // Set the control width to accomodate the furthest right groupbox
            this.m_intFullWh = this.grpboxCreated.Location.X + this.grpboxCreated.Size.Width + 30;

			this.txtRootDirectory.Enabled=false;
			this.txtShared.Enabled=false;
			m_oResizeForm.ScrollBarParentControl=panel1;

            m_strDebugFile = frmMain.g_oEnv.strTempDir + @"\FIA_Biosum_DebugLog_" + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".txt";


		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>C:\FIA_BIOSUM\source\cs\fia_biosum\fia_biosum_manager\frmScenario.cs.bak
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
		public void OpenProjectTable(string strRootDir, string strFile)
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//uc_project.OpenProjectTable \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.OpenProjectTable: strRootDir=" + strRootDir + " strFile=" + strFile + "\r\n");
			frmMain.g_sbpInfo.Text = "Loading Project...Stand By";
            this.m_intError = 0;
			this.m_strError = "";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.OpenProjectTable: Instantiate ado_data_access \r\n");

			ado_data_access p_ado = new ado_data_access();
			
			string strFullPath = strRootDir + "\\DB\\" + strFile;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.OpenProjectTable: strFullPath=" + strFullPath + "\r\n");

			string strConn=p_ado.getMDBConnString(strFullPath,"admin","");

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.OpenProjectTable: Open DBFile with Connection String=" + strConn + "\r\n");

			p_ado.OpenConnection(strConn);

           
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.OpenProjectTable: OpenConnection error Value=" + p_ado.m_intError.ToString() + "\r\n");
			if (p_ado.m_intError==0)
			{
				try
				{
                    
					bool bAppVerColumnExist = p_ado.ColumnExist(p_ado.m_OleDbConnection,"project","application_version");
					//string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFullPath + ";User Id=admin;Password=;";
					p_ado.SqlQueryReader(p_ado.m_OleDbConnection,"select * from project");

					p_ado.m_OleDbDataReader.Read();
                    // Adding trim function to remove extra spaces
					m_strNewProjectId=p_ado.m_OleDbDataReader["proj_id"].ToString().Trim();
                    m_strNewName=p_ado.m_OleDbDataReader["created_by"].ToString().Trim();
				    m_strNewDate=p_ado.m_OleDbDataReader["created_date"].ToString();
		            m_strNewCompany=p_ado.m_OleDbDataReader["company"].ToString().Trim();
		            m_strNewDescription=p_ado.m_OleDbDataReader["description"].ToString().Trim();
                    m_strNewShared=p_ado.m_OleDbDataReader["shared_file"].ToString();
		            m_strNewRootDirectory=p_ado.m_OleDbDataReader["project_root_directory"].ToString();
                    
					if (bAppVerColumnExist)
					{
						if (p_ado.m_OleDbDataReader["application_version"] != System.DBNull.Value)
							this.m_strNewProjectVersion = p_ado.m_OleDbDataReader["application_version"].ToString().Trim();
						else
							this.m_strNewProjectVersion="";
					}
					else
					{
						this.m_strNewProjectVersion="";
					}

				
				}
				catch (Exception caught)
				{
					MessageBox.Show(caught.Message);
				}
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbDataReader = null;
				p_ado.m_OleDbCommand = null;
				p_ado.m_OleDbConnection.Close();
				p_ado.m_OleDbConnection = null;
			}
			else 
			{
				this.m_intError = p_ado.m_intError;
                this.m_strError = p_ado.m_strError;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.OpenProjectTable: !!Failed to open project file!! Error=" + m_strError + "\r\n");

			}
            p_ado = null;
			//m_strProjectId = this.txtProjectId.Text.ToString();  
			if (this.m_strAction=="VIEW") 
			{
				this.grpboxDescription.Enabled=false;
				this.grpboxProjectFiles.Enabled=false;
				this.grpboxProjectId.Enabled=false;
				this.grpboxCompany.Enabled=false;
				this.grpboxCreated.Enabled=false;
				this.btnEdit.Enabled=true;
				this.btnCancel.Enabled=false;
                this.btnSave.Enabled=false;
                
			}
			frmMain.g_sbpInfo.Text = "Ready";
		}
		public void OpenUserConfigTable(string strDir, string strFile)
       	{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//uc_project.OpenUserConfigTable \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			this.m_intError = 0;
			this.m_strError = "";
			ado_data_access p_ado = new ado_data_access();
			
			string strFullPath = strDir + "\\" + strFile;
			string strConn = p_ado.getMDBConnString(strFullPath,"admin","");
			//string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFullPath + ";User Id=admin;Password=;";
			p_ado.SqlQueryReader(strConn,"select * from user_config where trim(ucase(user_name)) = '" + System.Environment.UserName.ToString().Trim().ToUpper() + "'");
		
			if (p_ado.m_intError==0)
			{
				try
				{
					while (p_ado.m_OleDbDataReader.Read())
					{
                       this.txtPersonal.Text = p_ado.m_OleDbDataReader["personal_directory"].ToString();
						if (this.txtPersonal.Text.Trim().Length > 0) 
						{
							if (((frmMain)this.ParentForm.ParentForm).tlbMain.Buttons[3].Enabled==false)
							{
								((frmMain)this.ParentForm.ParentForm).tlbMain.Buttons[3].Enabled=true;
							}
						}
						else
						{
							if (((frmMain)this.ParentForm.ParentForm).tlbMain.Buttons[3].Enabled==true)
							{
								((frmMain)this.ParentForm.ParentForm).tlbMain.Buttons[3].Enabled=false;
							}
						}
					}
				}
				catch (Exception caught)
				{
					MessageBox.Show(caught.Message);
				}
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbDataReader = null;
				p_ado.m_OleDbCommand = null;
				p_ado.m_OleDbConnection.Close();
				p_ado.m_OleDbConnection = null;
			}
			else 
			{
				this.m_intError = p_ado.m_intError;
			}
			p_ado = null;
		}
		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_project));
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.grpboxDescription = new System.Windows.Forms.GroupBox();
            this.txtDescription = new System.Windows.Forms.TextBox();
            this.grpboxCompany = new System.Windows.Forms.GroupBox();
            this.txtCompany = new System.Windows.Forms.TextBox();
            this.grpboxProjectId = new System.Windows.Forms.GroupBox();
            this.txtProjectId = new System.Windows.Forms.TextBox();
            this.grpboxCreated = new System.Windows.Forms.GroupBox();
            this.txtDate = new System.Windows.Forms.TextBox();
            this.lblDate = new System.Windows.Forms.Label();
            this.txtName = new System.Windows.Forms.TextBox();
            this.lblName = new System.Windows.Forms.Label();
            this.grpboxProjectFiles = new System.Windows.Forms.GroupBox();
            this.btnPersonalHelp = new System.Windows.Forms.Button();
            this.btnPersonalDirectory = new System.Windows.Forms.Button();
            this.lblLocal = new System.Windows.Forms.Label();
            this.txtPersonal = new System.Windows.Forms.TextBox();
            this.btnSharedHelp = new System.Windows.Forms.Button();
            this.btnSharedDirectory = new System.Windows.Forms.Button();
            this.lblShared = new System.Windows.Forms.Label();
            this.btnRootDirectoryHelp = new System.Windows.Forms.Button();
            this.btnRootDirectory = new System.Windows.Forms.Button();
            this.lblRootDirectory = new System.Windows.Forms.Label();
            this.txtRootDirectory = new System.Windows.Forms.TextBox();
            this.txtShared = new System.Windows.Forms.TextBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnHelp = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.grpboxDescription.SuspendLayout();
            this.grpboxCompany.SuspendLayout();
            this.grpboxProjectId.SuspendLayout();
            this.grpboxCreated.SuspendLayout();
            this.grpboxProjectFiles.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(672, 520);
            this.panel1.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.grpboxDescription);
            this.groupBox1.Controls.Add(this.grpboxCompany);
            this.groupBox1.Controls.Add(this.grpboxProjectId);
            this.groupBox1.Controls.Add(this.grpboxCreated);
            this.groupBox1.Controls.Add(this.grpboxProjectFiles);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.btnEdit);
            this.groupBox1.Controls.Add(this.btnSave);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(672, 520);
            this.groupBox1.TabIndex = 29;
            this.groupBox1.TabStop = false;
            // 
            // grpboxDescription
            // 
            this.grpboxDescription.Controls.Add(this.txtDescription);
            this.grpboxDescription.Enabled = false;
            this.grpboxDescription.Location = new System.Drawing.Point(8, 151);
            this.grpboxDescription.Name = "grpboxDescription";
            this.grpboxDescription.Size = new System.Drawing.Size(640, 121);
            this.grpboxDescription.TabIndex = 28;
            this.grpboxDescription.TabStop = false;
            this.grpboxDescription.Text = "Project Description";
            // 
            // txtDescription
            // 
            this.txtDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDescription.Location = new System.Drawing.Point(16, 16);
            this.txtDescription.Multiline = true;
            this.txtDescription.Name = "txtDescription";
            this.txtDescription.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtDescription.Size = new System.Drawing.Size(624, 96);
            this.txtDescription.TabIndex = 0;
            // 
            // grpboxCompany
            // 
            this.grpboxCompany.Controls.Add(this.txtCompany);
            this.grpboxCompany.Enabled = false;
            this.grpboxCompany.Location = new System.Drawing.Point(8, 97);
            this.grpboxCompany.Name = "grpboxCompany";
            this.grpboxCompany.Size = new System.Drawing.Size(640, 48);
            this.grpboxCompany.TabIndex = 27;
            this.grpboxCompany.TabStop = false;
            this.grpboxCompany.Text = "Company";
            // 
            // txtCompany
            // 
            this.txtCompany.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCompany.Location = new System.Drawing.Point(8, 14);
            this.txtCompany.MaxLength = 100;
            this.txtCompany.Name = "txtCompany";
            this.txtCompany.Size = new System.Drawing.Size(624, 23);
            this.txtCompany.TabIndex = 0;
            // 
            // grpboxProjectId
            // 
            this.grpboxProjectId.Controls.Add(this.txtProjectId);
            this.grpboxProjectId.Enabled = false;
            this.grpboxProjectId.Location = new System.Drawing.Point(8, 42);
            this.grpboxProjectId.Name = "grpboxProjectId";
            this.grpboxProjectId.Size = new System.Drawing.Size(184, 48);
            this.grpboxProjectId.TabIndex = 0;
            this.grpboxProjectId.TabStop = false;
            this.grpboxProjectId.Text = "Project Id";
            // 
            // txtProjectId
            // 
            this.txtProjectId.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtProjectId.Location = new System.Drawing.Point(8, 16);
            this.txtProjectId.MaxLength = 20;
            this.txtProjectId.Name = "txtProjectId";
            this.txtProjectId.Size = new System.Drawing.Size(166, 23);
            this.txtProjectId.TabIndex = 0;
            this.txtProjectId.Leave += new System.EventHandler(this.txtProjectId_Leave);
            // 
            // grpboxCreated
            // 
            this.grpboxCreated.Controls.Add(this.txtDate);
            this.grpboxCreated.Controls.Add(this.lblDate);
            this.grpboxCreated.Controls.Add(this.txtName);
            this.grpboxCreated.Controls.Add(this.lblName);
            this.grpboxCreated.Enabled = false;
            this.grpboxCreated.Location = new System.Drawing.Point(200, 42);
            this.grpboxCreated.Name = "grpboxCreated";
            this.grpboxCreated.Size = new System.Drawing.Size(450, 48);
            this.grpboxCreated.TabIndex = 1;
            this.grpboxCreated.TabStop = false;
            this.grpboxCreated.Text = "Created";
            // 
            // txtDate
            // 
            this.txtDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDate.Location = new System.Drawing.Point(296, 16);
            this.txtDate.MaxLength = 8;
            this.txtDate.Name = "txtDate";
            this.txtDate.Size = new System.Drawing.Size(144, 23);
            this.txtDate.TabIndex = 3;
            // 
            // lblDate
            // 
            this.lblDate.Location = new System.Drawing.Point(264, 21);
            this.lblDate.Name = "lblDate";
            this.lblDate.Size = new System.Drawing.Size(32, 16);
            this.lblDate.TabIndex = 2;
            this.lblDate.Text = "Date";
            // 
            // txtName
            // 
            this.txtName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtName.Location = new System.Drawing.Point(42, 15);
            this.txtName.MaxLength = 30;
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(222, 23);
            this.txtName.TabIndex = 1;
            // 
            // lblName
            // 
            this.lblName.Location = new System.Drawing.Point(7, 21);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(40, 16);
            this.lblName.TabIndex = 0;
            this.lblName.Text = "Name";
            // 
            // grpboxProjectFiles
            // 
            this.grpboxProjectFiles.Controls.Add(this.btnPersonalHelp);
            this.grpboxProjectFiles.Controls.Add(this.btnPersonalDirectory);
            this.grpboxProjectFiles.Controls.Add(this.lblLocal);
            this.grpboxProjectFiles.Controls.Add(this.txtPersonal);
            this.grpboxProjectFiles.Controls.Add(this.btnSharedHelp);
            this.grpboxProjectFiles.Controls.Add(this.btnSharedDirectory);
            this.grpboxProjectFiles.Controls.Add(this.lblShared);
            this.grpboxProjectFiles.Controls.Add(this.btnRootDirectoryHelp);
            this.grpboxProjectFiles.Controls.Add(this.btnRootDirectory);
            this.grpboxProjectFiles.Controls.Add(this.lblRootDirectory);
            this.grpboxProjectFiles.Controls.Add(this.txtRootDirectory);
            this.grpboxProjectFiles.Controls.Add(this.txtShared);
            this.grpboxProjectFiles.Location = new System.Drawing.Point(8, 280);
            this.grpboxProjectFiles.Name = "grpboxProjectFiles";
            this.grpboxProjectFiles.Size = new System.Drawing.Size(640, 152);
            this.grpboxProjectFiles.TabIndex = 29;
            this.grpboxProjectFiles.TabStop = false;
            this.grpboxProjectFiles.Text = "Project Files";
            // 
            // btnPersonalHelp
            // 
            this.btnPersonalHelp.Image = ((System.Drawing.Image)(resources.GetObject("btnPersonalHelp.Image")));
            this.btnPersonalHelp.Location = new System.Drawing.Point(600, 112);
            this.btnPersonalHelp.Name = "btnPersonalHelp";
            this.btnPersonalHelp.Size = new System.Drawing.Size(32, 32);
            this.btnPersonalHelp.TabIndex = 11;
            this.btnPersonalHelp.Visible = false;
            // 
            // btnPersonalDirectory
            // 
            this.btnPersonalDirectory.Enabled = false;
            this.btnPersonalDirectory.Image = ((System.Drawing.Image)(resources.GetObject("btnPersonalDirectory.Image")));
            this.btnPersonalDirectory.Location = new System.Drawing.Point(561, 112);
            this.btnPersonalDirectory.Name = "btnPersonalDirectory";
            this.btnPersonalDirectory.Size = new System.Drawing.Size(32, 32);
            this.btnPersonalDirectory.TabIndex = 10;
            this.btnPersonalDirectory.Click += new System.EventHandler(this.btnPersonalDirectory_Click);
            // 
            // lblLocal
            // 
            this.lblLocal.Location = new System.Drawing.Point(16, 90);
            this.lblLocal.Name = "lblLocal";
            this.lblLocal.Size = new System.Drawing.Size(80, 53);
            this.lblLocal.TabIndex = 9;
            this.lblLocal.Text = "Personal Notes And Document Links Directory";
            // 
            // txtPersonal
            // 
            this.txtPersonal.Enabled = false;
            this.txtPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPersonal.ForeColor = System.Drawing.SystemColors.WindowFrame;
            this.txtPersonal.Location = new System.Drawing.Point(112, 116);
            this.txtPersonal.Name = "txtPersonal";
            this.txtPersonal.Size = new System.Drawing.Size(416, 23);
            this.txtPersonal.TabIndex = 8;
            // 
            // btnSharedHelp
            // 
            this.btnSharedHelp.Image = ((System.Drawing.Image)(resources.GetObject("btnSharedHelp.Image")));
            this.btnSharedHelp.Location = new System.Drawing.Point(601, 44);
            this.btnSharedHelp.Name = "btnSharedHelp";
            this.btnSharedHelp.Size = new System.Drawing.Size(32, 32);
            this.btnSharedHelp.TabIndex = 7;
            this.btnSharedHelp.Visible = false;
            // 
            // btnSharedDirectory
            // 
            this.btnSharedDirectory.Enabled = false;
            this.btnSharedDirectory.Image = ((System.Drawing.Image)(resources.GetObject("btnSharedDirectory.Image")));
            this.btnSharedDirectory.Location = new System.Drawing.Point(561, 44);
            this.btnSharedDirectory.Name = "btnSharedDirectory";
            this.btnSharedDirectory.Size = new System.Drawing.Size(32, 32);
            this.btnSharedDirectory.TabIndex = 6;
            this.btnSharedDirectory.Click += new System.EventHandler(this.btnSharedDirectory_Click);
            // 
            // lblShared
            // 
            this.lblShared.Location = new System.Drawing.Point(16, 44);
            this.lblShared.Name = "lblShared";
            this.lblShared.Size = new System.Drawing.Size(80, 40);
            this.lblShared.TabIndex = 4;
            this.lblShared.Text = "Shared Notes And Document Links Directory";
            // 
            // btnRootDirectoryHelp
            // 
            this.btnRootDirectoryHelp.Image = ((System.Drawing.Image)(resources.GetObject("btnRootDirectoryHelp.Image")));
            this.btnRootDirectoryHelp.Location = new System.Drawing.Point(601, 10);
            this.btnRootDirectoryHelp.Name = "btnRootDirectoryHelp";
            this.btnRootDirectoryHelp.Size = new System.Drawing.Size(32, 32);
            this.btnRootDirectoryHelp.TabIndex = 3;
            this.btnRootDirectoryHelp.Visible = false;
            // 
            // btnRootDirectory
            // 
            this.btnRootDirectory.Enabled = false;
            this.btnRootDirectory.Image = ((System.Drawing.Image)(resources.GetObject("btnRootDirectory.Image")));
            this.btnRootDirectory.Location = new System.Drawing.Point(561, 10);
            this.btnRootDirectory.Name = "btnRootDirectory";
            this.btnRootDirectory.Size = new System.Drawing.Size(32, 32);
            this.btnRootDirectory.TabIndex = 2;
            this.btnRootDirectory.Click += new System.EventHandler(this.btnRootDirectory_Click);
            // 
            // lblRootDirectory
            // 
            this.lblRootDirectory.Location = new System.Drawing.Point(16, 19);
            this.lblRootDirectory.Name = "lblRootDirectory";
            this.lblRootDirectory.Size = new System.Drawing.Size(88, 16);
            this.lblRootDirectory.TabIndex = 0;
            this.lblRootDirectory.Text = "Project Directory";
            // 
            // txtRootDirectory
            // 
            this.txtRootDirectory.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtRootDirectory.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRootDirectory.Location = new System.Drawing.Point(113, 16);
            this.txtRootDirectory.Name = "txtRootDirectory";
            this.txtRootDirectory.Size = new System.Drawing.Size(416, 23);
            this.txtRootDirectory.TabIndex = 0;
            // 
            // txtShared
            // 
            this.txtShared.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtShared.ForeColor = System.Drawing.SystemColors.WindowFrame;
            this.txtShared.Location = new System.Drawing.Point(112, 48);
            this.txtShared.Name = "txtShared";
            this.txtShared.Size = new System.Drawing.Size(416, 23);
            this.txtShared.TabIndex = 0;
            // 
            // lblTitle
            // 
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(8, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(500, 24);
            this.lblTitle.TabIndex = 26;
            this.lblTitle.Text = "Open Or Create A Project";
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(208, 440);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(96, 32);
            this.btnEdit.TabIndex = 7;
            this.btnEdit.Text = "Edit";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnSave
            // 
            this.btnSave.Enabled = false;
            this.btnSave.Location = new System.Drawing.Point(304, 440);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(96, 32);
            this.btnSave.TabIndex = 5;
            this.btnSave.Text = "Save";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(400, 440);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(96, 32);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Location = new System.Drawing.Point(560, 480);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 8;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(112, 440);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(96, 32);
            this.btnHelp.TabIndex = 31;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // uc_project
            // 
            this.Controls.Add(this.panel1);
            this.Name = "uc_project";
            this.Size = new System.Drawing.Size(672, 520);
            this.Resize += new System.EventHandler(this.uc_project_Resize);
            this.panel1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.grpboxDescription.ResumeLayout(false);
            this.grpboxDescription.PerformLayout();
            this.grpboxCompany.ResumeLayout(false);
            this.grpboxCompany.PerformLayout();
            this.grpboxProjectId.ResumeLayout(false);
            this.grpboxProjectId.PerformLayout();
            this.grpboxCreated.ResumeLayout(false);
            this.grpboxCreated.PerformLayout();
            this.grpboxProjectFiles.ResumeLayout(false);
            this.grpboxProjectFiles.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			this.grpboxProjectFiles.Enabled=true;
			this.txtDate.Enabled=false;

			this.grpboxCreated.Enabled=true;
			this.grpboxCompany.Enabled=true;
			this.grpboxDescription.Enabled=true;
			this.btnSave.Enabled=true;
			this.btnCancel.Enabled=true;

			this.btnPersonalDirectory.Enabled=true;
			this.btnSharedDirectory.Enabled=true;
			this.btnEdit.Enabled=false;
            this.grpboxProjectId.Enabled = true;
			

		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.grpboxProjectId.Enabled=false;
			this.grpboxProjectFiles.Enabled=false;
			this.grpboxCreated.Enabled=false;
			this.grpboxCompany.Enabled=false;
			this.grpboxDescription.Enabled=false;
			this.btnSave.Enabled=false;
			this.btnCancel.Enabled=false;
			this.btnEdit.Enabled=true;
			if (this.m_strAction == "NEW") 
			{
				this.Parent.Visible=false;
				
			}
			else 
			{
				this.OpenProjectTable(this.m_strProjectDirectory,this.m_strProjectFile);
			}
		    this.m_strAction="";
		    
		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{

			this.SaveProjectProperties();
				

			
		}
		public void SaveProjectProperties()
		{
			string strDestFile;
			string strSourceFile;
			string strConn;
			string strSQL;
			string strFullPath;
			DialogResult result = new DialogResult();
			int x;
			int intAt = 0;
			string strDesc="";

			//validate the input
			//project id
			if (this.txtProjectId.Text.Length == 0) 
			{
				MessageBox.Show("Enter A Project Id ");
				this.txtProjectId.Focus();
				return;
			}


			//project root directory
			if (this.txtRootDirectory.Text.Length == 0 ) 
			{
				MessageBox.Show("Enter A Project Root Directory");
				this.txtRootDirectory.Focus();
				return;
			}

			try
			{
				if (this.m_strAction == "NEW") 
				{
				
					if (System.IO.Directory.Exists(this.txtRootDirectory.Text))
					{
						for (x=1;x<=1000;x++)
						{
							strFullPath = this.txtRootDirectory.Text.Trim() + x.ToString().Trim();
							if (!System.IO.Directory.Exists(strFullPath))
							{
								this.txtRootDirectory.Text = strFullPath.Trim();
								break;
							}
						}
					}
				}
				strFullPath = this.txtRootDirectory.Text.Trim() + "\\db";
				if (!System.IO.Directory.Exists(strFullPath))
					System.IO.Directory.CreateDirectory(strFullPath);

				    
				strFullPath = this.txtRootDirectory.Text.Trim() + "\\optimizer\\db";
				if (!System.IO.Directory.Exists(strFullPath))
					System.IO.Directory.CreateDirectory(strFullPath);

				strFullPath = this.txtRootDirectory.Text.Trim() + "\\gis\\db";
				if (!System.IO.Directory.Exists(strFullPath))
					System.IO.Directory.CreateDirectory(strFullPath);

				strFullPath = this.txtRootDirectory.Text.Trim() + "\\fvs\\db";
				if (!System.IO.Directory.Exists(strFullPath))
					System.IO.Directory.CreateDirectory(strFullPath);

				strFullPath = this.txtRootDirectory.Text.Trim() + "\\fvs\\data";
				if (!System.IO.Directory.Exists(strFullPath))
					System.IO.Directory.CreateDirectory(strFullPath);

				strFullPath = this.txtRootDirectory.Text.Trim() + "\\processor\\db";
				if (!System.IO.Directory.Exists(strFullPath))
					System.IO.Directory.CreateDirectory(strFullPath);
		
				
			}
			catch 
			{
				MessageBox.Show("Error Creating Project Folder");
				
				return;
			}

		
			//check if new project
			ado_data_access p_ado = new ado_data_access();
			dao_data_access p_dao = new dao_data_access();
            DataMgr p_dataMgr = new DataMgr();
			if (this.m_strAction == "NEW") 
			{
				//new project
				//copy default project file to new project directory
				

				strSourceFile = this.m_oEnv.strAppDir + "\\db\\project.mdb";
				strDestFile = this.txtRootDirectory.Text.Trim() + "\\db\\project.mdb";
				
				FIA_Biosum_Manager.frmTherm p_frmTherm;
				p_frmTherm = new FIA_Biosum_Manager.frmTherm();
				p_frmTherm.btnCancel.Visible=false;
				p_frmTherm.Show();
				p_frmTherm.Focus();
			
				
				p_frmTherm.Text = "Creating Project Files";
				p_frmTherm.Refresh();
				p_frmTherm.progressBar1.Minimum = 1;
				p_frmTherm.AbortProcess = false;
				p_frmTherm.progressBar1.Maximum = 13;
				p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Visible=true;
				p_frmTherm.lblMsg.Refresh();


				
				//
				//project file and tables
				//
				p_dao.CreateMDB(strDestFile);
				strConn = p_ado.getMDBConnString(strDestFile,"admin","");
				p_ado.OpenConnection(strConn);
				//contacts table
				frmMain.g_oTables.m_oProject.CreateContactsTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectContactsTableName);
				//datasource table
				frmMain.g_oTables.m_oProject.CreateDatasourceTable(p_ado,p_ado.m_OleDbConnection,Tables.Project.DefaultProjectDatasourceTableName);
				//project table
				frmMain.g_oTables.m_oProject.CreateProjectTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectTableName);
				//user config table
				frmMain.g_oTables.m_oProject.CreateUserConfigTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectUserConfigTableName);
				p_ado.CloseConnection(p_ado.m_OleDbConnection);
				//
				//travel times file and tables
				//
				p_frmTherm.lblMsg.Text = strDestFile;
                strDestFile = this.txtRootDirectory.Text.Trim() + "\\" + Tables.TravelTime.DefaultTravelTimePathAndDbFile;
                p_frmTherm.Increment(2);
                p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
                p_dataMgr.CreateDbFile(strDestFile);
				strConn = p_dataMgr.GetConnectionString(strDestFile);
                using (var oConn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    oConn.Open();
                    //processing site table
                    frmMain.g_oTables.m_oTravelTime.CreateSqliteProcessingSiteTable(p_dataMgr, oConn, Tables.TravelTime.DefaultProcessingSiteTableName);
                    //travel time table
                    frmMain.g_oTables.m_oTravelTime.CreateSqliteTravelTimeTable(p_dataMgr, oConn, Tables.TravelTime.DefaultTravelTimeTableName);
                }
                strDestFile = this.txtRootDirectory.Text.Trim() + "\\" + Tables.TravelTime.DefaultGisAuditPathAndDbFile;
                p_dataMgr.CreateDbFile(strDestFile);

                //
                //master file
                //
                //copy default master database to the new project directory (not currently used)
                strSourceFile = this.m_oEnv.strAppDir + "\\db\\master.mdb";
				strDestFile = this.txtRootDirectory.Text.Trim() + "\\db\\master.mdb";
				p_frmTherm.Increment(3);
				p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
				//System.IO.File.Copy(strSourceFile, strDestFile,true);	
				p_dao.CreateMDB(strDestFile);
				strConn = p_ado.getMDBConnString(strDestFile,"admin","");
				p_ado.OpenConnection(strConn);
				//plot table
				frmMain.g_oTables.m_oFIAPlot.CreatePlotTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName);
				//cond table
				frmMain.g_oTables.m_oFIAPlot.CreateConditionTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName);
				//pop estimation unit table
				frmMain.g_oTables.m_oFIAPlot.CreatePopEstnUnitTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName);
				//pop eval table
				frmMain.g_oTables.m_oFIAPlot.CreatePopEvalTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName);
				//pop plot stratum assignment table
				frmMain.g_oTables.m_oFIAPlot.CreatePopPlotStratumAssgnTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName);
				//pop stratum table
				frmMain.g_oTables.m_oFIAPlot.CreatePopStratumTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName);
				//site tree table
				frmMain.g_oTables.m_oFIAPlot.CreateSiteTreeTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName);
				//tree table
				frmMain.g_oTables.m_oFIAPlot.CreateTreeTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName);
                //biosum pop stratum adjustment factors table
                frmMain.g_oTables.m_oFIAPlot.CreateBiosumPopStratumAdjustmentFactorsTable(p_ado, p_ado.m_OleDbConnection, frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName);

				p_ado.CloseConnection(p_ado.m_OleDbConnection);

                //master.db file: Migrating POP tables to SQLite
                strDestFile = this.txtRootDirectory.Text.Trim() + "\\" + frmMain.g_oTables.m_oFIAPlot.DefaultPopTableDbFile;
                p_dataMgr.CreateDbFile(strDestFile);
                strConn = p_dataMgr.GetConnectionString(strDestFile);
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    con.Open();
                    //pop estimation unit table
                    frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopEstnUnitTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName);
                    //pop eval table
                    frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopEvalTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName);
                    //pop plot stratum assignment table
                    frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopPlotStratumAssgnTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName);
                    //pop stratum table
                    frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopStratumTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName);
                }

                //
                //master_aux file: Where additional input tables are stored
                //
                strDestFile = this.txtRootDirectory.Text.Trim() + "\\db\\master_aux.accdb";
				p_frmTherm.Increment(4);
				p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
				//System.IO.File.Copy(strSourceFile, strDestFile,true);	
				p_dao.CreateMDB(strDestFile);
				strConn = p_ado.getMDBConnString(strDestFile,"admin","");
				p_ado.OpenConnection(strConn);
                //DWM Section
                frmMain.g_oTables.m_oFIAPlot.CreateDWMCoarseWoodyDebrisTable(p_ado, p_ado.m_OleDbConnection, frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName);
                frmMain.g_oTables.m_oFIAPlot.CreateDWMFineWoodyDebrisTable(p_ado, p_ado.m_OleDbConnection, frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName);
                frmMain.g_oTables.m_oFIAPlot.CreateDWMDuffLitterFuelTable(p_ado, p_ado.m_OleDbConnection, frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName);
                frmMain.g_oTables.m_oFIAPlot.CreateDWMTransectSegmentTable(p_ado, p_ado.m_OleDbConnection, frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName);

				p_ado.CloseConnection(p_ado.m_OleDbConnection);

                //
                //master_aux.db
                //
                //strDestFile = this.txtRootDirectory.Text.Trim() + "\\db\\master_aux.db";
                //p_frmTherm.Increment(4);
                //p_frmTherm.lblMsg.Text = strDestFile;
                //p_frmTherm.lblMsg.Refresh();
                //p_dataMgr.CreateDbFile(strDestFile);
                //strConn = p_dataMgr.GetConnectionString(strDestFile);
                //using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                //{
                //    conn.Open();
                //    frmMain.g_oTables.m_oFIAPlot.CreateSqliteDWMCoarseWoodyDebrisTable(p_dataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName);
                //    frmMain.g_oTables.m_oFIAPlot.CreateSqliteDWMFineWoodyDebrisTable(p_dataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName);
                //    frmMain.g_oTables.m_oFIAPlot.CreateSqliteDWMDuffLitterFuelTable(p_dataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName);
                //    frmMain.g_oTables.m_oFIAPlot.CreateSqliteDWMTransectSegmentTable(p_dataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName);
                //}

                //
                //fvsmaster file
                //
                //copy default fvsmaster database to the new project directory
                //strSourceFile = this.m_oEnv.strAppDir + "\\db\\fvsmaster.mdb";
				strDestFile = this.txtRootDirectory.Text.Trim() + "\\db\\fvsmaster.mdb";
				p_frmTherm.Increment(5);
				p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
				//System.IO.File.Copy(strSourceFile, strDestFile,true);	
				p_dao.CreateMDB(strDestFile);
				strConn = p_ado.getMDBConnString(strDestFile,"admin","");
				p_ado.OpenConnection(strConn);
				//rx table
				frmMain.g_oTables.m_oFvs.CreateRxTable(p_ado,p_ado.m_OleDbConnection,Tables.FVS.DefaultRxTableName);
				//rx harvest cost column table
				frmMain.g_oTables.m_oFvs.CreateRxHarvestCostColumnTable(p_ado,p_ado.m_OleDbConnection, Tables.FVS.DefaultRxHarvestCostColumnsTableName);
				//rx packages table
				frmMain.g_oTables.m_oFvs.CreateRxPackageTable(p_ado,p_ado.m_OleDbConnection,Tables.FVS.DefaultRxPackageTableName);
				//rx package members table
				//frmMain.g_oTables.m_oFvs.CreateRxPackageMembersTable(p_ado,p_ado.m_OleDbConnection,Tables.FVS.DefaultRxPackageMembersTableName);
				//rx package fvs commands table
				//frmMain.g_oTables.m_oFvs.CreateRxPackageFvsCommandsTable(p_ado,p_ado.m_OleDbConnection,Tables.FVS.DefaultRxPackageFvsCommandTableName);
				//rx package fvs commands order
				//frmMain.g_oTables.m_oFvs.CreateRxPackageFvsCommandsOrderTable(p_ado,p_ado.m_OleDbConnection,Tables.FVS.DefaultRxPackageFvsCommandsOrderTableName);
				p_ado.CloseConnection(p_ado.m_OleDbConnection);
                //fvs output pre-post seqnum processing
                uc_fvs_output_prepost_seqnum.InitializePrePostSeqNumTables();   // This is the SQLite version

                // Moving these tables to master.db during SQLite conversion
                strDestFile = this.txtRootDirectory.Text.Trim() + "\\" + Tables.FVS.DefaultRxPackageDbFile;
                strConn = p_dataMgr.GetConnectionString(strDestFile);
                //using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                //{
                //    conn.Open();
                //    //rx table
                //    frmMain.g_oTables.m_oFvs.CreateSQLiteRxTable(p_dataMgr, conn, Tables.FVS.DefaultRxTableName);
                //    //rx packages table
                //    frmMain.g_oTables.m_oFvs.CreateSQLiteRxPackageTable(p_dataMgr, conn, Tables.FVS.DefaultRxPackageTableName);
                //    //rx harvest cost column table
                //    frmMain.g_oTables.m_oFvs.CreateSqliteRxHarvestCostColumnTable(p_dataMgr, conn, Tables.FVS.DefaultRxHarvestCostColumnsTableName);
                //}
                //@ToDo: Change destination for uc_fvs_output_prepost_seqnum.InitializePrePostSeqNumTables()


                //
                //prepopulated ref master file
                //
                //copy default master database to the new project directory
                strSourceFile = this.m_oEnv.strAppDir + "\\db\\ref_master.mdb";
				strDestFile = this.txtRootDirectory.Text.Trim() + "\\db\\ref_master.mdb";
				p_frmTherm.Increment(6);
				p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
				System.IO.File.Copy(strSourceFile, strDestFile,true);
                //
                //prepopulated OPCOST ref file file
                //
                //copy default master database to the new project directory
                strSourceFile = this.m_oEnv.strAppDir + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile;
                strDestFile = this.txtRootDirectory.Text.Trim() + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile;
                System.IO.File.Copy(strSourceFile, strDestFile, true);
                //
                //prepopulated ref fvs commands file
                //
                //copy default master database to the new project directory
                strSourceFile = this.m_oEnv.strAppDir + "\\db\\ref_fvscommands.mdb";
                strDestFile = this.txtRootDirectory.Text.Trim() + "\\db\\ref_fvscommands.mdb";
                p_frmTherm.Increment(7);
                p_frmTherm.lblMsg.Text = strDestFile;
                p_frmTherm.lblMsg.Refresh();
                System.IO.File.Copy(strSourceFile, strDestFile, true);
                //
                //prepopulated weighted variable optimizer_definitions.db file
                //
				//copy default optimizer_definitions.db to new project directory
				strSourceFile = this.m_oEnv.strAppDir + "\\db\\optimizer_definitions.db";
				strDestFile = this.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerDefinitions.DefaultSqliteDbFile;
				p_frmTherm.Increment(8);
				p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
				System.IO.File.Copy(strSourceFile, strDestFile, true);

                //
                //optimizer scenario rule definitions
                //
				p_frmTherm.Increment(9);
                p_frmTherm.lblMsg.Text = this.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;
				p_frmTherm.lblMsg.Refresh();
                //CreateOptimizerScenarioRuleDefinitionDbAndTables(this.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile);
				CreateOptimizerScenarioRuleDefinitionSqliteDbAndTables(this.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile);
				//
				//processor scenario rule definitions
				//
				p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
                CreateProcessorScenarioRuleDefinitionDbAndTables($@"{this.txtRootDirectory.Text.Trim()}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile}");

                p_frmTherm.Increment(10);
				//p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
				//System.IO.File.Copy(strSourceFile, strDestFile,true);		
				
				if (this.txtShared.Text.Trim().Length > 0)
				{
					strDestFile = this.txtRootDirectory.Text.Trim() + "\\db\\shared_project_links_and_notes.mdb";
					p_dao.CreateMDB(strDestFile);
					p_frmTherm.Increment(11);
					p_frmTherm.lblMsg.Text = strDestFile;
					p_frmTherm.lblMsg.Refresh();
					strConn = p_ado.getMDBConnString(strDestFile,"admin","");
					p_ado.OpenConnection(strConn);
					frmMain.g_oTables.m_oProject.CreateProjectLinksCategoryTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksCategoryTableName);
					frmMain.g_oTables.m_oProject.CreateProjectLinksDepositoryTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksDepositoryTableName);
					frmMain.g_oTables.m_oProject.CreateProjectNotesTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectNotesTableName);
					p_ado.CloseConnection(p_ado.m_OleDbConnection);

				}

				
				if (this.txtPersonal.Text.Length == 0)
				{
					result = MessageBox.Show("Do you want to keep personal project related notes and document links (Y/N)","test", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
					switch (result) 
					{
						case DialogResult.Yes:
							//check to see if the project drive is on the users computer
							for (x=0; x<= ((frmMain)this.ParentForm.ParentForm).m_LocalHardDrive.Length-1; x++)
							{
								intAt = this.txtRootDirectory.Text.IndexOf(((frmMain)this.ParentForm.ParentForm).m_LocalHardDrive[x].Trim().ToUpper(),0, this.txtRootDirectory.Text.Length);
								if (intAt >=0) 
								{
									break;
								}
							}
							if (x <= ((frmMain)this.ParentForm.ParentForm).m_LocalHardDrive.Length-1) 
							{
								//create a personal copy of the project file
								if (!System.IO.Directory.Exists(this.txtRootDirectory.Text.Trim() + "\\db\\" + System.Environment.UserName.ToString().Trim())) 
								{
									strFullPath = this.txtRootDirectory.Text.Trim() + "\\db\\" + System.Environment.UserName.ToString().Trim();
									System.IO.Directory.CreateDirectory(strFullPath);
								}
								strDestFile = this.txtRootDirectory.Text.Trim() + "\\db\\" +  System.Environment.UserName.ToString().Trim() + "\\personal_project_links_and_notes.mdb";
								p_dao.CreateMDB(strDestFile);
								p_frmTherm.Increment(12);
								p_frmTherm.lblMsg.Text = strDestFile;
								p_frmTherm.lblMsg.Refresh();
								this.txtPersonal.Text = this.txtRootDirectory.Text.Trim() + "\\db\\" +  System.Environment.UserName.ToString().Trim();
								strConn = p_ado.getMDBConnString(strDestFile,"admin","");
								p_ado.OpenConnection(strConn);
								frmMain.g_oTables.m_oProject.CreateProjectLinksCategoryTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksCategoryTableName);
								frmMain.g_oTables.m_oProject.CreateProjectLinksDepositoryTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksDepositoryTableName);
								frmMain.g_oTables.m_oProject.CreateProjectNotesTable(p_ado,p_ado.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectNotesTableName);
								p_ado.CloseConnection(p_ado.m_OleDbConnection);
							}
							else 
							{


							}
							break;
						case DialogResult.No:
							break;
					}                
				}
				
				p_frmTherm.Increment(13);
				strSourceFile = this.txtRootDirectory.Text.Trim() + "\\db\\project.mdb";
				strConn = p_ado.getMDBConnString(strSourceFile,"admin","");
				p_frmTherm.Close();
				p_frmTherm.Dispose();
				p_frmTherm = null;

				p_ado.OpenConnection(strConn);
				if (p_ado.m_intError == 0)
				{
					if (this.txtDescription.Text.Trim().Length > 0)
						strDesc = p_ado.FixString(this.txtDescription.Text.Trim(),"'","''");
					strSQL = "INSERT INTO project (proj_id,created_by,created_date,company,description,shared_file,project_root_directory,application_version) VALUES " + "(" +  
						"'" + this.txtProjectId.Text.Trim() + "'," + 
						"'" + this.txtName.Text.Trim() + "'," + 
						"'" + this.txtDate.Text +  "'," + 
						"'" + this.txtCompany.Text.Trim() + "'," + 
						"'" + strDesc + "'," + 
						"'" + this.txtShared.Text.Trim() + "'," + 
						"'" + this.txtRootDirectory.Text.Trim() + "'," + 
						"'" + frmMain.g_strAppVer + "');";

					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

					if (this.txtPersonal.Text.Trim().Length > 0)
					{
						strSQL = "DELETE * FROM user_config WHERE user_name =  " + "'" + System.Environment.UserName.ToString().Trim() + "'";
						p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);
						strSQL = "INSERT INTO user_config (user_name,personal_directory) VALUES " + "(" +  
							"'" + System.Environment.UserName.ToString().Trim() + "'," + 
							"'" + this.txtPersonal.Text.Trim() + "');";

						p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

					}

					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " + 
						"('Plot'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim() + "\\db'," + 
						"'master.mdb'," + 
						"'plot');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " + 
						"('Condition'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim() + "\\db'," + 
						"'master.mdb'," + 
						"'cond');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " + 
						"('Tree'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim() + "\\db'," + 
						"'master.mdb'," + 
						"'tree');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
						"('Owner Groups'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim() + "\\db'," + 
						"'ref_master.mdb'," + 
						"'owner_groups');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
						"('Treatment Prescriptions'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim()  + "\\db'," + 
						"'fvsmaster.mdb'," + 
						"'rx');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

                    strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
						"('Treatment Prescriptions Harvest Cost Columns'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim()  + "\\db'," + 
						"'fvsmaster.mdb'," + 
						"'" + Tables.FVS.DefaultRxHarvestCostColumnsTableName + "');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);
					
					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
						"('Treatment Prescription Categories'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim()  + "\\db'," + 
						"'ref_master.mdb'," + 
						"'fvs_rx_category');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
						"('Treatment Prescription Subcategories'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim()  + "\\db'," + 
						"'ref_master.mdb'," + 
						"'fvs_rx_subcategory');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
						"('Treatment Packages'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim()  + "\\db'," + 
						"'fvsmaster.mdb'," + 
						"'rxpackage');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

                    strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                        "('FVS PRE-POST SeqNum Definitions'," +
                        "'" + this.txtRootDirectory.Text.ToString().Trim() + "\\db'," +
                        "'fvsmaster.db'," +
                        "'" + Tables.FVS.DefaultFVSPrePostSeqNumTable + "');";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);

                    strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                        "('FVS PRE-POST SeqNum Treatment Package Assign'," +
                        "'" + this.txtRootDirectory.Text.ToString().Trim() + "\\db'," +
                        "'fvsmaster.db'," +
                        "'" + Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable + "');";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);

					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
						"('Tree Species'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim() + "\\db'," + 
						"'ref_master.mdb'," + 
						"'tree_species');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
						"('" + Datasource.TableTypes.FvsTreeSpecies + "'," + 
						"'@@AppData@@" + frmMain.g_strBiosumDataDir + "'," + 
						"'" + Tables.Reference.DefaultBiosumReferenceDbFile + "'," + 
						"'fvs_tree_species');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

                    strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                        "('FVS Western Tree Species Translator'," +
                        "'" + this.txtRootDirectory.Text.ToString().Trim() + "\\db'," +
                        "'ref_master.mdb'," +
                        "'FVS_WesternTreeSpeciesTranslator');";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);

                    strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                        "('FVS Eastern Tree Species Translator'," +
                        "'" + this.txtRootDirectory.Text.ToString().Trim() + "\\db'," +
                        "'ref_master.mdb'," +
                        "'FVS_EasternTreeSpeciesTranslator');";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);

					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
						"('Travel Times'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim() + "\\gis\\db'," + 
						"'" + Tables.TravelTime.DefaultTravelTimeDbFile + "'," + 
						"'travel_time');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " + 
						"('Processing Sites'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim() + "\\gis\\db'," +
                        "'" + Tables.TravelTime.DefaultTravelTimeDbFile + "'," + 
						"'processing_site');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);
					
                    strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                        "('" + Datasource.TableTypes.FVSVariant + "'," +
                        "'@@AppData@@" + frmMain.g_strBiosumDataDir + "'," +
                        "'" + Tables.Reference.DefaultBiosumReferenceDbFile + "'," +
						"'fiadb_fvs_variant');";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);


                    strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                        "('FIA Tree Macro Plot Breakpoint Diameter'," +
                        "'" + this.txtRootDirectory.Text.ToString().Trim() + "\\db'," +
                        "'ref_master.mdb'," +
                        "'TreeMacroPlotBreakPointDia');";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);

					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " + 
						"('Harvest Methods'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim() + "\\db'," + 
						"'ref_master.mdb'," +
                        "'harvest_methods');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

                    strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                        "('BIOSUM Pop Stratum Adjustment Factors'," +
                        "'" + this.txtRootDirectory.Text.ToString().Trim() + "\\db'," +
                        "'master.mdb'," +
                        "'" + frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName + "');";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);

					strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
						"('Site Tree'," + 
						"'" + this.txtRootDirectory.Text.ToString().Trim() + "\\db'," + 
						"'master.mdb'," + 
						"'sitetree');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

                    strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                             "('" + Datasource.TableTypes.FiaTreeSpeciesReference + "'," +
                             "'@@AppData@@" + frmMain.g_strBiosumDataDir + "'," +
                             "'" + Tables.Reference.DefaultBiosumReferenceDbFile + "'," +
                             "'" + Tables.ProcessorScenarioRun.DefaultFiaTreeSpeciesRefTableName + "');";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);

                    frmMain.g_oGeneralMacroSubstitutionVariable_Collection.Item(frmMain.PROJDIR).VariableSubstitutionString = this.txtRootDirectory.Text.Trim();
					
				}
				p_ado.CloseConnection(p_ado.m_OleDbConnection);
				p_ado.m_OleDbConnection = null;

				frmMain.g_oUtils.WriteText(this.txtRootDirectory.Text.Trim() + "\\application.version",frmMain.g_strAppVer);
				



				//make the new project the current project
				this.OpenProjectTable(this.txtRootDirectory.Text,"project.mdb");

				if (this.m_intError == 0)
				{
					this.lblTitle.Text = "Project Properties";
				}
				((frmMain)this.ParentForm.ParentForm).OpenProject(this.txtRootDirectory.Text,"project.mdb");

			}
			else 
			{
				System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
				strFullPath = this.m_strProjectDirectory.Trim() + "\\db\\" + this.m_strProjectFile;
				strConn = p_ado.getMDBConnString(strFullPath,"admin","");
				if (this.txtDescription.Text.Trim().Length > 0)
					strDesc = p_ado.FixString(this.txtDescription.Text.Trim(),"'","''");
                strSQL = "UPDATE project SET created_by = '" + this.txtName.Text + "', " +
                    "proj_id = '" + this.txtProjectId.Text.Trim() + "', " +
                    "company = '" + this.txtCompany.Text + "', " +
                    "description = '" + strDesc + "', " +
                    "shared_file = '" + this.txtShared.Text + "', " +
                    "project_root_directory = '" + this.txtRootDirectory.Text + "' ";
				p_ado.SqlNonQuery(strConn,strSQL);
                this.m_strProjectId = this.txtProjectId.Text.Trim();
                ((frmMain)this.ParentForm.ParentForm).Text = "Fia Biosum Manager (" + this.m_strProjectId + ")";
				System.Data.OleDb.OleDbCommand oCommand = new System.Data.OleDb.OleDbCommand();
				if (this.txtPersonal.Text.Trim().Length > 0)
				{
					strSQL = "DELETE * FROM user_config WHERE user_name =  " + "'" + System.Environment.UserName.ToString().Trim() + "'";
					p_ado.SqlNonQuery(strConn,strSQL);
					strSQL = "INSERT INTO user_config (user_name,personal_directory) VALUES " + "(" +  
						"'" + System.Environment.UserName.ToString().Trim() + "'," + 
						"'" + this.txtPersonal.Text.Trim() + "');";

					p_ado.SqlNonQuery(strConn,strSQL);

				}			
				
			}
			p_ado=null;
            p_dataMgr = null;
			this.btnSave.Enabled=false;
			this.btnCancel.Enabled=false;
			this.btnEdit.Enabled=true;
			this.grpboxProjectFiles.Enabled=false;
			this.grpboxCreated.Enabled=false;
			this.grpboxCompany.Enabled=false;
			this.grpboxDescription.Enabled=false;
			this.grpboxProjectId.Enabled=false;
			string tempstr = this.txtRootDirectory.Text ;
			this.m_strProjectDirectory = tempstr; 
			this.m_strProjectFile = "project.mdb";
			this.m_strProjectId = this.txtProjectId.Text.Trim();
			this.m_strAction="";
			if (((frmMain)this.ParentForm.ParentForm).tlbMain.Buttons[4].Enabled==false)
			{
				if (this.txtPersonal.Text.Trim().Length > 0 || this.txtShared.Text.Trim().Length > 0)
				{
					((frmMain)this.ParentForm.ParentForm).tlbMain.Buttons[4].Enabled=true;
					((frmMain)this.ParentForm.ParentForm).tlbMain.Buttons[3].Enabled=true;
				}
				else
				{
					((frmMain)this.ParentForm.ParentForm).tlbMain.Buttons[4].Enabled=false;
					((frmMain)this.ParentForm.ParentForm).tlbMain.Buttons[3].Enabled=false;
				}
			}


		}
		public void New_Project()
		{
			this.m_strAction="NEW";
		    
            this.txtCompany.Text = "";
			this.txtDate.Text = System.DateTime.Now.ToString();
			this.txtDescription.Text = "";
			this.txtName.Text = "";
			this.txtProjectId.Text = "";
			
			this.txtRootDirectory.Text = this.m_oEnv.strAppDir.Substring(0,2) + "\\FIA_Biosum";
			this.txtShared.Text = "";
			this.txtPersonal.Text = "";
			this.txtProjectId.Enabled=true;
			this.txtName.Enabled=true;
			this.txtCompany.Enabled=true;
			this.txtDescription.Enabled=true;
			this.txtDate.Enabled=false;
			this.btnRootDirectory.Enabled=true;
			this.btnSharedDirectory.Enabled=true;
			this.btnPersonalDirectory.Enabled=true;
			this.btnEdit.Enabled=false;
			this.btnSave.Enabled=true;
			this.btnCancel.Enabled=true;
			this.grpboxProjectFiles.Enabled=true;
			this.grpboxCreated.Enabled=true;
			this.grpboxCompany.Enabled=true;
			this.grpboxDescription.Enabled=true;
			this.grpboxProjectId.Enabled=true;
			this.Parent.Visible = true;
		}
		private void grpboxProjectFiles_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnRootDirectory_Click(object sender, System.EventArgs e)
		{
			DialogResult result =  this.folderBrowserDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				string strTemp = this.folderBrowserDialog1.SelectedPath;
			
				if (strTemp.Length > 0) 
				{
    				this.txtRootDirectory.Text = strTemp + "\\" + this.txtProjectId.Text.Trim();
				}
			}
		}
		public void CreateOptimizerScenarioRuleDefinitionDbAndTables(string p_strPathAndFile)
		{
			dao_data_access oDao = new dao_data_access();
			ado_data_access oAdo = new ado_data_access();

			string strDestFile = p_strPathAndFile;
			oDao.CreateMDB(strDestFile);
			string strConn = oAdo.getMDBConnString(strDestFile,"admin","");
			oAdo.OpenConnection(strConn);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioCostsTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName);
			frmMain.g_oTables.m_oScenario.CreateScenarioDatasourceTable(oAdo,oAdo.m_OleDbConnection,Tables.Scenario.DefaultScenarioDatasourceTableName);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioHarvestCostColumnsTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioHarvestCostColumnsTableName);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioLandOwnerGroupsTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableName);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioMergeTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioMergeTableName);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioPlotFilterMiscTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterMiscTableName);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioPlotFilterTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterTableName);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioPSitesTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableName);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioLastTieBreakRankTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLastTieBreakRankTableName);
			frmMain.g_oTables.m_oScenario.CreateScenarioTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableName);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesOverallEffectiveTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesOptimizationTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesTieBreakerTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioCondFilterMiscTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableName);
			frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioCondFilterTable(oAdo,oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterTableName);
            frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioProcessorScenarioSelectTable(oAdo, oAdo.m_OleDbConnection, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName);

            //create empty prepost_fvs_weighted.accdb
            strDestFile = this.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile;
            oDao.CreateMDB(strDestFile);
            
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
			oDao = null;
		}

		public void CreateOptimizerScenarioRuleDefinitionSqliteDbAndTables(string p_strPathAndFile)
        {
			DataMgr dataMgr = new DataMgr();

			dataMgr.CreateDbFile(p_strPathAndFile);
			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dataMgr.GetConnectionString(p_strPathAndFile)))
            {
				conn.Open();
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioCostsTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName);
				frmMain.g_oTables.m_oScenario.CreateSqliteScenarioDatasourceTable(dataMgr, conn, Tables.Scenario.DefaultScenarioDatasourceTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioHarvestCostColumnsTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioHarvestCostColumnsTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioLandOwnerGroupsTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioPlotFilterTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioPSitesTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioLastTieBreakRankTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLastTieBreakRankTableName);
				frmMain.g_oTables.m_oScenario.CreateSqliteScenarioTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioFVSVariablesTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioFVSVariableOverallEffectiveTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioFVSVariablesOptimizationTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioFVSVariablesTieBreakerTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioCondFilterMiscTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioCondFilterTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioProcessorScenarioSelectTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName);
				conn.Close();
            }

			//create empty prepost_fvs_weighted.db
			string strDestFile = this.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile;
			if (!System.IO.File.Exists(strDestFile))
            {
				dataMgr.CreateDbFile(strDestFile);
			}
		}

		public void CreateProcessorScenarioRuleDefinitionAccessDbAndTables(string p_strPathAndFile)
		{
			dao_data_access oDao = new dao_data_access();
			ado_data_access oAdo = new ado_data_access();

			string strDestFile = p_strPathAndFile;
			oDao.CreateMDB(strDestFile);
			string strConn = oAdo.getMDBConnString(strDestFile,"admin","");
			oAdo.OpenConnection(strConn);
			frmMain.g_oTables.m_oScenario.CreateScenarioDatasourceTable(oAdo,oAdo.m_OleDbConnection,Tables.Scenario.DefaultScenarioDatasourceTableName);
			frmMain.g_oTables.m_oScenario.CreateScenarioTable(oAdo,oAdo.m_OleDbConnection,Tables.Scenario.DefaultScenarioTableName);
			frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioTreeSpeciesDollarValuesTable(oAdo,oAdo.m_OleDbConnection,Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName);
			frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioHarvestMethodTable(oAdo,oAdo.m_OleDbConnection,Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName);
			frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioCostRevenueEscalatorsTable(oAdo,oAdo.m_OleDbConnection,Tables.ProcessorScenarioRuleDefinitions.DefaultCostRevenueEscalatorsTableName);
            frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioHarvestCostColumnsTable(oAdo, oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestCostColumnsTableName);
            frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioAdditionalHarvestCostsTable(oAdo, oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName);
            frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioMoveInCostsTable(oAdo, oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName);
            frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioTreeDiamGroupsTable(oAdo, oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName);
            frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioTreeSpeciesGroupsListTable(oAdo, oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName);
            frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioTreeSpeciesGroupsTable(oAdo, oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName);
			oAdo.CloseConnection(oAdo.m_OleDbConnection);

			oDao = null;
		}

        public void CreateProcessorScenarioRuleDefinitionDbAndTables(string p_strPathAndFile)
        {
            DataMgr dataMgr = new DataMgr();

            dataMgr.CreateDbFile(p_strPathAndFile);
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dataMgr.GetConnectionString(p_strPathAndFile)))
            {
                conn.Open();
                frmMain.g_oTables.m_oScenario.CreateSqliteScenarioDatasourceTable(dataMgr, conn, Tables.Scenario.DefaultScenarioDatasourceTableName);
                frmMain.g_oTables.m_oScenario.CreateSqliteScenarioTable(dataMgr, conn, Tables.Scenario.DefaultScenarioTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateSqliteScenarioTreeSpeciesDollarValuesTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateSqliteScenarioHarvestMethodTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateSqliteScenarioCostRevenueEscalatorsTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultCostRevenueEscalatorsTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateSqliteScenarioHarvestCostColumnsTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestCostColumnsTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateSqliteScenarioAdditionalHarvestCostsTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateSqliteScenarioMoveInCostsTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateSqliteScenarioTreeDiamGroupsTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateSqliteScenarioTreeSpeciesGroupsListTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateSqliteScenarioTreeSpeciesGroupsTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName);
            }
        }

        private void btnSharedDirectory_Click(object sender, System.EventArgs e)
		{
		    string strSourceFile="";
			string strDestFile="";
			string strNewShared="";
			string strConn;
			DialogResult result2 = new DialogResult();
			//prompt for the location of the shared project.mdb table used for notes and document links
			DialogResult result =  this.folderBrowserDialog1.ShowDialog();
			
			if (result == DialogResult.OK) 
			{
				dao_data_access oDao = new dao_data_access();
				ado_data_access oAdo = new ado_data_access();
				string strTemp = this.folderBrowserDialog1.SelectedPath;
			
				if (strTemp.Length > 0) 
				{
					strNewShared=strTemp;
					//check if the selected directory is the project directory
					if (strTemp.Trim().ToUpper() == this.m_strProjectDirectory.Trim().ToUpper() + "\\DB")
					{
						//cannot overwrite the main project table so just 
						//designate it as the shared document link and notes table
						this.txtShared.Text = strTemp;
					}
					else if (strTemp.Trim().ToUpper() == this.txtPersonal.Text.Trim().ToUpper()) 
					{
						MessageBox.Show("!!Personal Directory And Shared Directory Cannot Be The Same!!","Personal Notes And Document Links",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					}
					else 
					{

						//see if we currently have a shared directory
						
						if (this.m_strSharedDirectory.Trim().Length == 0)
						{
							//no current shared directory so copy an empty project.mdb file to the 
							//new shared directory
							
							strDestFile = strNewShared + "\\shared_project_links_and_notes.mdb";
							if (System.IO.File.Exists(strDestFile)==false)
							{
								this.txtShared.Text = strNewShared;
								//create new shared project links and documents db file
								oDao.CreateMDB(strDestFile);
								strConn = oAdo.getMDBConnString(strDestFile,"admin","");
								oAdo.OpenConnection(strConn);
								frmMain.g_oTables.m_oProject.CreateProjectLinksCategoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksCategoryTableName);
								frmMain.g_oTables.m_oProject.CreateProjectLinksDepositoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksDepositoryTableName);
								frmMain.g_oTables.m_oProject.CreateProjectNotesTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectNotesTableName);
								while (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
								{
									oAdo.m_OleDbConnection.Close();
								}



							}
							else
							{
                                
								//a project file already exists in the directory
                                result = MessageBox.Show("A project notes and document links file already exists.\r\n\r\n NOTE: Previously saved shared notes and document links will be overwritten if choosing <Yes>.\r\n\r\nOverwrite the file? (Y/N)", "Shared Notes And Document Links", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
								switch (result) 
								{
									case DialogResult.Yes:
										System.IO.File.Delete(strDestFile);
										this.txtShared.Text = strNewShared;
										oDao.CreateMDB(strDestFile);
										strConn = oAdo.getMDBConnString(strDestFile,"admin","");
										oAdo.OpenConnection(strConn);
										frmMain.g_oTables.m_oProject.CreateProjectLinksCategoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksCategoryTableName);
										frmMain.g_oTables.m_oProject.CreateProjectLinksDepositoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksDepositoryTableName);
										frmMain.g_oTables.m_oProject.CreateProjectNotesTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectTableName);
										oAdo.CloseConnection(oAdo.m_OleDbConnection);
										break;
									case DialogResult.No:
                                        result = MessageBox.Show("Reference the existing notes and document links file (Y/N)", "Shared Notes And Document Links", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                        if (result==DialogResult.OK)
                                            this.txtShared.Text = strNewShared;
										break;
								}                
							}
						}
						else
						{
							//we have a shared directory already
							
							strDestFile = strNewShared + "\\shared_project_links_and_notes.mdb";
							if (System.IO.File.Exists(strDestFile)==false)
							{
								//destination file does not exist
								//a current shared notes and document links project.mdb exists 
								//  so lets prompt user whether to copy the previous
								//  project.mdb or a new one.
								result2 = MessageBox.Show("A Current Shared Notes and Document Links table exists. Do you want to copy it to the new location?(Y/N)", "Shared Notes And Document Links", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
								switch (result2) 
								{
									case DialogResult.Yes:
										this.txtShared.Text = strNewShared;
										strSourceFile = this.m_strSharedDirectory + "\\shared_project_links_and_notes.mdb";
										if (System.IO.File.Exists(strSourceFile.Trim())==false)
										{
											strSourceFile = this.m_oEnv.strAppDir + "\\db\\shared_project_links_and_notes.mdb";
										}
										System.IO.File.Copy(strSourceFile, strDestFile,true);
										break;
									case DialogResult.No:
										this.txtShared.Text = strNewShared;
										oDao.CreateMDB(strDestFile);
										strConn = oAdo.getMDBConnString(strDestFile,"admin","");
										oAdo.OpenConnection(strConn);
										frmMain.g_oTables.m_oProject.CreateProjectLinksCategoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksCategoryTableName);
										frmMain.g_oTables.m_oProject.CreateProjectLinksDepositoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksDepositoryTableName);
										frmMain.g_oTables.m_oProject.CreateProjectNotesTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectNotesTableName);
										while (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
										{
											oAdo.m_OleDbConnection.Close();
										}							
										break;
								}                
								

							}
							else
							{
								//destination file already exists.
								//a current shared notes and document links project.mdb exists 
								//  and a notes and document links project.mdb exists in the 
								//  newly designated directory 
								strDestFile = strNewShared + "\\shared_project_links_and_notes.mdb";
								this.m_frmDialog1 =  new frmDialog();
								this.m_frmDialog1.Height = 100;
								
								this.m_frmDialog1.Width = 100;
								this.m_frmDialog1.MinimizeBox = false;
								this.m_frmDialog1.MaximizeBox = false;
								this.m_frmDialog1.WindowState=System.Windows.Forms.FormWindowState.Normal;
								this.m_frmDialog1.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
								//label message
								System.Windows.Forms.Label lblMsg = new System.Windows.Forms.Label();
								lblMsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Regular))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
								lblMsg.ForeColor = System.Drawing.Color.Black;
								lblMsg.Location = new System.Drawing.Point(8, 16);
								lblMsg.Size = new System.Drawing.Size(208, 24);
								this.m_frmDialog1.Controls.Add(lblMsg);

								lblMsg.Name = "label1";
								lblMsg.Top = 2;
								lblMsg.Left = 2;
								lblMsg.Width = 200;
								lblMsg.Height = 50;

								lblMsg.Text = "The selected shared directory already has a notes and document link database." + 
									" Do you want to overwrite it with a new database file, or overwrite it with the current database " + 
									" file, or keep the existing database file?";
								lblMsg.Visible=true;
								//copy current button
								System.Windows.Forms.Button btnCopyCurrent = new System.Windows.Forms.Button();
								btnCopyCurrent.Name="btnCopyCurrent";
								btnCopyCurrent.FlatStyle = System.Windows.Forms.FlatStyle.Standard;
								btnCopyCurrent.BackColor = System.Drawing.SystemColors.Control;
								btnCopyCurrent.Location = new System.Drawing.Point(8, 256);
								btnCopyCurrent.Size = new System.Drawing.Size(128, 24);
								btnCopyCurrent.TabIndex = 4;
								btnCopyCurrent.Text = "Copy Current";
								btnCopyCurrent.Click += new EventHandler(this.m_frmDialog1_btnPressed_Click);
								this.m_frmDialog1.Controls.Add(btnCopyCurrent);

								//copy new

								System.Windows.Forms.Button btnCopyNew = new System.Windows.Forms.Button();
								btnCopyNew.Name="btnCopyNew";
								btnCopyNew.FlatStyle = System.Windows.Forms.FlatStyle.Standard;
								btnCopyNew.BackColor = System.Drawing.SystemColors.Control;
								btnCopyNew.Location = new System.Drawing.Point(8, 256);
								btnCopyNew.Size = new System.Drawing.Size(128, 24);
								btnCopyNew.TabIndex = 4;
								btnCopyNew.Text = "Copy New";
								btnCopyNew.Click += new EventHandler(this.m_frmDialog1_btnPressed_Click);
								this.m_frmDialog1.Controls.Add(btnCopyNew);


								System.Windows.Forms.Button btnKeep = new System.Windows.Forms.Button();
								btnKeep.Name="btnKeep";
								btnKeep.FlatStyle = System.Windows.Forms.FlatStyle.Standard;
								btnKeep.BackColor = System.Drawing.SystemColors.Control;
								btnKeep.Location = new System.Drawing.Point(8, 256);
								btnKeep.Size = new System.Drawing.Size(128, 24);
								btnKeep.TabIndex = 4;
								btnKeep.Text = "Keep";
								btnKeep.Click += new EventHandler(this.m_frmDialog1_btnPressed_Click);
								this.m_frmDialog1.Controls.Add(btnKeep);


								System.Windows.Forms.Button btnCancel = new System.Windows.Forms.Button();
								btnCancel.Name="btnCancel";
								btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Standard;
								btnCancel.BackColor = System.Drawing.SystemColors.Control;
								btnCancel.Location = new System.Drawing.Point(8, 256);
								btnCancel.Size = new System.Drawing.Size(128, 24);
								btnCancel.TabIndex = 4;
								btnCancel.Text = "Cancel";
								btnCancel.Click += new EventHandler(this.m_frmDialog1_btnPressed_Click);
								this.m_frmDialog1.Controls.Add(btnCancel);



								lblMsg.Width = btnCopyCurrent.Width + btnCopyNew.Width + btnKeep.Width + btnCancel.Width;
								btnCopyCurrent.Left = 2;
								btnCopyNew.Left = 2 + btnCopyCurrent.Width;
								btnKeep.Left = btnCopyNew.Left + btnCopyNew.Width;
								btnCancel.Left = btnKeep.Left + btnKeep.Width;
								btnCopyCurrent.Top = lblMsg.Top + lblMsg.Height;
								btnCopyNew.Top = btnCopyCurrent.Top;
								btnKeep.Top = btnCopyCurrent.Top;
								btnCancel.Top = btnCopyCurrent.Top;
								this.m_frmDialog1.Width = lblMsg.Left + lblMsg.Width + 6;
								this.m_frmDialog1.Text = "Select A Shared Notes And Document Link Option";
								this.m_frmDialog1.Height = btnCopyCurrent.Top + (int)(btnCopyCurrent.Height * 1.8) + 10;

								result2 = this.m_frmDialog1.ShowDialog();
								switch (result2)
								{
									case DialogResult.OK:
									switch (m_strAction)
									{
										case "COPY CURRENT":
											this.txtShared.Text = strNewShared;
											strSourceFile = this.m_strSharedDirectory + "\\shared_project_links_and_notes.mdb";
											System.IO.File.Copy(strSourceFile, strDestFile,true);
											break;
										case "COPY NEW":
											oDao.CreateMDB(strDestFile);
											strConn = oAdo.getMDBConnString(strDestFile,"admin","");
											oAdo.OpenConnection(strConn);
											frmMain.g_oTables.m_oProject.CreateProjectLinksCategoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksCategoryTableName);
											frmMain.g_oTables.m_oProject.CreateProjectLinksDepositoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksDepositoryTableName);
											frmMain.g_oTables.m_oProject.CreateProjectNotesTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectNotesTableName);
											while (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
											{
												oAdo.m_OleDbConnection.Close();
											}
											break;
										case "KEEP":
											this.txtShared.Text = strNewShared;
											break;
										default:
											break;
									}
										break;
									default:
										break;

								}
							}
						}
					}
					if (this.txtShared.Text.Trim().ToUpper() == strNewShared.Trim().ToUpper())
					{
						((frmDialog)this.ParentForm).uc_project1.btnSave.Enabled=true;
					}
				}
				oDao=null;
				oAdo=null;
			}

		}
		public  void Open_Project()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//uc_project.Open_Profect \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			//string  strTemp;
			//int x;
			this.m_strNewProjectFile = "";
			this.m_strNewProjectDirectory = "";
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "Open FIA Biosum Project Access File";
			OpenFileDialog1.Filter = "MS Access Database File (*.MDB,*.MDE,*.ACCDB) |*.mdb;*.mde;*.accdb";
			
			DialogResult result =  OpenFileDialog1.ShowDialog();
			this.m_intError=0;
			this.m_strError="";
			if (result == DialogResult.OK) 
			{
				
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					this.m_strNewProjectFile = OpenFileDialog1.FileName.Substring(OpenFileDialog1.FileName.LastIndexOf("\\") + 1);
					this.m_strNewProjectDirectory = OpenFileDialog1.FileName.Substring(0,OpenFileDialog1.FileName.LastIndexOf("\\") - 3);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.Open_Profect: Open Project Table \r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.Open_Profect: strNewProjectDirectory=" + m_strNewProjectDirectory + " strNewProjectFile=" + m_strNewProjectFile + "\r\n");
                    }
					this.OpenProjectTable(this.m_strNewProjectDirectory, this.m_strNewProjectFile);
                    if (this.m_intError == 0)
                        this.OpenUserConfigTable(this.m_strNewProjectDirectory + "\\db", this.m_strNewProjectFile);
                    else
                    {

                    }


				}
			}
			else 
			{
				this.m_intError = -1;
			}
			OpenFileDialog1 = null;
			

		}

		public  void Open_Project_No_Dialog(string strDirectoryAndFile)
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//uc_project.Open_Profect_No_Dialog \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			
			this.m_strNewProjectFile = "";
			this.m_strNewProjectDirectory = "";
			this.m_intError=0;
			this.m_strError="";
			
			
			
			if (strDirectoryAndFile.Trim().Length > 0) 
			{
				this.m_strNewProjectFile =   strDirectoryAndFile.Substring(strDirectoryAndFile.LastIndexOf("\\") + 1);    //OpenFileDialog1.FileName.Substring(OpenFileDialog1.FileName.LastIndexOf("\\") + 1);
				this.m_strNewProjectDirectory = strDirectoryAndFile.Substring(0,strDirectoryAndFile.LastIndexOf("\\") - 3);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.Open_Profect_No_Dialog: Open Project Table \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.Open_Profect_No_Dialog: strNewProjectDirectory=" + m_strNewProjectDirectory + " strNewProjectFile=" + m_strNewProjectFile + "\r\n");
                }
                    
                
				this.OpenProjectTable(this.m_strNewProjectDirectory, this.m_strNewProjectFile);
                if (this.m_intError == 0)
                    this.OpenUserConfigTable(this.m_strNewProjectDirectory + "\\db", this.m_strNewProjectFile);
                
			}
			else 
			{
				this.m_intError=-1;
			}
			

		}
		private void txtProjectId_Leave(object sender, System.EventArgs e)
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//uc_project.txtProjectId_Leave \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
			if (this.txtProjectId.Text.Length > 0 && txtProjectId.Enabled==true) 
			{
				this.txtProjectId.Text = this.txtProjectId.Text.Trim();
				//replace spaces with underscores
				this.txtProjectId.Text = this.txtProjectId.Text.Replace(" ","_");

                if (this.m_strAction == "NEW")
                {
                    if (txtRootDirectory.Text.Trim().Length == 0)
                    {
                        this.txtRootDirectory.Text = this.m_oEnv.strAppDir.Substring(0, 2) + "\\FIA_Biosum\\" + this.txtProjectId.Text.ToLower();
                    }
                    else
                    {
                        if (txtRootDirectory.Text.Trim().Substring(txtRootDirectory.Text.Trim().Length - 1, 1) == @"\")
                            this.txtRootDirectory.Text = this.txtRootDirectory.Text.Trim() + this.txtProjectId.Text.ToLower();
                        else
                            this.txtRootDirectory.Text = this.txtRootDirectory.Text.Trim() + "\\" + this.txtProjectId.Text.ToLower();
                    }
                }
			}
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.txtProjectId_Leave: txtProjectId.Text=" + txtProjectId.Text.Trim() + " txtRootDirectory.Text=" + txtRootDirectory.Text.Trim() + "\r\n");
		}

		private void grpboxProjectId_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void grpboxCreated_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void grpboxDescription_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void uc_project_Resize(object sender, System.EventArgs e)
		{
			resize_uc_project();
			

		

		}
		public void resize_uc_project()
		{
			this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
			this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
			this.grpboxProjectId.Top = this.lblTitle.Top + this.lblTitle.Height + 2;
			this.grpboxCreated.Top = this.grpboxProjectId.Top;
			this.grpboxCompany.Top = this.grpboxProjectId.Top + this.grpboxProjectId.Height + 2;
			this.grpboxDescription.Top = this.grpboxCompany.Top + this.grpboxCompany.Height + 2;
			this.grpboxProjectFiles.Top = this.grpboxDescription.Top + this.grpboxDescription.Height + 2;
			
			this.grpboxDescription.Left = 2;
			this.grpboxProjectFiles.Left = 2;
			this.grpboxProjectId.Left = 2;
			this.grpboxCompany.Left = 2;
			this.grpboxDescription.Width  = this.Width - 4;
			this.grpboxProjectFiles.Width = this.grpboxDescription.Width;
			this.grpboxCompany.Width = this.grpboxDescription.Width;
			

			this.lblRootDirectory.Left = this.grpboxProjectFiles.Left + 2;
			this.lblShared.Left = this.lblRootDirectory.Left;
			this.lblLocal.Left = this.lblRootDirectory.Left;

			this.txtRootDirectory.Left = this.lblRootDirectory.Left + this.lblRootDirectory.Width + 2;
			this.txtShared.Left = this.txtRootDirectory.Left;
			this.txtPersonal.Left = this.txtRootDirectory.Left;

			this.btnRootDirectoryHelp.Left = this.grpboxProjectFiles.Width - this.btnRootDirectoryHelp.Width - 2;
			this.btnSharedHelp.Left = this.btnRootDirectoryHelp.Left;
			this.btnPersonalHelp.Left = this.btnRootDirectoryHelp.Left;

			this.btnRootDirectory.Left = this.btnRootDirectoryHelp.Left - this.btnRootDirectoryHelp.Width - 2;
			this.btnSharedDirectory.Left = this.btnRootDirectory.Left;
			this.btnPersonalDirectory.Left = this.btnRootDirectory.Left;

			this.btnSharedHelp.Top = this.btnSharedDirectory.Top;
			this.btnPersonalHelp.Top = this.btnPersonalDirectory.Top;
			this.btnRootDirectoryHelp.Top = this.btnRootDirectory.Top;

			this.txtRootDirectory.Width = this.grpboxProjectFiles.Width  - this.lblRootDirectory.Width  - this.btnRootDirectory.Width - this.btnRootDirectoryHelp.Width   - 15;
			this.txtShared.Width = this.txtRootDirectory.Width;
			this.txtPersonal.Width = this.txtRootDirectory.Width;
			
			
			this.txtDescription.Left = this.grpboxDescription.Left + 2;
			this.txtDescription.Width = this.grpboxDescription.Width - 8;
			this.txtCompany.Left = this.grpboxCompany.Left + 2;
			this.txtCompany.Width = this.grpboxCompany.Width - 8;
			

			this.grpboxProjectFiles.Height = this.btnPersonalDirectory.Top + this.btnPersonalDirectory.Height + 4;

			this.btnCancel.Top = this.grpboxProjectFiles.Top + this.grpboxProjectFiles.Height + 5;
			this.btnCancel.Left = (int) (this.Width * .50) + (int) (this.btnCancel.Width / 2);
			this.btnSave.Top = this.btnCancel.Top;
			this.btnSave.Left = this.btnCancel.Left - this.btnCancel.Width;
			this.btnEdit.Top = this.btnCancel.Top;
			this.btnEdit.Left = this.btnSave.Left - this.btnSave.Width;
            this.btnHelp.Top = this.btnCancel.Top;
            this.btnHelp.Left = this.btnEdit.Left - this.btnEdit.Width;

		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
		   
		        this.ParentForm.Close();
		}

		private void btnPersonalDirectory_Click(object sender, System.EventArgs e)
		{

			string strSourceFile="";
			string strDestFile="";
			string strNewPersonal="";
			string strConn="";
			DialogResult result2 = new DialogResult();
			this.folderBrowserDialog1.SelectedPath = this.txtRootDirectory.Text;
			DialogResult result =  this.folderBrowserDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				dao_data_access oDao = new dao_data_access();
				ado_data_access oAdo = new ado_data_access();

				string strTemp = this.folderBrowserDialog1.SelectedPath;
			
				if (strTemp.Length > 0) 
				{
					strNewPersonal=strTemp;
					//check if the selected directory is the project directory
					if (strTemp.Trim().ToUpper() == this.m_strProjectDirectory.Trim().ToUpper())
					{
						//cannot overwrite the main project table so just 
						//designate it as the shared document link and notes table
						MessageBox.Show("!!Personal Directory And Project Directory Cannot Be The Same!!","Personal Notes And Document Links", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					}
					else if (strTemp.Trim().ToUpper() == this.txtShared.Text.Trim().ToUpper()) 
					{
						MessageBox.Show("!!Personal Directory And Shared Directory Cannot Be The Same!!","Personal Notes And Document Links", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					}
					else 
					{

						//see if we currently have a shared directory

						if (this.m_strPersonalDirectory.Trim().Length == 0)
						{
							//no current shared directory so copy an empty project.mdb file to the 
							//new shared directory
							//strDestFile = this.txtShared.Text + "\\project.mdb";
							strDestFile = strNewPersonal + "\\personal_project_links_and_notes.mdb";
							if (System.IO.File.Exists(strDestFile)==false)
							{
								this.txtPersonal.Text = strNewPersonal;
								oDao.CreateMDB(strDestFile);
								strConn = oAdo.getMDBConnString(strDestFile,"admin","");
								oAdo.OpenConnection(strConn);
								frmMain.g_oTables.m_oProject.CreateProjectLinksCategoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksCategoryTableName);
								frmMain.g_oTables.m_oProject.CreateProjectLinksDepositoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksDepositoryTableName);
								frmMain.g_oTables.m_oProject.CreateProjectNotesTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectNotesTableName);
								while (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
								{
									oAdo.m_OleDbConnection.Close();
								}
							}
							else
							{
								//a project file already exists in the directory
                                result = MessageBox.Show("A project notes and document links file already exists.\r\n\r\n NOTE: Previously saved personal notes and document links will be overwritten if choosing <Yes>.\r\n\r\nOverwrite the file? (Y/N)", "Personal Notes And Document Links", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
								switch (result) 
								{
									case DialogResult.Yes:
										System.IO.File.Delete(strDestFile);
										this.txtPersonal.Text = strNewPersonal;
										oDao.CreateMDB(strDestFile);
										strConn = oAdo.getMDBConnString(strDestFile,"admin","");
										oAdo.OpenConnection(strConn);
										frmMain.g_oTables.m_oProject.CreateProjectLinksCategoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksCategoryTableName);
										frmMain.g_oTables.m_oProject.CreateProjectLinksDepositoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksDepositoryTableName);
										frmMain.g_oTables.m_oProject.CreateProjectNotesTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectNotesTableName);
										while (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
										{
											oAdo.m_OleDbConnection.Close();
										}
                                        frmMain.g_oFrmMain.frmProject.uc_project_document_links1.loadvalues(
                                            txtShared.Text,
                                            txtPersonal.Text, true);
										break;
									case DialogResult.No:
                                         result = MessageBox.Show("Reference the existing notes and document links file (Y/N)", "Shared Notes And Document Links", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                         if (result == DialogResult.Yes)
                                         {
                                             this.txtPersonal.Text = strNewPersonal;
                                             frmMain.g_oFrmMain.frmProject.uc_project_document_links1.loadvalues(
                                             txtShared.Text,
                                             txtPersonal.Text, true);
                                         }
										break;
								}                
							}
						}
						else
						{
							//we have a shared directory already
							//strDestFile = this.txtShared.Text + "\\project.mdb";
							strDestFile = strNewPersonal + "\\personal_project_links_and_notes.mdb";
							if (System.IO.File.Exists(strDestFile)==false)
							{
								//destination file does not exist
								//a current shared notes and document links project.mdb exists 
								//  so lets prompt user whether to copy the previous
								//  project.mdb or a new one.
								result2 = MessageBox.Show("A Current Personal Notes and Document Links table exists. Do you want to copy it to the new location?(Y/N)", "Personal Notes And Document Links", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
								switch (result2) 
								{
									case DialogResult.Yes:
										this.txtPersonal.Text = strNewPersonal;
										strSourceFile = this.m_strPersonalDirectory + "\\personal_project_links_and_notes.mdb";
										if (System.IO.File.Exists(strSourceFile.Trim())==false)
										{
											oDao.CreateMDB(strDestFile);
											strConn = oAdo.getMDBConnString(strDestFile,"admin","");
											oAdo.OpenConnection(strConn);
											frmMain.g_oTables.m_oProject.CreateProjectLinksCategoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksCategoryTableName);
											frmMain.g_oTables.m_oProject.CreateProjectLinksDepositoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksDepositoryTableName);
											frmMain.g_oTables.m_oProject.CreateProjectNotesTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectNotesTableName);
											while (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
											{
												oAdo.m_OleDbConnection.Close();
											}
										}
										else
										{
											System.IO.File.Copy(strSourceFile, strDestFile,true);
										}
										break;
									case DialogResult.No:
										this.txtPersonal.Text = strNewPersonal;
										//strSourceFile = this.m_oEnv.strAppDir + "\\db\\personal_project_links_and_notes.mdb";
										//System.IO.File.Copy(strSourceFile, strDestFile,true);									
										oDao.CreateMDB(strDestFile);
										strConn = oAdo.getMDBConnString(strDestFile,"admin","");
										oAdo.OpenConnection(strConn);
										frmMain.g_oTables.m_oProject.CreateProjectLinksCategoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksCategoryTableName);
										frmMain.g_oTables.m_oProject.CreateProjectLinksDepositoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksDepositoryTableName);
										frmMain.g_oTables.m_oProject.CreateProjectNotesTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectNotesTableName);
										while (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
										{
											oAdo.m_OleDbConnection.Close();
										}
										break;
								}                
								

							}
							else
							{
								//destination file already exists.
								//a current shared notes and document links project.mdb exists 
								//  and a notes and document links project.mdb exists in the 
								//  newly designated directory 
								strDestFile = strNewPersonal + "\\personal_project_links_and_notes.mdb";
								this.m_frmDialog1 =  new frmDialog();
								this.m_frmDialog1.Height = 100;
								
								this.m_frmDialog1.Width = 100;
								this.m_frmDialog1.MinimizeBox = false;
								this.m_frmDialog1.MaximizeBox = false;
								this.m_frmDialog1.WindowState=System.Windows.Forms.FormWindowState.Normal;
								this.m_frmDialog1.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
								//label message
								System.Windows.Forms.Label lblMsg = new System.Windows.Forms.Label();
								lblMsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Regular))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
								lblMsg.ForeColor = System.Drawing.Color.Black;
								lblMsg.Location = new System.Drawing.Point(8, 16);
								lblMsg.Size = new System.Drawing.Size(208, 24);
								this.m_frmDialog1.Controls.Add(lblMsg);

								lblMsg.Name = "label1";
								lblMsg.Top = 2;
								lblMsg.Left = 2;
								lblMsg.Width = 200;
								lblMsg.Height = 50;

								lblMsg.Text = "The selected personal directory already has a notes and document link database." + 
									" Do you want to overwrite it with a new database file, or overwrite it with the current database " + 
									" file, or keep the existing database file?";
								lblMsg.Visible=true;
								//copy current button
								System.Windows.Forms.Button btnCopyCurrent = new System.Windows.Forms.Button();
								btnCopyCurrent.Name="btnCopyCurrent";
								btnCopyCurrent.FlatStyle = System.Windows.Forms.FlatStyle.Standard;
								btnCopyCurrent.BackColor = System.Drawing.SystemColors.Control;
								btnCopyCurrent.Location = new System.Drawing.Point(8, 256);
								btnCopyCurrent.Size = new System.Drawing.Size(128, 24);
								btnCopyCurrent.TabIndex = 4;
								btnCopyCurrent.Text = "Copy Current";
								btnCopyCurrent.Click += new EventHandler(this.m_frmDialog1_btnPressed_Click);
								this.m_frmDialog1.Controls.Add(btnCopyCurrent);

								//copy new

								System.Windows.Forms.Button btnCopyNew = new System.Windows.Forms.Button();
								btnCopyNew.Name="btnCopyNew";
								btnCopyNew.FlatStyle = System.Windows.Forms.FlatStyle.Standard;
								btnCopyNew.BackColor = System.Drawing.SystemColors.Control;
								btnCopyNew.Location = new System.Drawing.Point(8, 256);
								btnCopyNew.Size = new System.Drawing.Size(128, 24);
								btnCopyNew.TabIndex = 4;
								btnCopyNew.Text = "Copy New";
								btnCopyNew.Click += new EventHandler(this.m_frmDialog1_btnPressed_Click);
								this.m_frmDialog1.Controls.Add(btnCopyNew);


								System.Windows.Forms.Button btnKeep = new System.Windows.Forms.Button();
								btnKeep.Name="btnKeep";
								btnKeep.FlatStyle = System.Windows.Forms.FlatStyle.Standard;
								btnKeep.BackColor = System.Drawing.SystemColors.Control;
								btnKeep.Location = new System.Drawing.Point(8, 256);
								btnKeep.Size = new System.Drawing.Size(128, 24);
								btnKeep.TabIndex = 4;
								btnKeep.Text = "Keep";
								btnKeep.Click += new EventHandler(this.m_frmDialog1_btnPressed_Click);
								this.m_frmDialog1.Controls.Add(btnKeep);


								System.Windows.Forms.Button btnCancel = new System.Windows.Forms.Button();
								btnCancel.Name="btnCancel";
								btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Standard;
								btnCancel.BackColor = System.Drawing.SystemColors.Control;
								btnCancel.Location = new System.Drawing.Point(8, 256);
								btnCancel.Size = new System.Drawing.Size(128, 24);
								btnCancel.TabIndex = 4;
								btnCancel.Text = "Cancel";
								btnCancel.Click += new EventHandler(this.m_frmDialog1_btnPressed_Click);
								this.m_frmDialog1.Controls.Add(btnCancel);



								lblMsg.Width = btnCopyCurrent.Width + btnCopyNew.Width + btnKeep.Width + btnCancel.Width;
								btnCopyCurrent.Left = 2;
								btnCopyNew.Left = 2 + btnCopyCurrent.Width;
								btnKeep.Left = btnCopyNew.Left + btnCopyNew.Width;
								btnCancel.Left = btnKeep.Left + btnKeep.Width;
								btnCopyCurrent.Top = lblMsg.Top + lblMsg.Height;
								btnCopyNew.Top = btnCopyCurrent.Top;
								btnKeep.Top = btnCopyCurrent.Top;
								btnCancel.Top = btnCopyCurrent.Top;
								this.m_frmDialog1.Width = lblMsg.Left + lblMsg.Width + 6;
								this.m_frmDialog1.Text = "Select A Personal Notes And Document Link Option";
								this.m_frmDialog1.Height = btnCopyCurrent.Top + (int)(btnCopyCurrent.Height * 1.8) + 10;

								result2 = this.m_frmDialog1.ShowDialog();
								switch (result2)
								{
									case DialogResult.OK:
									switch (m_strAction)
									{
										case "COPY CURRENT":
											this.txtPersonal.Text = strNewPersonal;
											strSourceFile = this.m_strSharedDirectory + "\\personal_project_links_and_notes.mdb";
											System.IO.File.Copy(strSourceFile, strDestFile,true);
											break;
										case "COPY NEW":
											oDao.CreateMDB(strDestFile);
											strConn = oAdo.getMDBConnString(strDestFile,"admin","");
											oAdo.OpenConnection(strConn);
											frmMain.g_oTables.m_oProject.CreateProjectLinksCategoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksCategoryTableName);
											frmMain.g_oTables.m_oProject.CreateProjectLinksDepositoryTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectLinksDepositoryTableName);
											frmMain.g_oTables.m_oProject.CreateProjectNotesTable(oAdo,oAdo.m_OleDbConnection,frmMain.g_oTables.m_oProject.DefaultProjectNotesTableName);
											while (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
											{
												oAdo.m_OleDbConnection.Close();
											}
											break;
										case "KEEP":
											this.txtPersonal.Text = strNewPersonal;
											break;
										default:
											break;
									}
										break;
									default:
										break;

								}
							}
						}
					}
					if (this.txtPersonal.Text.Trim().ToUpper() == strNewPersonal.Trim().ToUpper())
					{
						((frmDialog)this.ParentForm).uc_project1.btnSave.Enabled=true;
					}
				}
				oAdo=null;
				oDao=null;

			}

		}
		private void m_frmDialog1_btnPressed_Click(object sender, System.EventArgs e)
		{
			
			if (sender.ToString().ToUpper().IndexOf("COPY CURRENT") > 0)
			{
				this.m_strAction="COPY CURRENT";
				this.m_frmDialog1.DialogResult = System.Windows.Forms.DialogResult.OK;
			}
			else if (sender.ToString().ToUpper().IndexOf("COPY NEW") > 0)
			{
				this.m_strAction= "COPY NEW";
				this.m_frmDialog1.DialogResult = System.Windows.Forms.DialogResult.OK;
			}
			else if (sender.ToString().ToUpper().IndexOf("KEEP") > 0)
			{
				this.m_strAction="KEEP";
				this.m_frmDialog1.DialogResult = System.Windows.Forms.DialogResult.OK;
			}  
			else
			{
				this.m_strAction="";
				this.m_frmDialog1.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			}
			
		}
		
		public void SetProjectPathEnvironmentVariables()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//uc_project.SetProjectPathEnvironmentVariables \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            int x;
			
			string strFullPath = "";
			
			string strConn = "";
			string strSQL = "";
			string strOldProjDir = "";
            string strProjDir = "";
            
            frmMain.g_oGeneralMacroSubstitutionVariable_Collection.Item(frmMain.OLDPROJDIR).VariableSubstitutionString = this.txtRootDirectory.Text.Trim();
            frmMain.g_oGeneralMacroSubstitutionVariable_Collection.Item(frmMain.PROJDIR).VariableSubstitutionString = this.m_strProjectDirectory.Trim();

            strProjDir = m_strProjectDirectory.Trim();
            strOldProjDir = this.txtRootDirectory.Text.Trim();

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Replace old project directory (" + strOldProjDir + ") with new project directory (" + strProjDir + ")\r\n");
            

            /**********************************************
			 **instantiate the ado_data_access class
			 **********************************************/
			ado_data_access oAdo = new ado_data_access();
            DataMgr oDataMgr = new DataMgr();
            //
            //PROJECT DATA SOURCE
            //
            strFullPath = strProjDir + "\\db\\" + this.m_strProjectFile;
            strConn = oAdo.getMDBConnString(strFullPath, "", "");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Open Connection to Project Dbfile " + strConn + ")\r\n");
            oAdo.OpenConnection(strConn);
            
            strSQL = "UPDATE project SET project_root_directory = '" + strProjDir + "' " +
                     "WHERE proj_id = '" + this.txtProjectId.Text.Trim() + "';";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Execute SQL \r\n" + strSQL + "\r\n");

            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);

            strSQL = "UPDATE datasource " + 
                     "SET path = REPLACE(TRIM(LCASE(path))," + 
                                "'" + strOldProjDir.Trim().ToLower() + "'," + 
                                "'" + strProjDir.Trim().ToLower() + "')";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Execute SQL \r\n" + strSQL + "\r\n");

            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            //
            //TREATMENT OPTIMIZER SCENARIO DATA SOURCE
            //
            strFullPath = strProjDir + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
            if (System.IO.File.Exists(strFullPath))
            {
				strConn = oDataMgr.GetConnectionString(strFullPath);
				if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
					frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Open Connection to Treatment Optimizer Scenario Dbfile " + strConn + ")\r\n");

				using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
					conn.Open();
					strSQL = "UPDATE scenario_datasource " +
					 "SET path = REPLACE(TRIM(LOWER(path))," +
								"'" + strOldProjDir.Trim().ToLower() + "'," +
								"'" + strProjDir.Trim().ToLower() + "')";
					if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
						frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Execute SQL \r\n" + strSQL + "\r\n");
					oDataMgr.SqlNonQuery(conn, strSQL);
					strSQL = "UPDATE scenario " +
					"SET path = REPLACE(TRIM(LOWER(path))," +
							   "'" + strOldProjDir.Trim().ToLower() + "'," +
							   "'" + strProjDir.Trim().ToLower() + "')";
					if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
						frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Execute SQL \r\n" + strSQL + "\r\n");
					oDataMgr.SqlNonQuery(conn, strSQL);
				}
			}
            //
            //PROCESSOR SCENARIO DATA SOURCE
            //
            strFullPath = $@"{strProjDir}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile}";
            if (System.IO.File.Exists(strFullPath))
            {
                strConn = oDataMgr.GetConnectionString(strFullPath);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Open Connection to Processor Scenario Dbfile " + strConn + ")\r\n");
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    strSQL = "UPDATE scenario_datasource " +
                         "SET path = REPLACE(TRIM(LOWER(path))," +
                                    "'" + strOldProjDir.Trim().ToLower() + "'," +
                                    "'" + strProjDir.Trim().ToLower() + "')";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Execute SQL \r\n" + strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(conn, strSQL);
                    strSQL = "UPDATE scenario " +
                         "SET path = REPLACE(TRIM(LOWER(path))," +
                                    "'" + strOldProjDir.Trim().ToLower() + "'," +
                                    "'" + strProjDir.Trim().ToLower() + "')";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Execute SQL \r\n" + strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(conn, strSQL);
                }
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: frmMain.g_oUtils.getDriveLetter for project \r\n");
            m_strProjectDirectoryDrive = frmMain.g_oUtils.getDriveLetter(strProjDir);

            this.txtRootDirectory.Text = strProjDir;


            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Leaving \r\n");
			
		    oAdo = null;


 		}

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            //m_oHelp.ShowHelp(new string[] { "DATABASE", "NEWPROJECT" });
            m_oHelp.ShowPdfHelp(m_pdfFile, "1");
        }

		
	}
}



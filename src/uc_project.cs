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


		//project variables
		public bool boolProjectOpen = false;
        public string m_strDebugFile = "";

		//new project variables
		public string m_strNewProjectFile = "";
		public string m_strNewProjectDirectory = "";
		public string m_strNewProjectId="";
        public string m_strNewName="";
		public string m_strNewDate="";
		public string m_strNewOrganization="";
		public string m_strNewDescription="";
		public string m_strNewRootDirectory="";
		public string m_strNewProjectVersion="";

		public string m_strOldProjectDirectory = "";

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
		private System.Windows.Forms.GroupBox grpboxOrganization;
		public System.Windows.Forms.TextBox txtOrganization;
		private System.Windows.Forms.GroupBox grpboxProjectId;
		public System.Windows.Forms.TextBox txtProjectId;
		private System.Windows.Forms.GroupBox grpboxCreated;
		public System.Windows.Forms.TextBox txtDate;
		private System.Windows.Forms.Label lblDate;
		public System.Windows.Forms.TextBox txtName;
		private System.Windows.Forms.Label lblName;
		private System.Windows.Forms.GroupBox grpboxProjectDirectory;
        private System.Windows.Forms.Button btnRootDirectoryHelp;
		private System.Windows.Forms.Button btnRootDirectory;
		private System.Windows.Forms.Label lblRootDirectory;
		public System.Windows.Forms.TextBox txtRootDirectory;
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
			this.m_intError = 0;
			this.m_strError="";
            // Set the control height to accomodate the lowest button
            this.m_intFullHt = this.btnClose.Location.Y + this.btnClose.Size.Height + 10;
            // Set the control width to accomodate the furthest right groupbox
            this.m_intFullWh = this.grpboxCreated.Location.X + this.grpboxCreated.Size.Width + 30;

			this.txtRootDirectory.Enabled=false;
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

			try
			{
				p_ado.OpenConnection(strConn);
			}
			catch (Exception e)
            {
				string er = e.StackTrace;
            }

           
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
		            m_strNewOrganization=p_ado.m_OleDbDataReader["company"].ToString().Trim();
		            m_strNewDescription=p_ado.m_OleDbDataReader["description"].ToString().Trim();
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
				this.grpboxProjectDirectory.Enabled=false;
				this.grpboxProjectId.Enabled=false;
				this.grpboxOrganization.Enabled=false;
				this.grpboxCreated.Enabled=false;
				this.btnEdit.Enabled=true;
				this.btnCancel.Enabled=false;
                this.btnSave.Enabled=false;
                
			}
			frmMain.g_sbpInfo.Text = "Ready";
		}

		public void OpenProjectTableNew(string strProjDir, string strFile)
        {
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
			{
				frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
				frmMain.g_oUtils.WriteText(m_strDebugFile, "//uc_project.OpenProjectTable \r\n");
				frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
			}
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
				frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.OpenProjectTable: strRootDir=" + strProjDir + " strFile=" + strFile + "\r\n");
			frmMain.g_sbpInfo.Text = "Loading Project...Stand By";
			this.m_intError = 0;
			this.m_strError = "";
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
				frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.OpenProjectTable: Instantiate DataMgr \r\n");

			DataMgr p_dataMgr = new DataMgr();

			string strFullPath = strProjDir + "\\DB\\" + strFile;
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
				frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.OpenProjectTable: strFullPath=" + strFullPath + "\r\n");

			string strConn = p_dataMgr.GetConnectionString(strFullPath);

			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
				frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.OpenProjectTable: Open DBFile with Connection String=" + strConn + "\r\n");

			using (var oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
				oConn.Open();

				if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
					frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.OpenProjectTable: OpenConnection error Value=" + p_dataMgr.m_intError.ToString() + "\r\n");

				if (p_dataMgr.m_intError == 0)
                {
					try
					{
						bool bAppVerColumnExist = p_dataMgr.ColumnExist(oConn, "project", "application_version");
						p_dataMgr.SqlQueryReader(oConn, "SELECT * FROM project");

						p_dataMgr.m_DataReader.Read();
						// Adding trim function to remove extra spaces
						m_strNewProjectId = p_dataMgr.m_DataReader["proj_id"].ToString().Trim();
						m_strNewName = p_dataMgr.m_DataReader["created_by"].ToString().Trim();
						m_strNewDate = p_dataMgr.m_DataReader["created_date"].ToString();
						m_strNewOrganization = p_dataMgr.m_DataReader["organization"].ToString().Trim();
						m_strNewDescription = p_dataMgr.m_DataReader["description"].ToString().Trim();
						m_strOldProjectDirectory = p_dataMgr.m_DataReader["project_directory"].ToString().Trim();
						int intLastSlash = m_strOldProjectDirectory.LastIndexOf('\\');
						if (intLastSlash > 0)
						{
							m_strNewRootDirectory = m_strOldProjectDirectory.Substring(0, intLastSlash);
						}
						if (m_strNewProjectId != m_strOldProjectDirectory.Substring(intLastSlash + 1))
						{
							m_strNewProjectId = m_strOldProjectDirectory.Substring(intLastSlash + 1);

						}

						if (bAppVerColumnExist)
						{
							if (p_dataMgr.m_DataReader["application_version"] != System.DBNull.Value)
								this.m_strNewProjectVersion = p_dataMgr.m_DataReader["application_version"].ToString().Trim();
							else
								this.m_strNewProjectVersion = "";
						}
						else
						{
							this.m_strNewProjectVersion = "";
						}
					}
					catch (Exception caught)
					{
						MessageBox.Show(caught.Message);
					}
					p_dataMgr.m_DataReader.Close();
				}
				else
				{
					this.m_intError = p_dataMgr.m_intError;
					this.m_strError = p_dataMgr.m_strError;
					if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
						frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.OpenProjectTable: !!Failed to open project file!! Error=" + m_strError + "\r\n");
				}
			}
			p_dataMgr = null;

			if (this.m_strAction == "VIEW")
			{
				this.grpboxDescription.Enabled = false;
				this.grpboxProjectDirectory.Enabled = false;
				this.grpboxProjectId.Enabled = false;
				this.grpboxOrganization.Enabled = false;
				this.grpboxCreated.Enabled = false;
				this.btnEdit.Enabled = true;
				this.btnCancel.Enabled = false;
				this.btnSave.Enabled = false;

			}
			frmMain.g_sbpInfo.Text = "Ready";
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
            this.grpboxOrganization = new System.Windows.Forms.GroupBox();
            this.txtOrganization = new System.Windows.Forms.TextBox();
            this.grpboxProjectId = new System.Windows.Forms.GroupBox();
            this.txtProjectId = new System.Windows.Forms.TextBox();
            this.grpboxCreated = new System.Windows.Forms.GroupBox();
            this.txtDate = new System.Windows.Forms.TextBox();
            this.lblDate = new System.Windows.Forms.Label();
            this.txtName = new System.Windows.Forms.TextBox();
            this.lblName = new System.Windows.Forms.Label();
            this.grpboxProjectDirectory = new System.Windows.Forms.GroupBox();
            this.btnRootDirectoryHelp = new System.Windows.Forms.Button();
            this.btnRootDirectory = new System.Windows.Forms.Button();
            this.lblRootDirectory = new System.Windows.Forms.Label();
            this.txtRootDirectory = new System.Windows.Forms.TextBox();
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
            this.grpboxOrganization.SuspendLayout();
            this.grpboxProjectId.SuspendLayout();
            this.grpboxCreated.SuspendLayout();
            this.grpboxProjectDirectory.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(672, 440);
            this.panel1.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.grpboxDescription);
            this.groupBox1.Controls.Add(this.grpboxOrganization);
            this.groupBox1.Controls.Add(this.grpboxProjectId);
            this.groupBox1.Controls.Add(this.grpboxCreated);
            this.groupBox1.Controls.Add(this.grpboxProjectDirectory);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.btnEdit);
            this.groupBox1.Controls.Add(this.btnSave);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(672, 440);
            this.groupBox1.TabIndex = 29;
            this.groupBox1.TabStop = false;
            // 
            // grpboxDescription
            // 
            this.grpboxDescription.Controls.Add(this.txtDescription);
            this.grpboxDescription.Enabled = false;
            this.grpboxDescription.Location = new System.Drawing.Point(8, 209);
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
            // grpboxOrganization
            // 
            this.grpboxOrganization.Controls.Add(this.txtOrganization);
            this.grpboxOrganization.Enabled = false;
            this.grpboxOrganization.Location = new System.Drawing.Point(8, 157);
            this.grpboxOrganization.Name = "grpboxOrganization";
            this.grpboxOrganization.Size = new System.Drawing.Size(640, 48);
            this.grpboxOrganization.TabIndex = 27;
            this.grpboxOrganization.TabStop = false;
            this.grpboxOrganization.Text = "Organization";
            // 
            // txtOrganization
            // 
            this.txtOrganization.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOrganization.Location = new System.Drawing.Point(8, 14);
            this.txtOrganization.MaxLength = 100;
            this.txtOrganization.Name = "txtOrganization";
            this.txtOrganization.Size = new System.Drawing.Size(624, 23);
            this.txtOrganization.TabIndex = 0;
            // 
            // grpboxProjectId
            // 
            this.grpboxProjectId.Controls.Add(this.txtProjectId);
            this.grpboxProjectId.Enabled = false;
            this.grpboxProjectId.Location = new System.Drawing.Point(8, 103);
            this.grpboxProjectId.Name = "grpboxProjectId";
            this.grpboxProjectId.Size = new System.Drawing.Size(184, 48);
            this.grpboxProjectId.TabIndex = 0;
            this.grpboxProjectId.TabStop = false;
            this.grpboxProjectId.Text = "Project Workspace";
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
            this.grpboxCreated.Location = new System.Drawing.Point(200, 103);
            this.grpboxCreated.Name = "grpboxCreated";
            this.grpboxCreated.Size = new System.Drawing.Size(450, 48);
            this.grpboxCreated.TabIndex = 1;
            this.grpboxCreated.TabStop = false;
            this.grpboxCreated.Text = "Created";
            // 
            // txtDate
            // 
            this.txtDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDate.Location = new System.Drawing.Point(336, 16);
            this.txtDate.MaxLength = 8;
            this.txtDate.Name = "txtDate";
            this.txtDate.Size = new System.Drawing.Size(144, 23);
            this.txtDate.TabIndex = 3;
            // 
            // lblDate
            // 
            this.lblDate.Location = new System.Drawing.Point(264, 21);
            this.lblDate.Name = "lblDate";
            this.lblDate.Size = new System.Drawing.Size(72, 16);
            this.lblDate.TabIndex = 2;
            this.lblDate.Text = "Date Created";
            // 
            // txtName
            // 
            this.txtName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtName.Location = new System.Drawing.Point(82, 15);
            this.txtName.MaxLength = 30;
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(182, 23);
            this.txtName.TabIndex = 1;
            // 
            // lblName
            // 
            this.lblName.Location = new System.Drawing.Point(7, 21);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(80, 16);
            this.lblName.TabIndex = 0;
            this.lblName.Text = "Analyst Name";
            // 
            // grpboxProjectDirectory
            // 
            this.grpboxProjectDirectory.Controls.Add(this.btnRootDirectoryHelp);
            this.grpboxProjectDirectory.Controls.Add(this.btnRootDirectory);
            this.grpboxProjectDirectory.Controls.Add(this.lblRootDirectory);
            this.grpboxProjectDirectory.Controls.Add(this.txtRootDirectory);
            this.grpboxProjectDirectory.Location = new System.Drawing.Point(8, 42);
            this.grpboxProjectDirectory.Name = "grpboxProjectDirectory";
            this.grpboxProjectDirectory.Size = new System.Drawing.Size(640, 60);
            this.grpboxProjectDirectory.TabIndex = 29;
            this.grpboxProjectDirectory.TabStop = false;
            this.grpboxProjectDirectory.Text = "Project Directory";
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
            this.btnRootDirectory.Location = new System.Drawing.Point(561, 19);
            this.btnRootDirectory.Name = "btnRootDirectory";
            this.btnRootDirectory.Size = new System.Drawing.Size(32, 32);
            this.btnRootDirectory.TabIndex = 2;
            this.btnRootDirectory.Click += new System.EventHandler(this.btnRootDirectory_Click);
            // 
            // lblRootDirectory
            // 
            this.lblRootDirectory.Location = new System.Drawing.Point(16, 19);
            this.lblRootDirectory.Name = "lblRootDirectory";
            this.lblRootDirectory.Size = new System.Drawing.Size(88, 30);
            this.lblRootDirectory.TabIndex = 0;
            this.lblRootDirectory.Text = "Workspace Root Directory";
            // 
            // txtRootDirectory
            // 
            //this.txtRootDirectory.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtRootDirectory.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRootDirectory.Location = new System.Drawing.Point(113, 24);
            this.txtRootDirectory.Name = "txtRootDirectory";
            this.txtRootDirectory.Size = new System.Drawing.Size(416, 23);
            this.txtRootDirectory.TabIndex = 0;
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
            this.btnEdit.Location = new System.Drawing.Point(208, 380);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(96, 32);
            this.btnEdit.TabIndex = 7;
            this.btnEdit.Text = "Edit";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnSave
            // 
            this.btnSave.Enabled = false;
            this.btnSave.Location = new System.Drawing.Point(304, 380);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(96, 32);
            this.btnSave.TabIndex = 5;
            this.btnSave.Text = "Save";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(400, 380);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(96, 32);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Location = new System.Drawing.Point(560, 420);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 8;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(112, 380);
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
            this.Size = new System.Drawing.Size(672, 440);
            this.Resize += new System.EventHandler(this.uc_project_Resize);
            this.panel1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.grpboxDescription.ResumeLayout(false);
            this.grpboxDescription.PerformLayout();
            this.grpboxOrganization.ResumeLayout(false);
            this.grpboxOrganization.PerformLayout();
            this.grpboxProjectId.ResumeLayout(false);
            this.grpboxProjectId.PerformLayout();
            this.grpboxCreated.ResumeLayout(false);
            this.grpboxCreated.PerformLayout();
            this.grpboxProjectDirectory.ResumeLayout(false);
            this.grpboxProjectDirectory.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void btnEdit_Click(object sender, System.EventArgs e)
		{
            this.grpboxProjectDirectory.Enabled = true;
			this.txtProjectId.Enabled = false;
            this.txtDate.Enabled=false;

			this.grpboxCreated.Enabled=true;
			this.grpboxOrganization.Enabled=true;
			this.grpboxDescription.Enabled=true;
			this.btnSave.Enabled=true;
			this.btnCancel.Enabled=true;

			this.btnEdit.Enabled=false;
            this.grpboxProjectId.Enabled = true;
			

		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.grpboxProjectId.Enabled=false;
			this.grpboxProjectDirectory.Enabled=false;
			this.grpboxCreated.Enabled=false;
			this.grpboxOrganization.Enabled=false;
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
				this.OpenProjectTableNew(this.m_strProjectDirectory,this.m_strProjectFile);
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
			string strOrg = "";

			//validate the input
			//project id
			if (this.txtProjectId.Text.Length == 0) 
			{
				MessageBox.Show("Enter A Project Id ");
				this.txtProjectId.Focus();
				return;
			}
			string[] arrForbiddenChars = { "<", ">", ":", "\"", "/", "\\", "|", "?", "*" };
			bool bNameError = false;
			foreach (string character in arrForbiddenChars)
            {
				if (this.txtProjectId.Text.Contains(character))
                {
					bNameError = true;
					break;
                }
            }
			if (bNameError)
            {
				MessageBox.Show("Project Id cannot contain any of the following characters: " +
					"< > : \" / \\ | ? *");
				this.txtProjectId.Focus();
				return;
            }
			string[] arrReservedNames = {"CON", "PRN", "AUX", "NUL", "COM0", "COM1",
			"COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
			"LPT0", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"};
			foreach (string name in arrReservedNames)
            {
				if (this.txtProjectId.Text.ToUpper() == name)
                {
					bNameError = true;
					break;
                }
            }
			if (bNameError)
            {
				MessageBox.Show("Project Id cannot be any of the following names: " +
					"CON, PRN, AUX, NUL, COM0-COM9, LPT0-LPT9");
				this.txtProjectId.Focus();
				return;
            }
			if (this.txtProjectId.Text[0] == ' ' || this.txtProjectId.Text[0] == '.' ||
				this.txtProjectId.Text[this.txtProjectId.Text.Length-1] == ' ' || this.txtProjectId.Text[this.txtProjectId.Text.Length - 1] == '.')
            {
				MessageBox.Show("Project Id cannot contain leading or trailing spaces or periods");
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

			this.m_strProjectDirectory = this.txtRootDirectory.Text.Trim() + "\\" + this.txtProjectId.Text.Trim();

			try
			{
				if (this.m_strAction == "NEW") 
				{
				
					if (System.IO.Directory.Exists(this.m_strProjectDirectory))
					{
						for (x=1;x<=1000;x++)
						{
							strFullPath = this.m_strProjectDirectory + x.ToString().Trim();
							if (!System.IO.Directory.Exists(strFullPath))
							{
								this.m_strProjectDirectory = strFullPath.Trim();
								this.txtProjectId.Text = this.txtProjectId.Text.Trim() + x.ToString().Trim();
								break;
							}
						}
					}
				}
				strFullPath = this.m_strProjectDirectory + "\\db";
				if (!System.IO.Directory.Exists(strFullPath))
					System.IO.Directory.CreateDirectory(strFullPath);

				    
				strFullPath = this.m_strProjectDirectory + "\\optimizer\\db";
				if (!System.IO.Directory.Exists(strFullPath))
					System.IO.Directory.CreateDirectory(strFullPath);

				strFullPath = this.m_strProjectDirectory + "\\gis\\db";
				if (!System.IO.Directory.Exists(strFullPath))
					System.IO.Directory.CreateDirectory(strFullPath);

				strFullPath = this.m_strProjectDirectory + "\\fvs\\db";
				if (!System.IO.Directory.Exists(strFullPath))
					System.IO.Directory.CreateDirectory(strFullPath);

				strFullPath = this.m_strProjectDirectory + "\\fvs\\data";
				if (!System.IO.Directory.Exists(strFullPath))
					System.IO.Directory.CreateDirectory(strFullPath);

				strFullPath = this.m_strProjectDirectory + "\\processor\\db";
				if (!System.IO.Directory.Exists(strFullPath))
					System.IO.Directory.CreateDirectory(strFullPath);
		
				
			}
			catch 
			{
				MessageBox.Show("Error Creating Project Folder");
				
				return;
			}

		
			//check if new project
            DataMgr p_dataMgr = new DataMgr();
			if (this.m_strAction == "NEW") 
			{
				//new project
				//copy default project file to new project directory
				

				strDestFile = this.m_strProjectDirectory + "\\db\\project.db";
				
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
				p_dataMgr.CreateDbFile(strDestFile);
				strConn = p_dataMgr.GetConnectionString(strDestFile);
				using (var oConn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
					oConn.Open();

					//datasource table
					frmMain.g_oTables.m_oProject.CreateDatasourceTable(p_dataMgr, oConn, Tables.Project.DefaultProjectDatasourceTableName);
					//project table
					frmMain.g_oTables.m_oProject.CreateProjectTable(p_dataMgr, oConn, Tables.Project.DefaultProjectTableName);
                }

				//
				//travel times file and tables
				//
				p_frmTherm.lblMsg.Text = strDestFile;
                strDestFile = this.m_strProjectDirectory + "\\" + Tables.TravelTime.DefaultTravelTimePathAndDbFile;
                p_frmTherm.Increment(2);
                p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
                p_dataMgr.CreateDbFile(strDestFile);
				strConn = p_dataMgr.GetConnectionString(strDestFile);
                using (var oConn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    oConn.Open();
                    //processing site table
                    frmMain.g_oTables.m_oTravelTime.CreateProcessingSiteTable(p_dataMgr, oConn, Tables.TravelTime.DefaultProcessingSiteTableName);
                    //travel time table
                    frmMain.g_oTables.m_oTravelTime.CreateTravelTimeTable(p_dataMgr, oConn, Tables.TravelTime.DefaultTravelTimeTableName);
                }
                strDestFile = this.m_strProjectDirectory + "\\" + Tables.TravelTime.DefaultGisAuditPathAndDbFile;
                p_dataMgr.CreateDbFile(strDestFile);

                //
                //master file
                //
                strDestFile = $@"{this.m_strProjectDirectory}\{frmMain.g_oTables.m_oFIAPlot.DefaultPopTableDbFile}";
                strConn = p_dataMgr.GetConnectionString(strDestFile);
                p_frmTherm.Increment(3);
				p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
                p_dataMgr.CreateDbFile(strDestFile);
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    con.Open();
                    //plot table
                    frmMain.g_oTables.m_oFIAPlot.CreatePlotTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName);
                    //cond table
                    frmMain.g_oTables.m_oFIAPlot.CreateConditionTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName);
                    //site tree table
                    frmMain.g_oTables.m_oFIAPlot.CreateSiteTreeTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName);
                    //tree table
                    frmMain.g_oTables.m_oFIAPlot.CreateTreeTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName);
                    //biosum pop stratum adjustment factors table
                    frmMain.g_oTables.m_oFIAPlot.CreateBiosumPopStratumAdjustmentFactorsTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName);
                    //pop estimation unit table
                    frmMain.g_oTables.m_oFIAPlot.CreatePopEstnUnitTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName);
                    //pop eval table
                    frmMain.g_oTables.m_oFIAPlot.CreatePopEvalTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName);
                    //pop plot stratum assignment table
                    frmMain.g_oTables.m_oFIAPlot.CreatePopPlotStratumAssgnTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName);
                    //pop stratum table
                    frmMain.g_oTables.m_oFIAPlot.CreatePopStratumTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName);
                }

                //
                //master_aux.db
                //
                strDestFile = this.m_strProjectDirectory + "\\db\\master_aux.db";
                p_frmTherm.Increment(4);
                p_frmTherm.lblMsg.Text = strDestFile;
                p_frmTherm.lblMsg.Refresh();
                p_dataMgr.CreateDbFile(strDestFile);
                strConn = p_dataMgr.GetConnectionString(strDestFile);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    frmMain.g_oTables.m_oFIAPlot.CreateDWMCoarseWoodyDebrisTable(p_dataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName);
                    frmMain.g_oTables.m_oFIAPlot.CreateDWMFineWoodyDebrisTable(p_dataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName);
                    frmMain.g_oTables.m_oFIAPlot.CreateDWMDuffLitterFuelTable(p_dataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName);
                    frmMain.g_oTables.m_oFIAPlot.CreateDWMTransectSegmentTable(p_dataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName);
                }

                //
                //master file
                //
				p_frmTherm.Increment(5);
				p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
                strDestFile = this.m_strProjectDirectory + "\\" + Tables.FVS.DefaultRxPackageDbFile;
                strConn = p_dataMgr.GetConnectionString(strDestFile);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    //rx table
                    frmMain.g_oTables.m_oFvs.CreateRxTable(p_dataMgr, conn, Tables.FVS.DefaultRxTableName);
                    //rx harvest cost column table
                    frmMain.g_oTables.m_oFvs.CreateRxHarvestCostColumnTable(p_dataMgr, conn, Tables.FVS.DefaultRxHarvestCostColumnsTableName);
                    //rx packages table
                    frmMain.g_oTables.m_oFvs.CreateRxPackageTable(p_dataMgr, conn, Tables.FVS.DefaultRxPackageTableName);

                }
                //fvs output pre-post seqnum processing
                uc_fvs_output_prepost_seqnum.InitializePrePostSeqNumTables();

                //
                //prepopulated ref master file
                //
                //copy default ref_master SQLitedatabase to the new project directory
                strSourceFile = $@"{this.m_oEnv.strAppDir}\{Tables.Reference.DefaultRefMasterDbFile}";
				strDestFile = $@"{this.m_strProjectDirectory}\{Tables.Reference.DefaultRefMasterDbFile}";
				p_frmTherm.Increment(6);
				p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
				System.IO.File.Copy(strSourceFile, strDestFile, true);

				//
				//prepopulated OPCOST ref file file
				//
				//copy default master database to the new project directory
				strSourceFile = this.m_oEnv.strAppDir + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile;
                strDestFile = this.m_strProjectDirectory + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile;
                System.IO.File.Copy(strSourceFile, strDestFile, true);
                //
                //prepopulated weighted variable optimizer_definitions.db file
                //
				//copy default optimizer_definitions.db to new project directory
				strSourceFile = this.m_oEnv.strAppDir + "\\db\\optimizer_definitions.db";
				strDestFile = this.m_strProjectDirectory + "\\" + Tables.OptimizerDefinitions.DefaultDbFile;
				p_frmTherm.Increment(8);
				p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
				System.IO.File.Copy(strSourceFile, strDestFile, true);

                //
                //optimizer scenario rule definitions
                //
				p_frmTherm.Increment(9);
                p_frmTherm.lblMsg.Text = this.m_strProjectDirectory + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;
				p_frmTherm.lblMsg.Refresh();
				CreateOptimizerScenarioRuleDefinitionDbAndTables(this.m_strProjectDirectory + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile);
				//
				//processor scenario rule definitions
				//
				p_frmTherm.lblMsg.Text = strDestFile;
				p_frmTherm.lblMsg.Refresh();
                CreateProcessorScenarioRuleDefinitionDbAndTables($@"{this.m_strProjectDirectory}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultDbFile}");

                p_frmTherm.Increment(10);
				p_frmTherm.lblMsg.Refresh();
				

				p_frmTherm.Increment(13);
				strSourceFile = this.m_strProjectDirectory + "\\db\\project.db";
				strConn = p_dataMgr.GetConnectionString(strSourceFile);
				p_frmTherm.Close();
				p_frmTherm.Dispose();
				p_frmTherm = null;

				using (var oConn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
					oConn.Open();

					if (this.txtDescription.Text.Trim().Length > 0)
                    {
						strDesc = p_dataMgr.FixString(this.txtDescription.Text.Trim(), "'", "''");
                    }
					if (this.txtOrganization.Text.Trim().Length > 0)
                    {
						strOrg = p_dataMgr.FixString(this.txtOrganization.Text.Trim(), "'", "''");
                    }
					p_dataMgr.m_strSQL = "INSERT INTO project (" +
						"proj_id, created_by, created_date, organization, description, project_directory, application_version) " +
						"VALUES (" +
						"'" + this.txtProjectId.Text.Trim() + "', " +
						"'" + this.txtName.Text.Trim() + "', " +
						"'" + this.txtDate.Text + "', " +
						"'" + strOrg + "', " +
						"'" + strDesc + "', " +
						"'" + this.m_strProjectDirectory + "', " +
						"'" + frmMain.g_strAppVer + "')";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('Plot'," +
						"'" + this.m_strProjectDirectory + "\\db'," +
						"'master.db'," +
						"'plot');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('Condition'," +
						"'" + this.m_strProjectDirectory + "\\db'," +
						"'master.db'," +
						"'cond');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('Tree'," +
						"'" + this.m_strProjectDirectory + "\\db'," +
						"'master.db'," +
						"'tree');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('Treatment Prescriptions'," +
						"'" + this.m_strProjectDirectory + "\\db'," +
						"'master.db'," +
						"'rx');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('Treatment Prescriptions Harvest Cost Columns'," +
						"'" + this.m_strProjectDirectory + "\\db'," +
						"'master.db'," +
						"'" + Tables.FVS.DefaultRxHarvestCostColumnsTableName + "');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('Treatment Packages'," +
						"'" + this.m_strProjectDirectory + "\\db'," +
						"'master.db'," +
						"'rxpackage');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('FVS PRE-POST SeqNum Definitions'," +
						"'" + this.m_strProjectDirectory + "\\db'," +
						"'master.db'," +
						"'" + Tables.FVS.DefaultFVSPrePostSeqNumTable + "');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('FVS PRE-POST SeqNum Treatment Package Assign'," +
						"'" + this.m_strProjectDirectory + "\\db'," +
						"'master.db'," +
						"'" + Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable + "');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('Travel Times'," +
						"'" + this.m_strProjectDirectory + "\\gis\\db'," +
						"'" + Tables.TravelTime.DefaultTravelTimeDbFile + "'," +
						"'travel_time');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('Processing Sites'," +
						"'" + this.m_strProjectDirectory + "\\gis\\db'," +
						"'" + Tables.TravelTime.DefaultTravelTimeDbFile + "'," +
						"'processing_site');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('FIA Tree Macro Plot Breakpoint Diameter'," +
						"'@@appdata@@\\fiabiosum'," +
						"'biosum_ref.db'," +
						"'TreeMacroPlotBreakPointDia');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('" + Datasource.TableTypes.HarvestMethods + "'," +
						"'@@appdata@@\\fiabiosum'," +
						"'" + Tables.Reference.DefaultBiosumReferenceFile + "'," +
						"'" + Tables.Reference.DefaultHarvestMethodsTableName + "');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('" + Datasource.TableTypes.PopStratumAdjFactors + "'," +
						"'" + this.m_strProjectDirectory + "\\db'," +
						"'master.db'," +
						"'" + frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName + "');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
						"('Site Tree'," +
						"'" + this.m_strProjectDirectory + "\\db'," +
						"'master.db'," +
						"'sitetree');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					p_dataMgr.m_strSQL = "INSERT INTO datasource (table_type, path, file, table_name) VALUES " +
							 "('" + Datasource.TableTypes.FiaTreeSpeciesReference + "'," +
							 "'@@AppData@@" + frmMain.g_strBiosumDataDir + "'," +
							 "'" + Tables.Reference.DefaultBiosumReferenceFile + "'," +
							 "'" + Tables.ProcessorScenarioRun.DefaultFiaTreeSpeciesRefTableName + "');";
					p_dataMgr.SqlNonQuery(oConn, p_dataMgr.m_strSQL);

					frmMain.g_oGeneralMacroSubstitutionVariable_Collection.Item(frmMain.PROJDIR).VariableSubstitutionString = this.m_strProjectDirectory;
				}

				frmMain.g_oUtils.WriteText(this.m_strProjectDirectory + "\\application.version",frmMain.g_strAppVer);
				



				//make the new project the current project
				this.OpenProjectTableNew(this.m_strProjectDirectory, "project.db");

				if (this.m_intError == 0)
				{
					this.lblTitle.Text = "Project Properties";
				}
				((frmMain)this.ParentForm.ParentForm).OpenProject(this.m_strProjectDirectory, "project.db");

			}
			else 
			{
				strFullPath = this.m_strProjectDirectory + "\\db\\" + this.m_strProjectFile;
				strConn = p_dataMgr.GetConnectionString(strFullPath);
				if (this.txtDescription.Text.Trim().Length > 0)
                {
					strDesc = p_dataMgr.FixString(this.txtDescription.Text.Trim(), "'", "''");
                }
				if (this.txtOrganization.Text.Trim().Length > 0)
				{
					strOrg = p_dataMgr.FixString(this.txtOrganization.Text.Trim(), "'", "''");
				}
				p_dataMgr.m_strSQL = "UPDATE project " +
					"SET created_by = '" + this.txtName.Text + "', " +
					"proj_id = '" + this.txtProjectId.Text.Trim() + "', " +
					"organization = '" + strOrg + "', " +
					"description = '" + strDesc + "', " +
					"project_directory = '" + this.txtRootDirectory.Text + "'";
				p_dataMgr.SqlNonQuery(strConn, p_dataMgr.m_strSQL);
				this.m_strProjectId = this.txtProjectId.Text.Trim();
				((frmMain)this.ParentForm.ParentForm).Text = "Fia Biosum Manager (" + this.m_strProjectId + ")";

			}
            p_dataMgr = null;
			this.btnSave.Enabled=false;
			this.btnCancel.Enabled=false;
			this.btnEdit.Enabled=true;
			this.grpboxProjectDirectory.Enabled=false;
			this.grpboxCreated.Enabled=false;
			this.grpboxOrganization.Enabled=false;
			this.grpboxDescription.Enabled=false;
			this.grpboxProjectId.Enabled=false;
			this.m_strProjectFile = "project.db";
			this.m_strProjectId = this.txtProjectId.Text.Trim();
			this.m_strAction="";

		}
		public void New_Project()
		{
			this.m_strAction="NEW";
		    
            this.txtOrganization.Text = "";
			this.txtDate.Text = System.DateTime.Now.ToString();
			this.txtDescription.Text = "";
			this.txtName.Text = "";
			this.txtProjectId.Text = "";
			
			this.txtProjectId.Enabled=true;
			this.txtName.Enabled=true;
			this.txtOrganization.Enabled=true;
			this.txtDescription.Enabled=true;
			this.txtDate.Enabled=false;
			this.btnRootDirectory.Enabled=true;
			this.btnEdit.Enabled=false;
			this.btnSave.Enabled=true;
			this.btnCancel.Enabled=true;
			this.grpboxProjectDirectory.Enabled=true;
			this.grpboxCreated.Enabled=true;
			this.grpboxOrganization.Enabled=true;
			this.grpboxDescription.Enabled=true;
			this.grpboxProjectId.Enabled=true;
			this.Parent.Visible = true;
		}
		private void grpboxProjectDirectory_Enter(object sender, System.EventArgs e)
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
    				//this.txtRootDirectory.Text = strTemp + "\\" + this.txtProjectId.Text.Trim();
					this.txtRootDirectory.Text = strTemp;
				}
			}
		}
		

		public void CreateOptimizerScenarioRuleDefinitionDbAndTables(string p_strPathAndFile)
        {
			DataMgr dataMgr = new DataMgr();

			dataMgr.CreateDbFile(p_strPathAndFile);
			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dataMgr.GetConnectionString(p_strPathAndFile)))
            {
				conn.Open();
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioCostsTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName);
				frmMain.g_oTables.m_oScenario.CreateScenarioDatasourceTable(dataMgr, conn, Tables.Scenario.DefaultScenarioDatasourceTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioHarvestCostColumnsTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioHarvestCostColumnsTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioLandOwnerGroupsTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioPlotFilterTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioPSitesTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioLastTieBreakRankTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLastTieBreakRankTableName);
				frmMain.g_oTables.m_oScenario.CreateScenarioTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariableOverallEffectiveTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesOptimizationTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesTieBreakerTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioCondFilterMiscTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioCondFilterTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterTableName);
				frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioProcessorScenarioSelectTable(dataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName);
				conn.Close();
            }

			//create empty prepost_fvs_weighted.db
			string strDestFile = m_strProjectDirectory.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile;
			if (!System.IO.File.Exists(strDestFile))
            {
				dataMgr.CreateDbFile(strDestFile);
			}
		}
        public void CreateProcessorScenarioRuleDefinitionDbAndTables(string p_strPathAndFile)
        {
            DataMgr dataMgr = new DataMgr();

            dataMgr.CreateDbFile(p_strPathAndFile);
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dataMgr.GetConnectionString(p_strPathAndFile)))
            {
                conn.Open();
                frmMain.g_oTables.m_oScenario.CreateScenarioDatasourceTable(dataMgr, conn, Tables.Scenario.DefaultScenarioDatasourceTableName);
                frmMain.g_oTables.m_oScenario.CreateScenarioTable(dataMgr, conn, Tables.Scenario.DefaultScenarioTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioTreeSpeciesDollarValuesTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioHarvestMethodTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioCostRevenueEscalatorsTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultCostRevenueEscalatorsTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioHarvestCostColumnsTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestCostColumnsTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioAdditionalHarvestCostsTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioMoveInCostsTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioTreeDiamGroupsTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioTreeSpeciesGroupsListTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName);
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioTreeSpeciesGroupsTable(dataMgr, conn, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName);
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
			OpenFileDialog1.Title = "Open FIA Biosum Project Database File";
			OpenFileDialog1.Filter = "Database Files (*.MDB,*.MDE,*.ACCDB,*.DB,*.SQLITE,*SQLITE3) |*.mdb;*.mde;*.accdb;*.db;*.sqlite;*.sqlite3";
			
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

					if (this.m_strNewProjectFile.EndsWith(".mdb") || this.m_strNewProjectFile.EndsWith(".mde") || this.m_strNewProjectFile.EndsWith(".accdb"))
                    {
						// Warn if trying to open project.mdb with version 5.12.1
						string strProjVersionFile = m_strNewProjectDirectory + "\\application.version";
						string strProjVersion = "";
						if (System.IO.File.Exists(strProjVersionFile))
						{
							using (System.IO.StreamReader reader = new System.IO.StreamReader(strProjVersionFile))
							{
								string strTemp = reader.ReadLine();
								if (strTemp != null)
								{
									strProjVersion = strTemp.Trim();
								}
							}
						}
						if (strProjVersion == "5.12.1")
						{
							System.Text.StringBuilder sb = new System.Text.StringBuilder();
							sb.Append("WARNING: You are trying to open an Access project database instead of a SQLite project database! ");
							sb.Append("BioSum version v5.12.1 requires using a SQLite project.db.");
							System.Windows.Forms.DialogResult res =
								System.Windows.Forms.MessageBox.Show(sb.ToString(), "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK);
							this.m_intError = -1;
							return;
						}
						this.OpenProjectTable(this.m_strNewProjectDirectory, this.m_strNewProjectFile);
					}
                    else
                    {
						this.OpenProjectTableNew(this.m_strNewProjectDirectory, this.m_strNewProjectFile);
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
                    
                
				this.OpenProjectTableNew(this.m_strNewProjectDirectory, this.m_strNewProjectFile);
                
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
			this.grpboxProjectDirectory.Top = this.lblTitle.Top + this.lblTitle.Height + 2;
			this.grpboxProjectId.Top = this.grpboxProjectDirectory.Top + this.grpboxProjectDirectory.Height + 2;
			this.grpboxCreated.Top = this.grpboxProjectId.Top;
			this.grpboxOrganization.Top = this.grpboxProjectId.Top + this.grpboxProjectId.Height + 2;
			this.grpboxDescription.Top = this.grpboxOrganization.Top + this.grpboxOrganization.Height + 2;
			
			this.grpboxDescription.Left = 2;
			this.grpboxProjectDirectory.Left = 2;
			this.grpboxProjectId.Left = 2;
			this.grpboxOrganization.Left = 2;
			this.grpboxDescription.Width  = this.Width - 4;
			this.grpboxProjectDirectory.Width = this.grpboxDescription.Width;
			this.grpboxOrganization.Width = this.grpboxDescription.Width;
			this.grpboxCreated.Width = this.grpboxDescription.Width - this.grpboxProjectId.Left - this.grpboxProjectId.Width - 12;
			

			this.lblRootDirectory.Left = this.grpboxProjectDirectory.Left + 2;

			this.txtRootDirectory.Left = this.lblRootDirectory.Left + this.lblRootDirectory.Width + 2;

			this.btnRootDirectoryHelp.Left = this.grpboxProjectDirectory.Width - this.btnRootDirectoryHelp.Width - 2;

			this.btnRootDirectory.Left = this.btnRootDirectoryHelp.Left - this.btnRootDirectoryHelp.Width - 2;

			this.btnRootDirectoryHelp.Top = this.btnRootDirectory.Top;

			this.txtRootDirectory.Width = this.grpboxProjectDirectory.Width  - this.lblRootDirectory.Width  - this.btnRootDirectory.Width - this.btnRootDirectoryHelp.Width   - 15;
			
			
			this.txtDescription.Left = this.grpboxDescription.Left + 2;
			this.txtDescription.Width = this.grpboxDescription.Width - 8;
			this.txtOrganization.Left = this.grpboxOrganization.Left + 2;
			this.txtOrganization.Width = this.grpboxOrganization.Width - 8;
			

			this.btnCancel.Top = this.grpboxDescription.Top + this.grpboxDescription.Height + 5;
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
			strFullPath = strProjDir + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;
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
			strFullPath = $@"{strProjDir}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultDbFile}";
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


		public void SetProjectPathEnvironmentVariablesSqlite()
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

			frmMain.g_oGeneralMacroSubstitutionVariable_Collection.Item(frmMain.OLDPROJDIR).VariableSubstitutionString = this.m_strOldProjectDirectory.Trim();
            frmMain.g_oGeneralMacroSubstitutionVariable_Collection.Item(frmMain.PROJDIR).VariableSubstitutionString = this.m_strProjectDirectory.Trim();

            strProjDir = m_strProjectDirectory.Trim();
			strOldProjDir = m_strOldProjectDirectory.Trim();
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Replace old project directory (" + strOldProjDir + ") with new project directory (" + strProjDir + ")\r\n");
            

            /**********************************************
			 **instantiate the DataMgr class
			 **********************************************/
            DataMgr oDataMgr = new DataMgr();
            //
            //PROJECT DATA SOURCE
            //
			strFullPath = strProjDir + "\\db\\" + this.m_strProjectFile;
			strConn = oDataMgr.GetConnectionString(strFullPath);
			if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
				frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Open Connection to Project Dbfile " + strConn + ")\r\n");
			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
				conn.Open();

				strSQL = "UPDATE project SET proj_id = '" + this.txtProjectId.Text.Trim() + "'";
				if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
					frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Execute SQL \r\n" + strSQL + "\r\n");
				oDataMgr.SqlNonQuery(conn, strSQL);

				strSQL = "UPDATE project SET project_directory = '" + strProjDir + "' " +
					"WHERE proj_id = '" + this.txtProjectId.Text.Trim() + "'";
				if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
					frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Execute SQL \r\n" + strSQL + "\r\n");
				oDataMgr.SqlNonQuery(conn, strSQL);

				strSQL = "UPDATE datasource SET path = " +
					"REPLACE(TRIM(LOWER(path)), '" + strOldProjDir.Trim().ToLower() + "', " +
					"'" + strProjDir.Trim().ToLower() + "')";
				if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
					frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Execute SQL \r\n" + strSQL + "\r\n");
				oDataMgr.SqlNonQuery(conn, strSQL);
			}

			//
			//TREATMENT OPTIMIZER SCENARIO DATA SOURCE
			//
			strFullPath = strProjDir + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;
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
            strFullPath = $@"{strProjDir}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultDbFile}";
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



            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "uc_project.SetProjectPathEnvironmentVariables: Leaving \r\n");
			
 		}

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "DATABASE", "NEWPROJECT" });
            //m_oHelp.ShowPdfHelp(m_pdfFile, "1");
        }

		
	}
}



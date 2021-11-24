using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Text;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario.
	/// </summary>
	public class uc_scenario_open : System.Windows.Forms.UserControl
	{
		public System.Windows.Forms.Button btnOpen;
		public System.Windows.Forms.TextBox txtDescription;
		private System.Data.DataSet dataSet1;
		private System.Data.OleDb.OleDbCommand oleDbSelectCommand1;
		private System.Data.OleDb.OleDbCommand oleDbInsertCommand1;
		private System.Data.OleDb.OleDbCommand oleDbUpdateCommand1;
		private System.Data.OleDb.OleDbCommand oleDbDeleteCommand1;
		private System.Data.OleDb.OleDbDataAdapter oleDbDataAdapter1;
		private System.Data.OleDb.OleDbCommand oleDbCommand1;
		private System.Data.OleDb.OleDbConnection oleDbConnection1;
		private System.ComponentModel.Container components = null;
		public int intError;
		public System.Windows.Forms.Button btnCancel;
		public string strError;
		public System.Windows.Forms.ListBox lstScenario;
		public System.Windows.Forms.Label lblScenarioId;
		public System.Windows.Forms.Label lblScenarioDescription;
		public System.Windows.Forms.Label lblScenarioPath;
		public System.Windows.Forms.TextBox txtScenarioPath;
		public System.Windows.Forms.Label lblNewScenario;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Button btnClose;
		public System.Windows.Forms.Label lblTitle;
		public System.Windows.Forms.TextBox txtScenarioId;
		public int m_intFullHt=500;
		public int m_intFullWd=650;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario;
		private FIA_Biosum_Manager.frmProcessorScenario _frmProcessorScenario;
		private string _strScenarioType="optimizer";
		
		// public FIA_Biosum_Manager.frmScenario frmscenario1;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		

		public uc_scenario_open()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.lstScenario.Click += new System.EventHandler(this.lstScenario_Click);

			this.txtScenarioPath.Enabled=false;
			


			this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
			this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
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
            this.btnOpen = new System.Windows.Forms.Button();
            this.txtDescription = new System.Windows.Forms.TextBox();
            this.lstScenario = new System.Windows.Forms.ListBox();
            this.lblScenarioId = new System.Windows.Forms.Label();
            this.lblScenarioDescription = new System.Windows.Forms.Label();
            this.lblScenarioPath = new System.Windows.Forms.Label();
            this.txtScenarioPath = new System.Windows.Forms.TextBox();
            this.dataSet1 = new System.Data.DataSet();
            this.oleDbSelectCommand1 = new System.Data.OleDb.OleDbCommand();
            this.oleDbInsertCommand1 = new System.Data.OleDb.OleDbCommand();
            this.oleDbUpdateCommand1 = new System.Data.OleDb.OleDbCommand();
            this.oleDbDeleteCommand1 = new System.Data.OleDb.OleDbCommand();
            this.oleDbDataAdapter1 = new System.Data.OleDb.OleDbDataAdapter();
            this.oleDbCommand1 = new System.Data.OleDb.OleDbCommand();
            this.oleDbConnection1 = new System.Data.OleDb.OleDbConnection();
            this.lblNewScenario = new System.Windows.Forms.Label();
            this.txtScenarioId = new System.Windows.Forms.TextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOpen
            // 
            this.btnOpen.BackColor = System.Drawing.SystemColors.Control;
            this.btnOpen.Location = new System.Drawing.Point(224, 400);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(96, 32);
            this.btnOpen.TabIndex = 1;
            this.btnOpen.Text = "OK";
            this.btnOpen.UseVisualStyleBackColor = false;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // txtDescription
            // 
            this.txtDescription.Enabled = false;
            this.txtDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDescription.Location = new System.Drawing.Point(168, 193);
            this.txtDescription.Multiline = true;
            this.txtDescription.Name = "txtDescription";
            this.txtDescription.Size = new System.Drawing.Size(448, 152);
            this.txtDescription.TabIndex = 2;
            this.txtDescription.TextChanged += new System.EventHandler(this.txtDescription_TextChanged);
            this.txtDescription.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDescription_KeyPress);
            // 
            // lstScenario
            // 
            this.lstScenario.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstScenario.ItemHeight = 25;
            this.lstScenario.Location = new System.Drawing.Point(8, 74);
            this.lstScenario.Name = "lstScenario";
            this.lstScenario.Size = new System.Drawing.Size(144, 304);
            this.lstScenario.TabIndex = 3;
            this.lstScenario.SelectedIndexChanged += new System.EventHandler(this.lstScenario_SelectedIndexChanged);
            // 
            // lblScenarioId
            // 
            this.lblScenarioId.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblScenarioId.Location = new System.Drawing.Point(16, 49);
            this.lblScenarioId.Name = "lblScenarioId";
            this.lblScenarioId.Size = new System.Drawing.Size(120, 23);
            this.lblScenarioId.TabIndex = 4;
            this.lblScenarioId.Text = "Scenario List";
            // 
            // lblScenarioDescription
            // 
            this.lblScenarioDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblScenarioDescription.Location = new System.Drawing.Point(168, 166);
            this.lblScenarioDescription.Name = "lblScenarioDescription";
            this.lblScenarioDescription.Size = new System.Drawing.Size(160, 16);
            this.lblScenarioDescription.TabIndex = 5;
            this.lblScenarioDescription.Text = "Scenario Description";
            // 
            // lblScenarioPath
            // 
            this.lblScenarioPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblScenarioPath.Location = new System.Drawing.Point(168, 109);
            this.lblScenarioPath.Name = "lblScenarioPath";
            this.lblScenarioPath.Size = new System.Drawing.Size(136, 15);
            this.lblScenarioPath.TabIndex = 6;
            this.lblScenarioPath.Text = "Scenario Directory Path";
            // 
            // txtScenarioPath
            // 
            this.txtScenarioPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtScenarioPath.Location = new System.Drawing.Point(168, 129);
            this.txtScenarioPath.Name = "txtScenarioPath";
            this.txtScenarioPath.Size = new System.Drawing.Size(448, 30);
            this.txtScenarioPath.TabIndex = 7;
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "NewDataSet";
            this.dataSet1.Locale = new System.Globalization.CultureInfo("en-US");
            // 
            // oleDbDataAdapter1
            // 
            this.oleDbDataAdapter1.DeleteCommand = this.oleDbDeleteCommand1;
            this.oleDbDataAdapter1.InsertCommand = this.oleDbInsertCommand1;
            this.oleDbDataAdapter1.SelectCommand = this.oleDbSelectCommand1;
            this.oleDbDataAdapter1.UpdateCommand = this.oleDbUpdateCommand1;
            // 
            // lblNewScenario
            // 
            this.lblNewScenario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNewScenario.Location = new System.Drawing.Point(168, 53);
            this.lblNewScenario.Name = "lblNewScenario";
            this.lblNewScenario.Size = new System.Drawing.Size(128, 15);
            this.lblNewScenario.TabIndex = 9;
            this.lblNewScenario.Text = "Scenario Id";
            // 
            // txtScenarioId
            // 
            this.txtScenarioId.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtScenarioId.Location = new System.Drawing.Point(168, 75);
            this.txtScenarioId.MaxLength = 20;
            this.txtScenarioId.Name = "txtScenarioId";
            this.txtScenarioId.Size = new System.Drawing.Size(120, 30);
            this.txtScenarioId.TabIndex = 10;
            this.txtScenarioId.Leave += new System.EventHandler(this.txtScenarioId_Leave);
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.SystemColors.Control;
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(336, 400);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(96, 32);
            this.btnCancel.TabIndex = 14;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.lblNewScenario);
            this.groupBox1.Controls.Add(this.txtScenarioId);
            this.groupBox1.Controls.Add(this.lblScenarioPath);
            this.groupBox1.Controls.Add(this.txtScenarioPath);
            this.groupBox1.Controls.Add(this.lblScenarioDescription);
            this.groupBox1.Controls.Add(this.txtDescription);
            this.groupBox1.Controls.Add(this.btnOpen);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.lblScenarioId);
            this.groupBox1.Controls.Add(this.lstScenario);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(632, 480);
            this.groupBox1.TabIndex = 15;
            this.groupBox1.TabStop = false;
            this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(626, 32);
            this.lblTitle.TabIndex = 25;
            this.lblTitle.Text = "Open Scenario";
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.BackColor = System.Drawing.SystemColors.Control;
            this.btnClose.Location = new System.Drawing.Point(528, 440);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 15;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // uc_scenario_open
            // 
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_scenario_open";
            this.Size = new System.Drawing.Size(632, 480);
            this.Load += new System.EventHandler(this.uc_scenario_Load);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.uc_scenario_MouseDown);
            this.Resize += new System.EventHandler(this.uc_scenario_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void uc_scenario_Load(object sender, System.EventArgs e)
		{
		
		}

		private void uc_scenario_Resize(object sender, System.EventArgs e)
		{
			resize_uc_scenario();


		}
		public void resize_uc_scenario()
		{
			try
			{
				this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
				this.txtDescription.Width = this.Width - (this.txtDescription.Left * 2);

				btnOpen.Top = this.lstScenario.Top + this.lstScenario.Height + 5;
				btnCancel.Top = btnOpen.Top;
				this.btnOpen.Left = (int) (this.groupBox1.Width * .50) - (int) (this.btnOpen.Width / 2);

				this.btnCancel.Left = this.btnOpen.Left + this.btnOpen.Width;



			}
			catch
			{
			}
		}

		
		
		public void populate_scenario_listbox()
		{
			string strScenarioId="";
			string strDescription="";
			//string strScenarioMDBFile="";
			string strScenarioPath="";
	          
			System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
			string strProjDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();

			string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\db";
			string strFile = "scenario_" + ScenarioType + "_rule_definitions.mdb";
			StringBuilder strFullPath = new StringBuilder(strScenarioDir);
			strFullPath.Append("\\");
			strFullPath.Append(strFile);
			ado_data_access p_ado = new ado_data_access();
			string strConn=p_ado.getMDBConnString(strFullPath.ToString(),"admin","");

			p_ado.SqlQueryReader(strConn,"select * from scenario");
			if (p_ado.m_intError==0)
			{
				try
				{
					this.lstScenario.Items.Clear();
					while (p_ado.m_OleDbDataReader.Read())
					{
						strScenarioId = p_ado.m_OleDbDataReader["scenario_id"].ToString();
						//strScenarioMDBFile = p_ado.m_OleDbDataReader["file"].ToString();
						strDescription = p_ado.m_OleDbDataReader["description"].ToString();
						strScenarioPath = p_ado.m_OleDbDataReader["path"].ToString();
						this.lstScenario.Items.Add(p_ado.m_OleDbDataReader["scenario_id"].ToString());
					}
					this.lstScenario.SelectedIndex = this.lstScenario.Items.Count - 1;

					this.txtScenarioPath.Text = strScenarioPath;
					this.txtDescription.Text = strDescription;
                    


				}
				catch (Exception caught)
				{
					intError = -1;
					strError = caught.Message;
					MessageBox.Show(strError);
				}
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbDataReader = null;
				p_ado.m_OleDbCommand = null;
				p_ado.m_OleDbConnection.Close();
				p_ado.m_OleDbConnection = null;
			}
			p_ado = null;
		}
        public void populate_scenario_listbox_sqlite()
        {
            string strScenarioId = "";
            string strDescription = "";
            string strScenarioPath = "";
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();

            string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\" +
                    Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;

            string strConn = dataMgr.GetConnectionString(strScenarioDir);
            try
            {
                lstScenario.Items.Clear();
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    con.Open();
                    dataMgr.SqlQueryReader(con, "select * from scenario");
                    if (dataMgr.m_DataReader.HasRows)
                    {
                        while (dataMgr.m_DataReader.Read())
                        {
                            strScenarioId = dataMgr.m_DataReader["scenario_id"].ToString();
                            strDescription = dataMgr.m_DataReader["description"].ToString();
                            strScenarioPath = dataMgr.m_DataReader["path"].ToString();
                            this.lstScenario.Items.Add(strScenarioId);
                        }
                    }
                }
                this.lstScenario.SelectedIndex = this.lstScenario.Items.Count - 1;
                this.txtScenarioPath.Text = strScenarioPath;
                this.txtDescription.Text = strDescription;
            }
            catch (Exception caught)
            {
                intError = -1;
                strError = caught.Message;
                MessageBox.Show(strError, "FIA Biosum");
            }
        }
        private void btnFolder_Click(object sender, System.EventArgs e)
		{
            
			DialogResult result = ((frmMain)this.Parent.Parent).folderBrowserDialog1.ShowDialog();
			//the variable myPic contains the string of the full File Name,it includes the full path. 
			//string mymdb = OpenFileDialog1.FileName; 
			//MessageBox.Show(mymdb);
			if (result == DialogResult.OK) 
			{
				string strTemp = ((frmMain)this.Parent.Parent).folderBrowserDialog1.SelectedPath;
			
				if (strTemp.Length > 0) 
				{
					this.txtScenarioPath.Text = strTemp;
				}
			}
		}

		private void btnOpen_Click(object sender, System.EventArgs e)
		{
			//lets see if this scenario is already open
			try
			{
				utils oUtils = new utils();
				oUtils.m_intLevel=-1;
				if (this.ScenarioType.Trim().ToUpper()=="OPTIMIZER")
				{
					if (oUtils.FindWindowLike(frmMain.g_oFrmMain.Handle, "Treatment Optimizer: Optimization Scenario (" + this.lstScenario.SelectedItem.ToString().Trim() + ")","*",true,false) > 0)
					{
						MessageBox.Show("!!Scenario Already Open!!","Scenario Open",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
						return;
					}
					((frmOptimizerScenario)this.ParentForm).DialogResult=DialogResult.OK;
				}
				else
				{
					if (oUtils.FindWindowLike(frmMain.g_oFrmMain.Handle, "Processor: Scenario (" + this.lstScenario.SelectedItem.ToString().Trim() + ")","*",true,false) > 0)
					{
						MessageBox.Show("!!Scenario Already Open!!","Scenario Open",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
						return;
					}
					this.ReferenceProcessorScenarioForm.DialogResult=DialogResult.OK;
				}
				
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_scenario:btnOpen_Click  \n" + 
					"Err Msg - " + err.Message.ToString().Trim(),
					"Open Scenario",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
			
			}

			
			

		}
		public string strScenarioId
		{
			get {return this.txtScenarioId.Text;}
			set {this.txtScenarioId.Text = value;}
		}
		public string strScenarioPath
		{
			get {return this.txtScenarioPath.Text;}
			set {this.txtScenarioPath.Text = value;}
		}
		public string strScenarioDescription
		{
			get {return this.txtDescription.Text;}
			set {this.txtDescription.Text = value;}
		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
			this.SaveScenarioProperties();
		}
		public void SaveScenarioProperties()
		{
			string strTemp1;
			string strTemp2;
			string strSQL = "";
			bool bCore=false;
			string strDesc="";
			
			if (this.lstScenario.Visible == true) //new scenario
			{
				//validate the input
				//Optimization id
				if (this.txtScenarioId.Text.Length == 0 ) 
				{
					MessageBox.Show("Enter A Unique Optimization scenario Id");
					this.txtScenarioId.Focus();
					return;
				}

				//check for duplicate scenario id
				if (this.lstScenario.Items.Count > 0) 
				{
					strTemp2 = this.txtScenarioId.Text.Trim();
					for (int x = 0; x <= this.lstScenario.Items.Count - 1; x++)
					{
						strTemp1 = this.lstScenario.Items[x].ToString().Trim();
						if (strTemp1.ToUpper() == strTemp2.ToUpper()) 
						{
							MessageBox.Show("Cannot have a duplicate Optimization scenario id");
							this.txtScenarioId.Focus();
							return;
						}
						
					}
				}

				//make sure user entered scenario path
				if (this.txtScenarioPath.Text.Length > 0) 
				{
    				//create the scenario path if it does not exist and
					//copy the scenario_results.mdb to it
					try
					{
						if (!System.IO.Directory.Exists(this.txtScenarioPath.Text)) 
						{
							System.IO.Directory.CreateDirectory(this.txtScenarioPath.Text);
							System.IO.Directory.CreateDirectory(this.txtScenarioPath.Text.ToString() + "\\db");
							//copy default scenario_results database to the new project directory
							string strSourceFile = ((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\" + ScenarioType + "\\db\\scenario_results.mdb";
							string strDestFile = this.txtScenarioPath.Text + "\\db\\scenario_results.mdb";
							System.IO.File.Copy(strSourceFile, strDestFile,true);	
							
						}
					}
					catch 
					{
						MessageBox.Show("Error Creating Folder");
						return;
					}
				  
													   
				}
				else 
				{
					MessageBox.Show("Enter A Directory Location To Save Optimization scenario Files");
					this.txtScenarioPath.Focus();
					return;
				}	

				

				//copy the project data source values to the scenario data source
				ado_data_access p_ado = new ado_data_access();
				string strProjDBDir = ((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\db";
				string strProjFile = "project.mdb";
				StringBuilder strProjFullPath = new StringBuilder(strProjDBDir);
				strProjFullPath.Append("\\");
				strProjFullPath.Append(strProjFile);
				string strProjConn=p_ado.getMDBConnString(strProjFullPath.ToString(),"admin","");
				System.Data.OleDb.OleDbConnection p_OleDbProjConn = new System.Data.OleDb.OleDbConnection();
				p_ado.OpenConnection(strProjConn,ref p_OleDbProjConn);
                

				string strScenarioDBDir = ((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\" + ScenarioType + "\\db";
				string strScenarioFile = "scenario_" + ScenarioType + "_rule_definitions.mdb"; 
				StringBuilder strScenarioFullPath = new StringBuilder(strScenarioDBDir);
				strScenarioFullPath.Append("\\");
				strScenarioFullPath.Append(strScenarioFile);
				string strScenarioConn = p_ado.getMDBConnString(strScenarioFullPath.ToString(),"admin","");
				p_ado.OpenConnection(strScenarioConn);


				if (p_ado.m_intError==0)
				{
                   
					if (this.txtDescription.Text.Trim().Length > 0)
                        strDesc = p_ado.FixString(this.txtDescription.Text.Trim(),"'","''");
					strSQL = "INSERT INTO scenario (scenario_id,description,Path,File) VALUES " + "('" + this.txtScenarioId.Text.Trim() + "'," + 
						"'" + strDesc + "'," + 
						"'" + this.txtScenarioPath.Text.Trim() + "','scenario_" + ScenarioType + "_rule_definitions.mdb');";
						//"'" + this.txtScenarioMDBFile.Text.Trim() + "');";
					p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

					p_ado.SqlQueryReader(p_OleDbProjConn,"select * from datasource");
					if (p_ado.m_intError==0)
					{
						try
						{
							
							
							while (p_ado.m_OleDbDataReader.Read())
							{
								bCore=false;
								switch (p_ado.m_OleDbDataReader["table_type"].ToString().Trim().ToUpper())
								{
									case "PLOT":
										bCore=true;
										break;
									case "CONDITION":
										bCore = true;
										break;
									//case "FIRE AND FUEL EFFECTS":
									//	bCore = true;
									//	break;
                                    case "HARVEST COSTS":
										bCore = true;
										break;
									case "TREATMENT PRESCRIPTIONS":
										bCore = true;
										break;
									case "TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS":
										bCore = true;
										break;
									case "TRAVEL TIMES":
										bCore = true;
										break;
									case "PROCESSING SITES":
										bCore = true;
										break;
                                    case "TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES":
										bCore = true;
										break;
									default:
										break;
								}
								if (bCore == true)
								{
									strSQL = "INSERT INTO scenario_datasource (scenario_id,table_type,Path,file,table_name) VALUES " + "('" + this.txtScenarioId.Text.Trim() + "'," + 
										"'" + p_ado.m_OleDbDataReader["table_type"].ToString().Trim() + "'," + 
										"'" + p_ado.m_OleDbDataReader["path"].ToString().Trim() + "'," + 
										"'" + p_ado.m_OleDbDataReader["file"].ToString().Trim() + "'," +  
										"'" + p_ado.m_OleDbDataReader["table_name"].ToString().Trim() + "');";
									p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);
								}

							}
						}
						catch (Exception caught)
						{
							intError = -1;
							strError = caught.Message;
							MessageBox.Show(strError);
						}
						if (p_ado.m_intError==0)
						{
							if (ScenarioType.Trim().ToUpper() == "OPTIMIZER")
							{
								((frmOptimizerScenario)ParentForm).uc_datasource1.strScenarioId = this.txtScenarioId.Text.Trim();
								((frmOptimizerScenario)ParentForm).uc_datasource1.strDataSourceMDBFile = ((frmMain)ParentForm.ParentForm).frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\db\\scenario_" + ScenarioType + "_rule_definitions.mdb";
								((frmOptimizerScenario)ParentForm).uc_datasource1.strDataSourceTable = "scenario_datasource";
								((frmOptimizerScenario)ParentForm).uc_datasource1.strProjectDirectory = ((frmMain)ParentForm.ParentForm).frmProject.uc_project1.txtRootDirectory.Text.Trim();
							}
							else
							{
								this.ReferenceProcessorScenarioForm.uc_datasource1.strScenarioId = this.txtScenarioId.Text.Trim();
								this.ReferenceProcessorScenarioForm.uc_datasource1.strDataSourceMDBFile = ((frmMain)ParentForm.ParentForm).frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\db\\scenario_" + ScenarioType + "_rule_definitions.mdb";
								this.ReferenceProcessorScenarioForm.uc_datasource1.strDataSourceTable = "scenario_datasource";
								this.ReferenceProcessorScenarioForm.uc_datasource1.strProjectDirectory = ((frmMain)ParentForm.ParentForm).frmProject.uc_project1.txtRootDirectory.Text.Trim();
							}
						}
						p_ado.m_OleDbDataReader.Close();
						p_ado.m_OleDbDataReader = null;
						p_ado.m_OleDbCommand = null;
						p_OleDbProjConn.Close();
						p_OleDbProjConn = null;
					}
					if (ScenarioType.Trim().ToUpper() == "OPTIMIZER")
					{
						string strTemp=p_ado.FixString("SELECT @@PlotTable@@.* FROM @@PlotTable@@ ","'","''");
						strSQL = "INSERT INTO scenario_plot_filter (scenario_id,sql_command,current_yn) VALUES " + "('" + this.txtScenarioId.Text.Trim() + "'," + 
							"'" + strTemp + "'," + 
							"'Y');";
						p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);

						strTemp=p_ado.FixString("SELECT @@CondTable@@.* FROM @@CondTable@@","'","''");
						strSQL = "INSERT INTO scenario_cond_filter (scenario_id,sql_command,current_yn) VALUES " + "('" + this.txtScenarioId.Text.Trim() + "'," + 
							"'" + strTemp + "'," + 
							"'Y');";
						p_ado.SqlNonQuery(p_ado.m_OleDbConnection,strSQL);
					}
				}
				p_ado.m_OleDbConnection.Close();
				p_ado.m_OleDbConnection = null;
				p_ado = null;


				this.btnCancel.Enabled=false;
				this.btnOpen.Enabled=true;

				this.lstScenario.Enabled=true;
				this.txtScenarioId.Visible=false;
				this.lblNewScenario.Visible=false;
				this.txtScenarioPath.Enabled=false;
				this.lstScenario.Items.Add(this.txtScenarioId.Text);
				this.lstScenario.SelectedIndex = this.lstScenario.Items.Count -1 ;

			}
			else 
			{
				ado_data_access p_ado = new ado_data_access();

				System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
				string strProjDir = ((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory;
				string strScenarioDir = ((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\" + ScenarioType + "\\db";
				string strFile = "scenario_" + ScenarioType + "_rule_definitions.mdb"; 
				if (ScenarioType.Trim().ToUpper() == "OPTIMIZER")
				{
					((frmOptimizerScenario)ParentForm).uc_datasource1.strScenarioId = this.txtScenarioId.Text.Trim();
					((frmOptimizerScenario)ParentForm).uc_datasource1.strDataSourceMDBFile = strScenarioDir + "\\scenario_" + ScenarioType + "_rule_definitions.mdb";
					((frmOptimizerScenario)ParentForm).uc_datasource1.strDataSourceTable = "scenario_datasource";
					((frmOptimizerScenario)ParentForm).uc_datasource1.strProjectDirectory =  strProjDir;
				}
				else
				{
					this.ReferenceProcessorScenarioForm.uc_datasource1.strScenarioId = this.txtScenarioId.Text.Trim();
					this.ReferenceProcessorScenarioForm.uc_datasource1.strDataSourceMDBFile = strScenarioDir + "\\scenario_" + ScenarioType + "_rule_definitions.mdb";
					this.ReferenceProcessorScenarioForm.uc_datasource1.strDataSourceTable = "scenario_datasource";
					this.ReferenceProcessorScenarioForm.uc_datasource1.strProjectDirectory =  strProjDir;
				}
				StringBuilder strFullPath = new StringBuilder(strScenarioDir);
				strFullPath.Append("\\");
				strFullPath.Append(strFile);
				if (this.txtDescription.Text.Trim().Length > 0)
					strDesc = p_ado.FixString(this.txtDescription.Text.Trim(),"'","''");
				string strConn = p_ado.getMDBConnString(strFullPath.ToString(),"admin","");
				//string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFullPath.ToString() + ";User Id=admin;Password=;";
				strSQL = "UPDATE scenario SET description = '" + 
					strDesc + 
					"' WHERE scenario_id = '" + this.txtScenarioId.Text.ToLower() + "';";
				p_ado.SqlNonQuery(strConn,strSQL);
				p_ado=null;

			}
			if (ScenarioType.Trim().ToUpper() =="OPTIMIZER")
			{
				if (((frmOptimizerScenario)this.ParentForm).m_bScenarioOpen == false) 
				{
					((frmOptimizerScenario)this.ParentForm).Text = "Core Analysis: Optimization Scenario (" + this.txtScenarioId.Text.Trim() + ")";
					((frmOptimizerScenario)this.ParentForm).SetMenu("scenario");
					((frmOptimizerScenario)this.ParentForm).m_bScenarioOpen = true;
					this.lblTitle.Text = "";
					this.Visible=false;
				}
				
			}
			else 
			{
				if (this.ReferenceProcessorScenarioForm.m_bScenarioOpen==false)
				{
					this.ReferenceProcessorScenarioForm.Text = "Processor: Scenario (" + this.txtScenarioId.Text.Trim() + ")";
					this.ReferenceProcessorScenarioForm.m_bScenarioOpen=true;
					this.lblTitle.Text = "";
					this.Visible=false;
				}
			}

		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			if (ScenarioType.Trim().ToUpper()=="OPTIMIZER")
			{
			
				if (((frmOptimizerScenario)this.ParentForm).m_bScenarioOpen == false) 
				{
					((frmOptimizerScenario)this.ParentForm).Close();
				}
				else 
				{
					this.lblTitle.Text = "";
					//((frmScenario)this.ParentForm).lblTitle.Text = "";
					((frmOptimizerScenario)this.ParentForm).SetMenu("scenario");
					this.Visible =false;
					//v309((frmScenario)this.ParentForm).Height = ((frmScenario)this.ParentForm).grpboxMenu.Height*2;
				}
			}
			else
			{
				if (this.ReferenceProcessorScenarioForm.m_bScenarioOpen == false) 
				{
					this.ReferenceProcessorScenarioForm.Close();
				}
				else 
				{
					this.lblTitle.Text = "";
					this.Visible =false;
				}
			}

		
		}
		private void lstScenario_Click(object sender, System.EventArgs e)
		{
			
		   
		
		}

		private void txtDescription_TextChanged(object sender, System.EventArgs e)
		{
			
			
 

		}

		public void OpenScenario()
		{
            if (ReferenceProcessorScenarioForm != null && 
                ReferenceProcessorScenarioForm.m_bUsingSqlite == true)
            {
                this.populate_scenario_listbox_sqlite();
            }
            else
            {
                this.populate_scenario_listbox();
            }
            	        
			this.btnCancel.Enabled = true;
			this.btnOpen.Enabled = true;
			this.txtDescription.Enabled=false;

			this.txtScenarioPath.Enabled=false;
			this.txtScenarioId.Enabled=false;
			this.lstScenario.Enabled = true;
		}
		private void RefreshForm()
		{
			
			string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\db";
			string strFile = "scenario_" + ScenarioType + "_rule_definitions.mdb"; 
			StringBuilder strFullPath = new StringBuilder(strScenarioDir);
			strFullPath.Append("\\");
			strFullPath.Append(strFile);
			ado_data_access p_ado = new ado_data_access();
			string strConn = p_ado.getMDBConnString(strFullPath.ToString(),"admin","");
			//string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFullPath.ToString() + ";User Id=admin;Password=;";
			string strSQL = "select * from scenario where scenario_id = '" + this.lstScenario.SelectedItem.ToString().Trim() + "';";
			
            p_ado.SqlQueryReader(strConn,strSQL);
			if (p_ado.m_intError==0)
			{
				try
				{
					while (p_ado.m_OleDbDataReader.Read())
					{

						this.txtDescription.Text =  p_ado.m_OleDbDataReader["description"].ToString();
						this.txtScenarioPath.Text =  p_ado.m_OleDbDataReader["path"].ToString();
						this.txtScenarioId.Text = p_ado.m_OleDbDataReader["scenario_id"].ToString();
						//((frmScenario)this.ParentForm).uc_scenario_notes1.txtNotes.Text = p_ado.m_OleDbDataReader["notes"].ToString();
						break;
					}
					p_ado.m_OleDbDataReader.Close();

				}
				catch (Exception caught)
				{
                  this.strError = caught.Message;
				  this.intError=-1;
				  MessageBox.Show(this.strError);
				}
				p_ado.m_OleDbDataReader = null;
				p_ado.m_OleDbCommand = null;
				p_ado.m_OleDbConnection.Close();
				p_ado.m_OleDbConnection = null;
			}
		}
        private void RefreshFormSqlite()
        {
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\" +
                    Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            string strConn = dataMgr.GetConnectionString(strScenarioDir);
            try
            {
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    con.Open();
                    string strSQL = "select * from scenario where scenario_id = '" + this.lstScenario.SelectedItem.ToString() + "';";
                    dataMgr.SqlQueryReader(con, strSQL);
                    if (dataMgr.m_DataReader.HasRows)
                    {
                        while (dataMgr.m_DataReader.Read())
                        {
                            txtScenarioId.Text = dataMgr.m_DataReader["scenario_id"].ToString();
                            txtDescription.Text = dataMgr.m_DataReader["description"].ToString();
                            txtScenarioPath.Text = dataMgr.m_DataReader["path"].ToString();
                            break;

                        }
                    }
                    dataMgr.m_DataReader.Close();
                }
            }
            catch (Exception caught)
            {
                intError = -1;
                strError = caught.Message;
                MessageBox.Show(strError, "FIA Biosum");
            }
        }
		private void txtDescription_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
     	   //if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
		}

		private void lstScenario_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.lstScenario.SelectedIndex >= 0) 
			{
                if (ReferenceProcessorScenarioForm != null && 
                    ReferenceProcessorScenarioForm.m_bUsingSqlite == true)
                {
                    this.RefreshFormSqlite();
                }
                else
                {
                    this.RefreshForm();
                }                
			}
		}

		private void uc_scenario_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (ScenarioType.Trim().ToUpper()=="OPTIMIZER")
			{
				((frmOptimizerScenario)this.ParentForm).m_bPopup = false;
			}
			else
			{
				this.ReferenceProcessorScenarioForm.m_bPopup=false;
			}
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
            if (ScenarioType.Trim().ToUpper() == "OPTIMIZER")
				((frmOptimizerScenario)this.ParentForm).Close();
			else
				this.ReferenceProcessorScenarioForm.Close();

		}

		private void txtScenarioId_Leave(object sender, System.EventArgs e)
		{
			if (this.txtScenarioId.Text.Length > 0) 
			{
				this.txtScenarioId.Text = this.txtScenarioId.Text.Trim();
				//replace spaces with underscores
				this.txtScenarioId.Text = this.txtScenarioId.Text.Replace(" ","_");
				int intLastDir = this.txtScenarioPath.Text.LastIndexOf("\\");
				this.txtScenarioPath.Text = this.txtScenarioPath.Text.Substring(0,intLastDir) + "\\" + this.txtScenarioId.Text;
			}
		}

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
		
		}
		public FIA_Biosum_Manager.frmOptimizerScenario ReferenceCoreScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
		public FIA_Biosum_Manager.frmProcessorScenario ReferenceProcessorScenarioForm
		{
			get {return _frmProcessorScenario;}
			set {_frmProcessorScenario=value;}
		}
		public string ScenarioType
		{
			get {return this._strScenarioType;}
			set {this._strScenarioType=value;}
		}
		
	}
}

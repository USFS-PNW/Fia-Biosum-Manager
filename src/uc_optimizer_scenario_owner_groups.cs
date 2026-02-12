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
	/// Summary description for uc_scenario_owner_groups.
	/// </summary>
	public class uc_optimizer_scenario_owner_groups : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.ImageList imgSize;
		private System.ComponentModel.IContainer components;
		public int m_intFullHt = 304;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.CheckBox chkOwnGrp10;
		public System.Windows.Forms.CheckBox chkOwnGrp20;
		public System.Windows.Forms.CheckBox chkOwnGrp30;
		public System.Windows.Forms.CheckBox chkOwnGrp40;
		private System.Windows.Forms.GroupBox groupBox2;
		public string strScenarioId;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario=null;


		public uc_optimizer_scenario_owner_groups()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_optimizer_scenario_owner_groups));
            this.imgSize = new System.Windows.Forms.ImageList(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkOwnGrp10 = new System.Windows.Forms.CheckBox();
            this.chkOwnGrp20 = new System.Windows.Forms.CheckBox();
            this.chkOwnGrp30 = new System.Windows.Forms.CheckBox();
            this.chkOwnGrp40 = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // imgSize
            // 
            this.imgSize.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgSize.ImageStream")));
            this.imgSize.TransparentColor = System.Drawing.Color.Transparent;
            this.imgSize.Images.SetKeyName(0, "");
            this.imgSize.Images.SetKeyName(1, "");
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.chkOwnGrp10);
            this.groupBox1.Controls.Add(this.chkOwnGrp20);
            this.groupBox1.Controls.Add(this.chkOwnGrp30);
            this.groupBox1.Controls.Add(this.chkOwnGrp40);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(100, 51);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(487, 221);
            this.groupBox1.TabIndex = 21;
            this.groupBox1.TabStop = false;
            // 
            // chkOwnGrp10
            // 
            this.chkOwnGrp10.Checked = true;
            this.chkOwnGrp10.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkOwnGrp10.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.chkOwnGrp10.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkOwnGrp10.ForeColor = System.Drawing.Color.Black;
            this.chkOwnGrp10.Location = new System.Drawing.Point(8, 18);
            this.chkOwnGrp10.Name = "chkOwnGrp10";
            this.chkOwnGrp10.Size = new System.Drawing.Size(250, 40);
            this.chkOwnGrp10.TabIndex = 17;
            this.chkOwnGrp10.Text = "U.S. Forest Service";
            this.chkOwnGrp10.Click += new System.EventHandler(this.chkOwnGrp10_Click);
            // 
            // chkOwnGrp20
            // 
            this.chkOwnGrp20.Checked = true;
            this.chkOwnGrp20.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkOwnGrp20.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.chkOwnGrp20.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkOwnGrp20.ForeColor = System.Drawing.Color.Black;
            this.chkOwnGrp20.Location = new System.Drawing.Point(8, 64);
            this.chkOwnGrp20.Name = "chkOwnGrp20";
            this.chkOwnGrp20.Size = new System.Drawing.Size(200, 40);
            this.chkOwnGrp20.TabIndex = 18;
            this.chkOwnGrp20.Text = "Other Federal";
            this.chkOwnGrp20.Click += new System.EventHandler(this.chkOwnGrp20_Click);
            // 
            // chkOwnGrp30
            // 
            this.chkOwnGrp30.Checked = true;
            this.chkOwnGrp30.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkOwnGrp30.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.chkOwnGrp30.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkOwnGrp30.ForeColor = System.Drawing.Color.Black;
            this.chkOwnGrp30.Location = new System.Drawing.Point(8, 112);
            this.chkOwnGrp30.Name = "chkOwnGrp30";
            this.chkOwnGrp30.Size = new System.Drawing.Size(391, 40);
            this.chkOwnGrp30.TabIndex = 19;
            this.chkOwnGrp30.Text = "State, County, And Local Government";
            this.chkOwnGrp30.Click += new System.EventHandler(this.chkOwnGrp30_Click);
            // 
            // chkOwnGrp40
            // 
            this.chkOwnGrp40.Checked = true;
            this.chkOwnGrp40.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkOwnGrp40.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.chkOwnGrp40.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkOwnGrp40.ForeColor = System.Drawing.Color.Black;
            this.chkOwnGrp40.Location = new System.Drawing.Point(8, 160);
            this.chkOwnGrp40.Name = "chkOwnGrp40";
            this.chkOwnGrp40.Size = new System.Drawing.Size(200, 40);
            this.chkOwnGrp40.TabIndex = 20;
            this.chkOwnGrp40.Text = "Private";
            this.chkOwnGrp40.Click += new System.EventHandler(this.chkOwnGrp40_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox2.Controls.Add(this.lblTitle);
            this.groupBox2.Controls.Add(this.groupBox1);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 0);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(696, 304);
            this.groupBox2.TabIndex = 22;
            this.groupBox2.TabStop = false;
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(690, 32);
            this.lblTitle.TabIndex = 26;
            this.lblTitle.Text = "Owner Groups";
            // 
            // uc_scenario_owner_groups
            // 
            this.Controls.Add(this.groupBox2);
            this.Name = "uc_scenario_owner_groups";
            this.Size = new System.Drawing.Size(696, 304);
            this.Resize += new System.EventHandler(this.uc_scenario_owner_groups_Resize);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void grpboxOwnerGroups_Resize(object sender, System.EventArgs e)
		{
			try
			{
				//v309this.cmdOwnerGroups.Left = (int)(this.grpboxOwnerGroups.Width * .50) - (int)(this.cmdOwnerGroups.Width *.50);
				//v309this.btnHelp.Left = this.grpboxOwnerGroups.Width - this.btnHelp.Width - 4;
				//v309this.groupBox1.Left = Convert.ToInt32((this.grpboxOwnerGroups.Width * .50) - (this.groupBox1.Width * .50));
				

			}
			catch
			{
			}

		}

		public void loadvalues()
		{
			const int LANDOWNER_FORESTSERVICE = 10;
			const int LANDOWNER_OTHERFEDERAL = 20;
			const int LANDOWNER_STATELOCAL = 30;
			const int LANDOWNER_PRIVATE = 40;
            
					
			int x=0;
			
            string[] strArray = frmMain.g_oUtils.ConvertListToArray(ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.OwnerGroupCodeList,",");
			this.chkOwnGrp10.Checked=false;
			this.chkOwnGrp20.Checked=false;
			this.chkOwnGrp30.Checked=false;
			this.chkOwnGrp40.Checked=false;
            if (strArray != null)
            {
                for (x=0;x<=strArray.Length - 1;x++)
                {
						    switch (Convert.ToInt32(strArray[x]))
						    {
							    case LANDOWNER_FORESTSERVICE:
								    this.chkOwnGrp10.Checked=true;
								    break;
							    case LANDOWNER_OTHERFEDERAL:
								    this.chkOwnGrp20.Checked=true;
								    break;
							    case LANDOWNER_STATELOCAL:
								    this.chkOwnGrp30.Checked=true;
								    break;
							    case LANDOWNER_PRIVATE:
								    this.chkOwnGrp40.Checked=true;
								    break;
						    }
			
			    }
            }
			((frmOptimizerScenario)this.ParentForm).m_bSave=false;
			if (this.chkOwnGrp10.Checked==false &&
				this.chkOwnGrp20.Checked==false &&
				this.chkOwnGrp30.Checked==false &&
				this.chkOwnGrp40.Checked==false) 
			{
				this.chkOwnGrp10.Checked=true;
				this.chkOwnGrp20.Checked=true;
				this.chkOwnGrp30.Checked=true;
				this.chkOwnGrp40.Checked=true;
			}
		}

		public int savevalues()
		{
			int x = 0;


			string strSQL = "";
			string strConn = "";


			DataMgr oDataMgr = new DataMgr();
			string strScenarioId = ((frmOptimizerScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text.Trim();
			string strScenarioDB =
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\" +
				Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;

			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strScenarioDB)))
            {
				conn.Open();
				if (oDataMgr.m_intError != 0)
				{
					x = oDataMgr.m_intError;
					oDataMgr = null;
					return x;
				}

				//delete all records from the scenario wind speed class table
				strSQL = "DELETE FROM scenario_land_owner_groups WHERE " +
					" TRIM(UPPER(scenario_id)) = '" + strScenarioId.Trim().ToUpper() + "';";

				oDataMgr.SqlNonQuery(conn, strSQL);
				if (oDataMgr.m_intError < 0)
				{
					conn.Close();
					x = oDataMgr.m_intError;
					oDataMgr = null;
					return x;
				}

				if (this.chkOwnGrp10.Checked == true)
				{
					oDataMgr.m_strSQL = "INSERT INTO scenario_land_owner_groups (scenario_id,owngrpcd)" +
						" VALUES ('" + strScenarioId + "',10);";
					oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
				}

				if (this.chkOwnGrp20.Checked == true)
				{
					oDataMgr.m_strSQL = "INSERT INTO scenario_land_owner_groups (scenario_id,owngrpcd)" +
						" VALUES ('" + strScenarioId + "',20);";
					oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
				}

				if (this.chkOwnGrp30.Checked == true)
				{
					oDataMgr.m_strSQL = "INSERT INTO scenario_land_owner_groups (scenario_id,owngrpcd)" +
						" VALUES ('" + strScenarioId + "',30);";
					oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
				}
				if (this.chkOwnGrp40.Checked == true)
				{
					oDataMgr.m_strSQL = "INSERT INTO scenario_land_owner_groups (scenario_id,owngrpcd)" +
						" VALUES ('" + strScenarioId + "',40);";
					oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
				}
			}
			return 0;
		}
		public int ValInput()
		{
			if (this.chkOwnGrp10.Checked==false &&
				this.chkOwnGrp20.Checked==false &&
				this.chkOwnGrp30.Checked==false &&
				this.chkOwnGrp40.Checked==false) 
			{
				MessageBox.Show("Run Scenario Failed: Select at least one ownership group","FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return -1;
			}
			return 0;

		}

		private void chkOwnGrp10_Click(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void chkOwnGrp20_Click(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void chkOwnGrp30_Click(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
			
		}

		private void chkOwnGrp40_Click(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void uc_scenario_owner_groups_Resize(object sender, System.EventArgs e)
		{
			this.groupBox1.Left = (int)(this.ClientSize.Width * .5) - (int)(this.groupBox1.Width * .5);
			this.groupBox1.Top = (int)(this.ClientSize.Height * .5) - (int)(this.groupBox1.Height * .5);
		}
	
		public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}

		
	}
}

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
	/// Summary description for uc_scenario_psite.
	/// </summary>
	public class uc_optimizer_scenario_psite : System.Windows.Forms.UserControl
	{
		private ListViewEmbeddedControls.ListViewEx lstPSites;
		private System.Windows.Forms.ImageList imgSize;
		private System.Windows.Forms.ListView listView1;
		private System.ComponentModel.IContainer components;
		private string m_strTravelTimeTable;
		private string m_strPSiteTable;
		private string m_strTempDBFile;
		private env m_oEnv;
		public int m_intError=0;
		public string m_strError="";
		private System.Windows.Forms.Button btnSelectAll;
		private System.Windows.Forms.Button btnUnselectAll;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.ComboBox m_Combo;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario=null;
		private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvRowColors = new ListViewAlternateBackgroundColors();
        private ListViewColumnSorter lvwColumnSorter;

	    const int COLUMN_CHECKBOX=0;
		const int COLUMN_PSITEID=1;
		const int COLUMN_PSITECN=2;
		const int COLUMN_PSITENAME=3;
		const int COLUMN_PSITEEXIST=4;
		const int COLUMN_PSITEROADRAIL=5;
		const int COLUMN_PSITEBIOPROCESSTYPE=6;
		const int COLUMN_PSITESTATE=7;
		const int COLUMN_PSITECOUNTY=8;


		public uc_optimizer_scenario_psite()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
            lstPSites = new ListViewEmbeddedControls.ListViewEx();

			this.groupBox1.Controls.Add(lstPSites);
		    lstPSites.Size = this.listView1.Size;
			lstPSites.Location = this.listView1.Location;
			lstPSites.CheckBoxes = true;
			lstPSites.AllowColumnReorder=true;
			lstPSites.FullRowSelect = false;
			lstPSites.GridLines = true;
			lstPSites.MultiSelect=false;
			lstPSites.View = System.Windows.Forms.View.Details;
			lstPSites.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(lstPSites_ColumnClick);
			lstPSites.ItemCheck += new ItemCheckEventHandler(lstPSites_ItemCheck);
			lstPSites.MouseUp += new System.Windows.Forms.MouseEventHandler(lstPSites_MouseUp);
			lstPSites.SelectedIndexChanged += new System.EventHandler(lstPSites_SelectedIndexChanged);
			
			// 

			listView1.Hide();
			lstPSites.Show();

			if (frmMain.g_oGridViewFont != null) lstPSites.Font = frmMain.g_oGridViewFont;
			this.m_oLvRowColors.ReferenceListView = this.lstPSites;
			this.m_oLvRowColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvRowColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvRowColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvRowColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvRowColors.CustomFullRowSelect=true;

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

		public void loadvalues()
        {
			int x;
            byte byteTranCd = 9;
            byte byteBioCd = 9;
            int intPSiteId;

            this.lstPSites.Clear();
            this.m_oLvRowColors.InitializeRowCollection();

            this.lstPSites.Columns.Add("", 2, HorizontalAlignment.Left);
            this.lstPSites.Columns.Add("PSite ID", 55, HorizontalAlignment.Left);
            this.lstPSites.Columns.Add("PSite CN", 90, HorizontalAlignment.Left);
            this.lstPSites.Columns.Add("Name", 180, HorizontalAlignment.Left);
            this.lstPSites.Columns.Add("Exists", 40, HorizontalAlignment.Left);
            this.lstPSites.Columns.Add("Site Type", 110, HorizontalAlignment.Left);
            this.lstPSites.Columns.Add("Processing Type", 95, HorizontalAlignment.Left);
            this.lstPSites.Columns.Add("State", 40, HorizontalAlignment.Left);
            this.lstPSites.Columns.Add("County", 75, HorizontalAlignment.Left);

            this.lstPSites.Columns[COLUMN_CHECKBOX].Width = -2;

            // Create an instance of a ListView column sorter and assign it 
            // to the ListView control.
            lvwColumnSorter = new ListViewColumnSorter();
            this.lstPSites.ListViewItemSorter = lvwColumnSorter;

			for (x = 0; x <= ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Count - 1; x++)
			{
				//null column
				this.lstPSites.Items.Add(" ");
				this.lstPSites.Items[lstPSites.Items.Count - 1].UseItemStyleForSubItems = false;
				this.m_oLvRowColors.AddRow();
				this.m_oLvRowColors.AddColumns(lstPSites.Items.Count - 1, lstPSites.Columns.Count);

				//psite_id
				string strPsiteId = ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteId.Trim('\'');
				this.lstPSites.Items[lstPSites.Items.Count - 1].SubItems.Add(strPsiteId);
				this.m_oLvRowColors.ListViewSubItem(lstPSites.Items.Count - 1,
									COLUMN_PSITEID,
									lstPSites.Items[lstPSites.Items.Count - 1].SubItems[COLUMN_PSITEID], false);

				//psite cn
				this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems.Add(ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteCN);
				this.m_oLvRowColors.ListViewSubItem(lstPSites.Items.Count - 1,
									COLUMN_PSITECN,
									lstPSites.Items[lstPSites.Items.Count - 1].SubItems[COLUMN_PSITECN], false);

				//name
				this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems.Add(ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteName);
				this.m_oLvRowColors.ListViewSubItem(lstPSites.Items.Count - 1,
									COLUMN_PSITENAME,
									lstPSites.Items[lstPSites.Items.Count - 1].SubItems[COLUMN_PSITENAME], false);

				//exists
				this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems.Add(ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteExistYN);
				this.m_oLvRowColors.ListViewSubItem(lstPSites.Items.Count - 1,
									COLUMN_PSITEEXIST,
									lstPSites.Items[lstPSites.Items.Count - 1].SubItems[COLUMN_PSITEEXIST], false);

				//processing site type
				int intSubItemCount = 0;
				byteTranCd = Convert.ToByte(ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).TransportationCode);
				if (byteTranCd == 1)
				{
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems.Add("Facility on Road");
					intSubItemCount = this.lstPSites.Items[lstPSites.Items.Count - 1].SubItems.Count;
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems[intSubItemCount - 1].Font = new Font("Microsoft Sans Serif", (float)8.25, System.Drawing.FontStyle.Regular);

				}
				else if (byteTranCd == 2)
				{
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems.Add("Intermodal Transfer");
					intSubItemCount = this.lstPSites.Items[lstPSites.Items.Count - 1].SubItems.Count;
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems[intSubItemCount - 1].Font = new Font("Microsoft Sans Serif", (float)8.25, System.Drawing.FontStyle.Regular);
				}
				else
				{
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems.Add("Facility on Rail");
					intSubItemCount = this.lstPSites.Items[lstPSites.Items.Count - 1].SubItems.Count;
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems[intSubItemCount - 1].Font = new Font("Microsoft Sans Serif", (float)8.25, System.Drawing.FontStyle.Regular);
				}
				this.m_oLvRowColors.ListViewSubItem(lstPSites.Items.Count - 1,
                                            COLUMN_PSITEROADRAIL,
                                            lstPSites.Items[lstPSites.Items.Count - 1].SubItems[COLUMN_PSITEROADRAIL], false);

                //bio processing type
                intSubItemCount = 0;
				byteTranCd = Convert.ToByte(ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).BiomassCode);
				if (byteTranCd == 2)
				{
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems.Add("Chips");
					intSubItemCount = this.lstPSites.Items[lstPSites.Items.Count - 1].SubItems.Count;
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems[intSubItemCount - 1].Font = new Font("Microsoft Sans Serif", (float)8.25, System.Drawing.FontStyle.Regular);
				}
				else if (byteTranCd == 3)
				{
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems.Add("Both");
					intSubItemCount = this.lstPSites.Items[lstPSites.Items.Count - 1].SubItems.Count;
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems[intSubItemCount - 1].Font = new Font("Microsoft Sans Serif", (float)8.25, System.Drawing.FontStyle.Regular);
				}
				else if (byteTranCd == 4)
				{
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems.Add("Other");
					intSubItemCount = this.lstPSites.Items[lstPSites.Items.Count - 1].SubItems.Count;
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems[intSubItemCount - 1].Font = new Font("Microsoft Sans Serif", (float)8.25, System.Drawing.FontStyle.Regular);
				}
				else
				{
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems.Add("Merch");
					intSubItemCount = this.lstPSites.Items[lstPSites.Items.Count - 1].SubItems.Count;
					this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems[intSubItemCount - 1].Font = new Font("Microsoft Sans Serif", (float)8.25, System.Drawing.FontStyle.Regular);
				}
				this.m_oLvRowColors.ListViewSubItem(lstPSites.Items.Count - 1,
                                            COLUMN_PSITEBIOPROCESSTYPE,
                                            lstPSites.Items[lstPSites.Items.Count - 1].SubItems[COLUMN_PSITEBIOPROCESSTYPE], false);

                //state
                this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems.Add(ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).State);
				this.m_oLvRowColors.ListViewSubItem(lstPSites.Items.Count - 1,
									COLUMN_PSITESTATE,
									lstPSites.Items[lstPSites.Items.Count - 1].SubItems[COLUMN_PSITESTATE], false);

				//county
				this.lstPSites.Items[this.lstPSites.Items.Count - 1].SubItems.Add(ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).County);
				this.m_oLvRowColors.ListViewSubItem(lstPSites.Items.Count - 1,
									COLUMN_PSITECOUNTY,
									lstPSites.Items[lstPSites.Items.Count - 1].SubItems[COLUMN_PSITECOUNTY], false);

				//selected
				if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).Selected)
				{
					lstPSites.Items[x].Checked = true;
				}
				else
				{
					lstPSites.Items[x].Checked = false;
				}
			}
		}

		

		public int savevalues()
		{
			string strTranCd;
			string strBioCd;
			string strSelected;
			string strName;
			string strScenarioId;
			string strPSiteId;
			int x;

			DataMgr oDataMgr = new DataMgr();
			strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim();
			string strScenarioDB =
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
				Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;

			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strScenarioDB)))
			{
				conn.Open();
				if (oDataMgr.m_intError != 0)
				{
					x = oDataMgr.m_intError;
					oDataMgr = null;
					return x;
				}

				//delete all records from the scenario psites table
				oDataMgr.m_strSQL = "DELETE FROM scenario_psites WHERE " +
					" TRIM(UPPER(scenario_id)) = '" + strScenarioId.Trim().ToUpper() + "';";

				oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
				if (oDataMgr.m_intError < 0)
				{
					conn.Close();
					x = oDataMgr.m_intError;
					oDataMgr = null;
					return x;
				}
				for (x = 0; x <= this.lstPSites.Items.Count - 1; x++)
				{
					strTranCd = "";
					strBioCd = "";
					strSelected = "";
					strName = "";
					strScenarioId = "";
					strPSiteId = "";

					oDataMgr.m_strSQL = "INSERT INTO scenario_psites (scenario_id,psite_id,name,trancd,biocd,selected_yn)" +
								   " VALUES ";

					strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim();
					strPSiteId = lstPSites.Items[x].SubItems[COLUMN_PSITEID].Text.Trim();
					strName = lstPSites.Items[x].SubItems[COLUMN_PSITENAME].Text.Trim();
					strName = oDataMgr.FixString(strName.Trim(), "'", "''");
					if (lstPSites.Items[x].Checked == true)
					{
						strSelected = "Y";
					}
					else
					{
						strSelected = "N";
					}
					if (lstPSites.Items[x].SubItems[COLUMN_PSITEROADRAIL].Text.Trim() == "Facility on Road")
					{
						strTranCd = "1";
					}
					else if (lstPSites.Items[x].SubItems[COLUMN_PSITEROADRAIL].Text.Trim() == "Intermodal Transfer")
                    {
						strTranCd = "2";
                    }
					else if (lstPSites.Items[x].SubItems[COLUMN_PSITEROADRAIL].Text.Trim() == "Facility on Rail")
                    {
						strTranCd = "3";
                    }
					else
                    {
						strTranCd = "9";
                    }
					if (lstPSites.Items[x].SubItems[COLUMN_PSITEBIOPROCESSTYPE].Text.Trim() == "Merch")
                    {
						strBioCd = "1";
                    }
					else if (lstPSites.Items[x].SubItems[COLUMN_PSITEBIOPROCESSTYPE].Text.Trim() == "Chips")
                    {
						strBioCd = "2";
                    }
					else if (lstPSites.Items[x].SubItems[COLUMN_PSITEBIOPROCESSTYPE].Text.Trim() == "Both")
                    {
						strBioCd = "3";
                    }
                    else if (lstPSites.Items[x].SubItems[COLUMN_PSITEBIOPROCESSTYPE].Text.Trim() == "Other")
                    {
                        strBioCd = "4";
                    }
                    else
                    {
						strBioCd = "9";
                    }
					oDataMgr.m_strSQL = oDataMgr.m_strSQL + "('" + strScenarioId + "'," +
														   strPSiteId + ",'" +
														   strName + "'," +
														   strTranCd + "," +
														   strBioCd + ",'" +
														   strSelected + "')";
					oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
					if (oDataMgr.m_intError != 0) break;
				}
				x = oDataMgr.m_intError;
			}
			return x;
		}
		public int val_psites()
		{
			if (this.lstPSites.CheckedItems.Count == 0)
			{
				MessageBox.Show("Run Scenario Failed: Select at least one processing site in <Wood Processing Sites>","FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return -1;
			}
			return 0;

		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(uc_optimizer_scenario_psite));
			this.imgSize = new System.Windows.Forms.ImageList(this.components);
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblTitle = new System.Windows.Forms.Label();
			//this.btnScenarioPSiteUpdate = new System.Windows.Forms.Button();
			this.btnUnselectAll = new System.Windows.Forms.Button();
			this.btnSelectAll = new System.Windows.Forms.Button();
			//this.btnScenarioPSiteDefault = new System.Windows.Forms.Button();
			this.listView1 = new System.Windows.Forms.ListView();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// imgSize
			// 
			this.imgSize.ImageSize = ((System.Drawing.Size)(resources.GetObject("imgSize.ImageSize")));
			this.imgSize.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgSize.ImageStream")));
			this.imgSize.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// groupBox1
			// 
			this.groupBox1.AccessibleDescription = resources.GetString("groupBox1.AccessibleDescription");
			this.groupBox1.AccessibleName = resources.GetString("groupBox1.AccessibleName");
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("groupBox1.Anchor")));
			this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("groupBox1.BackgroundImage")));
			this.groupBox1.Controls.Add(this.lblTitle);
			//this.groupBox1.Controls.Add(this.btnScenarioPSiteUpdate);
			this.groupBox1.Controls.Add(this.btnUnselectAll);
			this.groupBox1.Controls.Add(this.btnSelectAll);
			//this.groupBox1.Controls.Add(this.btnScenarioPSiteDefault);
			this.groupBox1.Controls.Add(this.listView1);
			this.groupBox1.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("groupBox1.Dock")));
			this.groupBox1.Enabled = ((bool)(resources.GetObject("groupBox1.Enabled")));
			this.groupBox1.Font = ((System.Drawing.Font)(resources.GetObject("groupBox1.Font")));
			this.groupBox1.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("groupBox1.ImeMode")));
			this.groupBox1.Location = ((System.Drawing.Point)(resources.GetObject("groupBox1.Location")));
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("groupBox1.RightToLeft")));
			this.groupBox1.Size = ((System.Drawing.Size)(resources.GetObject("groupBox1.Size")));
			this.groupBox1.TabIndex = ((int)(resources.GetObject("groupBox1.TabIndex")));
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = resources.GetString("groupBox1.Text");
			this.toolTip1.SetToolTip(this.groupBox1, resources.GetString("groupBox1.ToolTip"));
			this.groupBox1.Visible = ((bool)(resources.GetObject("groupBox1.Visible")));
			this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
			// 
			// lblTitle
			// 
			this.lblTitle.AccessibleDescription = resources.GetString("lblTitle.AccessibleDescription");
			this.lblTitle.AccessibleName = resources.GetString("lblTitle.AccessibleName");
			this.lblTitle.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("lblTitle.Anchor")));
			this.lblTitle.AutoSize = ((bool)(resources.GetObject("lblTitle.AutoSize")));
			this.lblTitle.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("lblTitle.Dock")));
			this.lblTitle.Enabled = ((bool)(resources.GetObject("lblTitle.Enabled")));
			this.lblTitle.Font = ((System.Drawing.Font)(resources.GetObject("lblTitle.Font")));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Image = ((System.Drawing.Image)(resources.GetObject("lblTitle.Image")));
			this.lblTitle.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("lblTitle.ImageAlign")));
			this.lblTitle.ImageIndex = ((int)(resources.GetObject("lblTitle.ImageIndex")));
			this.lblTitle.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("lblTitle.ImeMode")));
			this.lblTitle.Location = ((System.Drawing.Point)(resources.GetObject("lblTitle.Location")));
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("lblTitle.RightToLeft")));
			this.lblTitle.Size = ((System.Drawing.Size)(resources.GetObject("lblTitle.Size")));
			this.lblTitle.TabIndex = ((int)(resources.GetObject("lblTitle.TabIndex")));
			this.lblTitle.Text = resources.GetString("lblTitle.Text");
			this.lblTitle.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("lblTitle.TextAlign")));
			this.toolTip1.SetToolTip(this.lblTitle, resources.GetString("lblTitle.ToolTip"));
			this.lblTitle.Visible = ((bool)(resources.GetObject("lblTitle.Visible")));
			// 
			// btnScenarioPSiteUpdate
			// 
			//this.btnScenarioPSiteUpdate.AccessibleDescription = resources.GetString("btnScenarioPSiteUpdate.AccessibleDescription");
			//this.btnScenarioPSiteUpdate.AccessibleName = resources.GetString("btnScenarioPSiteUpdate.AccessibleName");
			//this.btnScenarioPSiteUpdate.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnScenarioPSiteUpdate.Anchor")));
			//this.btnScenarioPSiteUpdate.BackColor = System.Drawing.SystemColors.Control;
			//this.btnScenarioPSiteUpdate.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnScenarioPSiteUpdate.BackgroundImage")));
			//this.btnScenarioPSiteUpdate.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnScenarioPSiteUpdate.Dock")));
			//this.btnScenarioPSiteUpdate.Enabled = ((bool)(resources.GetObject("btnScenarioPSiteUpdate.Enabled")));
			//this.btnScenarioPSiteUpdate.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnScenarioPSiteUpdate.FlatStyle")));
			//this.btnScenarioPSiteUpdate.Font = ((System.Drawing.Font)(resources.GetObject("btnScenarioPSiteUpdate.Font")));
			//this.btnScenarioPSiteUpdate.ForeColor = System.Drawing.Color.Black;
			//this.btnScenarioPSiteUpdate.Image = ((System.Drawing.Image)(resources.GetObject("btnScenarioPSiteUpdate.Image")));
			//this.btnScenarioPSiteUpdate.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnScenarioPSiteUpdate.ImageAlign")));
			//this.btnScenarioPSiteUpdate.ImageIndex = ((int)(resources.GetObject("btnScenarioPSiteUpdate.ImageIndex")));
			//this.btnScenarioPSiteUpdate.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnScenarioPSiteUpdate.ImeMode")));
			//this.btnScenarioPSiteUpdate.Location = ((System.Drawing.Point)(resources.GetObject("btnScenarioPSiteUpdate.Location")));
			//this.btnScenarioPSiteUpdate.Name = "btnScenarioPSiteUpdate";
			//this.btnScenarioPSiteUpdate.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnScenarioPSiteUpdate.RightToLeft")));
			//this.btnScenarioPSiteUpdate.Size = ((System.Drawing.Size)(resources.GetObject("btnScenarioPSiteUpdate.Size")));
			//this.btnScenarioPSiteUpdate.TabIndex = ((int)(resources.GetObject("btnScenarioPSiteUpdate.TabIndex")));
			//this.btnScenarioPSiteUpdate.Tag = "";
			//this.btnScenarioPSiteUpdate.Text = resources.GetString("btnScenarioPSiteUpdate.Text");
			//this.btnScenarioPSiteUpdate.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnScenarioPSiteUpdate.TextAlign")));
			//this.toolTip1.SetToolTip(this.btnScenarioPSiteUpdate, resources.GetString("btnScenarioPSiteUpdate.ToolTip"));
			//this.btnScenarioPSiteUpdate.Visible = ((bool)(resources.GetObject("btnScenarioPSiteUpdate.Visible")));
			//this.btnScenarioPSiteUpdate.Click += new System.EventHandler(this.btnScenarioPSiteUpdate_Click);
			// 
			// btnUnselectAll
			// 
			this.btnUnselectAll.AccessibleDescription = resources.GetString("btnUnselectAll.AccessibleDescription");
			this.btnUnselectAll.AccessibleName = resources.GetString("btnUnselectAll.AccessibleName");
			this.btnUnselectAll.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnUnselectAll.Anchor")));
			this.btnUnselectAll.BackColor = System.Drawing.SystemColors.Control;
			this.btnUnselectAll.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnUnselectAll.BackgroundImage")));
			this.btnUnselectAll.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnUnselectAll.Dock")));
			this.btnUnselectAll.Enabled = ((bool)(resources.GetObject("btnUnselectAll.Enabled")));
			this.btnUnselectAll.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnUnselectAll.FlatStyle")));
			this.btnUnselectAll.Font = ((System.Drawing.Font)(resources.GetObject("btnUnselectAll.Font")));
			this.btnUnselectAll.ForeColor = System.Drawing.Color.Black;
			this.btnUnselectAll.Image = ((System.Drawing.Image)(resources.GetObject("btnUnselectAll.Image")));
			this.btnUnselectAll.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnUnselectAll.ImageAlign")));
			this.btnUnselectAll.ImageIndex = ((int)(resources.GetObject("btnUnselectAll.ImageIndex")));
			this.btnUnselectAll.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnUnselectAll.ImeMode")));
			this.btnUnselectAll.Location = ((System.Drawing.Point)(resources.GetObject("btnUnselectAll.Location")));
			this.btnUnselectAll.Name = "btnUnselectAll";
			this.btnUnselectAll.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnUnselectAll.RightToLeft")));
			this.btnUnselectAll.Size = ((System.Drawing.Size)(resources.GetObject("btnUnselectAll.Size")));
			this.btnUnselectAll.TabIndex = ((int)(resources.GetObject("btnUnselectAll.TabIndex")));
			this.btnUnselectAll.Text = resources.GetString("btnUnselectAll.Text");
			this.btnUnselectAll.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnUnselectAll.TextAlign")));
			this.toolTip1.SetToolTip(this.btnUnselectAll, resources.GetString("btnUnselectAll.ToolTip"));
			this.btnUnselectAll.Visible = ((bool)(resources.GetObject("btnUnselectAll.Visible")));
			this.btnUnselectAll.Click += new System.EventHandler(this.btnUnselectAll_Click);
			// 
			// btnSelectAll
			// 
			this.btnSelectAll.AccessibleDescription = resources.GetString("btnSelectAll.AccessibleDescription");
			this.btnSelectAll.AccessibleName = resources.GetString("btnSelectAll.AccessibleName");
			this.btnSelectAll.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnSelectAll.Anchor")));
			this.btnSelectAll.BackColor = System.Drawing.SystemColors.Control;
			this.btnSelectAll.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnSelectAll.BackgroundImage")));
			this.btnSelectAll.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnSelectAll.Dock")));
			this.btnSelectAll.Enabled = ((bool)(resources.GetObject("btnSelectAll.Enabled")));
			this.btnSelectAll.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnSelectAll.FlatStyle")));
			this.btnSelectAll.Font = ((System.Drawing.Font)(resources.GetObject("btnSelectAll.Font")));
			this.btnSelectAll.ForeColor = System.Drawing.Color.Black;
			this.btnSelectAll.Image = ((System.Drawing.Image)(resources.GetObject("btnSelectAll.Image")));
			this.btnSelectAll.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnSelectAll.ImageAlign")));
			this.btnSelectAll.ImageIndex = ((int)(resources.GetObject("btnSelectAll.ImageIndex")));
			this.btnSelectAll.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnSelectAll.ImeMode")));
			this.btnSelectAll.Location = ((System.Drawing.Point)(resources.GetObject("btnSelectAll.Location")));
			this.btnSelectAll.Name = "btnSelectAll";
			this.btnSelectAll.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnSelectAll.RightToLeft")));
			this.btnSelectAll.Size = ((System.Drawing.Size)(resources.GetObject("btnSelectAll.Size")));
			this.btnSelectAll.TabIndex = ((int)(resources.GetObject("btnSelectAll.TabIndex")));
			this.btnSelectAll.Text = resources.GetString("btnSelectAll.Text");
			this.btnSelectAll.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnSelectAll.TextAlign")));
			this.toolTip1.SetToolTip(this.btnSelectAll, resources.GetString("btnSelectAll.ToolTip"));
			this.btnSelectAll.Visible = ((bool)(resources.GetObject("btnSelectAll.Visible")));
			this.btnSelectAll.Click += new System.EventHandler(this.btnSelectAll_Click);
			// 
			// btnScenarioPSiteDefault
			// 
			////this.btnScenarioPSiteDefault.AccessibleDescription = resources.GetString("btnScenarioPSiteDefault.AccessibleDescription");
			////this.btnScenarioPSiteDefault.AccessibleName = resources.GetString("btnScenarioPSiteDefault.AccessibleName");
			////this.btnScenarioPSiteDefault.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("btnScenarioPSiteDefault.Anchor")));
			////this.btnScenarioPSiteDefault.BackColor = System.Drawing.SystemColors.Control;
			////this.btnScenarioPSiteDefault.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnScenarioPSiteDefault.BackgroundImage")));
			////this.btnScenarioPSiteDefault.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("btnScenarioPSiteDefault.Dock")));
			////this.btnScenarioPSiteDefault.Enabled = ((bool)(resources.GetObject("btnScenarioPSiteDefault.Enabled")));
			////this.btnScenarioPSiteDefault.FlatStyle = ((System.Windows.Forms.FlatStyle)(resources.GetObject("btnScenarioPSiteDefault.FlatStyle")));
			////this.btnScenarioPSiteDefault.Font = ((System.Drawing.Font)(resources.GetObject("btnScenarioPSiteDefault.Font")));
			////this.btnScenarioPSiteDefault.ForeColor = System.Drawing.Color.Black;
			////this.btnScenarioPSiteDefault.Image = ((System.Drawing.Image)(resources.GetObject("btnScenarioPSiteDefault.Image")));
			////this.btnScenarioPSiteDefault.ImageAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnScenarioPSiteDefault.ImageAlign")));
			////this.btnScenarioPSiteDefault.ImageIndex = ((int)(resources.GetObject("btnScenarioPSiteDefault.ImageIndex")));
			////this.btnScenarioPSiteDefault.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("btnScenarioPSiteDefault.ImeMode")));
			////this.btnScenarioPSiteDefault.Location = ((System.Drawing.Point)(resources.GetObject("btnScenarioPSiteDefault.Location")));
			////this.btnScenarioPSiteDefault.Name = "btnScenarioPSiteDefault";
			////this.btnScenarioPSiteDefault.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("btnScenarioPSiteDefault.RightToLeft")));
			////this.btnScenarioPSiteDefault.Size = ((System.Drawing.Size)(resources.GetObject("btnScenarioPSiteDefault.Size")));
			////this.btnScenarioPSiteDefault.TabIndex = ((int)(resources.GetObject("btnScenarioPSiteDefault.TabIndex")));
			////this.btnScenarioPSiteDefault.Tag = "";
			////this.btnScenarioPSiteDefault.Text = resources.GetString("btnScenarioPSiteDefault.Text");
			////this.btnScenarioPSiteDefault.TextAlign = ((System.Drawing.ContentAlignment)(resources.GetObject("btnScenarioPSiteDefault.TextAlign")));
			////this.toolTip1.SetToolTip(this.btnScenarioPSiteDefault, resources.GetString("btnScenarioPSiteDefault.ToolTip"));
			////this.btnScenarioPSiteDefault.Visible = ((bool)(resources.GetObject("btnScenarioPSiteDefault.Visible")));
			////this.btnScenarioPSiteDefault.Click += new System.EventHandler(this.btnScenarioPSiteDefault_Click);
			// 
			// listView1
			// 
			this.listView1.AccessibleDescription = resources.GetString("listView1.AccessibleDescription");
			this.listView1.AccessibleName = resources.GetString("listView1.AccessibleName");
			this.listView1.Alignment = ((System.Windows.Forms.ListViewAlignment)(resources.GetObject("listView1.Alignment")));
			this.listView1.AllowColumnReorder = true;
			this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("listView1.Anchor")));
			this.listView1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("listView1.BackgroundImage")));
			this.listView1.CheckBoxes = true;
			this.listView1.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("listView1.Dock")));
			this.listView1.Enabled = ((bool)(resources.GetObject("listView1.Enabled")));
			this.listView1.Font = ((System.Drawing.Font)(resources.GetObject("listView1.Font")));
			this.listView1.GridLines = true;
			this.listView1.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("listView1.ImeMode")));
			this.listView1.LabelWrap = ((bool)(resources.GetObject("listView1.LabelWrap")));
			this.listView1.Location = ((System.Drawing.Point)(resources.GetObject("listView1.Location")));
			this.listView1.MultiSelect = false;
			this.listView1.Name = "listView1";
			this.listView1.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("listView1.RightToLeft")));
			this.listView1.Size = ((System.Drawing.Size)(resources.GetObject("listView1.Size")));
			this.listView1.TabIndex = ((int)(resources.GetObject("listView1.TabIndex")));
			this.listView1.Text = resources.GetString("listView1.Text");
			this.toolTip1.SetToolTip(this.listView1, resources.GetString("listView1.ToolTip"));
			this.listView1.View = System.Windows.Forms.View.Details;
			this.listView1.Visible = ((bool)(resources.GetObject("listView1.Visible")));
			// 
			// uc_scenario_psite
			// 
			this.AccessibleDescription = resources.GetString("$this.AccessibleDescription");
			this.AccessibleName = resources.GetString("$this.AccessibleName");
			this.AutoScroll = ((bool)(resources.GetObject("$this.AutoScroll")));
			this.AutoScrollMargin = ((System.Drawing.Size)(resources.GetObject("$this.AutoScrollMargin")));
			this.AutoScrollMinSize = ((System.Drawing.Size)(resources.GetObject("$this.AutoScrollMinSize")));
			this.BackColor = System.Drawing.SystemColors.Control;
			this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
			this.Controls.Add(this.groupBox1);
			this.Enabled = ((bool)(resources.GetObject("$this.Enabled")));
			this.Font = ((System.Drawing.Font)(resources.GetObject("$this.Font")));
			this.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("$this.ImeMode")));
			this.Location = ((System.Drawing.Point)(resources.GetObject("$this.Location")));
			this.Name = "uc_scenario_psite";
			this.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("$this.RightToLeft")));
			this.Size = ((System.Drawing.Size)(resources.GetObject("$this.Size")));
			this.toolTip1.SetToolTip(this, resources.GetString("$this.ToolTip"));
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion


		private void cmdPSites_Click(object sender, System.EventArgs e)
		{
		}

		private void grpboxPSites_Resize(object sender, System.EventArgs e)
		{
			this.lstPSites.Width = this.ClientSize.Width - (int)(this.lstPSites.Left * 2);
		}
		private void lstPSites_ColumnClick(object sender, System.Windows.Forms.ColumnClickEventArgs e)
		{
            int x, y;
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }
            // Perform the sort with these new sort options.
            this.lstPSites.Sort();
            //reinitialize the alternate row colors
            for (x = 0; x <= this.lstPSites.Items.Count - 1; x++)
            {
                for (y = 0; y <= this.lstPSites.Columns.Count - 1; y++)
                {
                    m_oLvRowColors.ListViewSubItem(this.lstPSites.Items[x].Index, y, this.lstPSites.Items[this.lstPSites.Items[x].Index].SubItems[y], false);
                }
            }
           
		}

		private void btnSelectAll_Click(object sender, System.EventArgs e)
		{
			//if (this.lstPSites.CheckedItems.Count<this.lstPSites.Items.Count)
			//{
			//	if (this.ReferenceCoreScenarioForm.btnSave.Enabled==false) 
			//		((frmScenario)this.ParentForm).btnSave.Enabled=true;
			//}
			for (int x=0;x<=this.lstPSites.Items.Count-1;x++)
			{
				this.lstPSites.Items[x].Checked=true;
			}
		}

		private void btnUnselectAll_Click(object sender, System.EventArgs e)
		{
			//if (this.lstPSites.CheckedItems.Count>0)
			//{
			//	if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//		((frmScenario)this.ParentForm).btnSave.Enabled=true;
			//}
			for (int x=0;x<=this.lstPSites.Items.Count-1;x++)
			{
				this.lstPSites.Items[x].Checked=false;
			}
		}
/// <summary>
/// reload the listview with values from the wood processing site table
/// </summary>
/// <param name="sender"></param>
/// <param name="e"></param>
		
		private void m_Combo_SelectedIndexChanged(object sender, EventArgs e)
		{
			//MessageBox.Show("SelectedIndexChanged");
//			if (((frmScenario)this.Parent).m_lrulesfirsttime==false)
//			{
//				if (this.ReferenceCoreScenarioForm.btnSave.Enabled==false) 
//					((frmScenario)this.ParentForm).btnSave.Enabled=true;
//			}
		}
		private void m_Combo_SelectedValueChanged(object sender, EventArgs e)
		{
			//MessageBox.Show("SelectedValueChanged");
		}
		private void lstPSites_ItemCheck(object sender, 
			System.Windows.Forms.ItemCheckEventArgs e)
		{
//			if (((frmScenario)this.Parent).m_lrulesfirsttime==false)
//			{
//				if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
//					((frmScenario)this.ParentForm).btnSave.Enabled=true;
//			}
		}
		private void lstPSites_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			int x;
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = lstPSites.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lstPSites.Items[lstPSites.TopItem.Index + (int)dblRow-1].Selected=true;
				}
			}
			catch 
			{
			}
		}
		private void lstPSites_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.lstPSites.SelectedItems.Count > 0)
				m_oLvRowColors.DelegateListViewItem(this.lstPSites.SelectedItems[0]);
		}
		

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			lstPSites.Width = this.ClientSize.Width - (lstPSites.Left * 2);
			//this.btnScenarioPSiteDefault.Top = this.ClientSize.Height - (int)(this.btnScenarioPSiteDefault.Height * 1.5);
			//this.btnScenarioPSiteUpdate.Top = this.btnScenarioPSiteDefault.Top;
			//this.btnSelectAll.Top = this.btnScenarioPSiteDefault.Top;
			//this.btnUnselectAll.Top = this.btnScenarioPSiteDefault.Top;
			//lstPSites.Height = this.ClientSize.Height - lstPSites.Top - (int)(this.btnScenarioPSiteDefault.Height * 1.5) - 3;
		}
	
		public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
		
	}

}

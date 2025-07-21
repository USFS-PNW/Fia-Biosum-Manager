using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Text;
using SQLite.ADO;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario_filter.
	/// </summary>
	public class uc_optimizer_scenario_filter : System.Windows.Forms.UserControl
	{
		private env m_oEnv;
		private System.Windows.Forms.ImageList imgSize;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.ComponentModel.IContainer components;
		public int m_intFullHt = 400;
		private string[] m_strOptimizerTables;
		private string m_strCurrentSQL;
		private string m_strCurrentYardDist;
		private string m_strCurrentYardDist2;
		public System.Windows.Forms.TextBox txtCurrentSQL;
		private System.Windows.Forms.Button btnCreateSQL;
		private System.Windows.Forms.Button btnExecuteSQL;
		private System.Windows.Forms.Button btnPrevSQL;
		private System.Windows.Forms.Button btnEditSQL;
		private System.Windows.Forms.TextBox txtYardDist;
		private System.Windows.Forms.GroupBox grpboxSQL;
		private System.Windows.Forms.GroupBox grpboxYardDist;
		private System.Windows.Forms.Label lblLowSlope;
		private System.Windows.Forms.Label lblHighSlope;
		private System.Windows.Forms.TextBox txtYardDist2;
        private System.Windows.Forms.Button btnYardDistDefault;
		private int m_intNumberOfOptimizerTablesLoadedIntoDatasets;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Button btnValidate;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario=null;
		private string _strFilterType="PLOT";
		private string m_strScenarioTable="scenario_plot_filter";
		public int m_intError=0;
		public string m_strError="";
		private string _strSteepYardingDistanceDefault="2500";
        private string _strLowYardingDistanceDefault = "2500";


		public uc_optimizer_scenario_filter()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_oEnv = new env();
			this.grpboxSQL.Left = 4;
			this.grpboxYardDist.Left = 4;
			this.txtCurrentSQL.Left = 4;
			this.m_strOptimizerTables = new string[50];
			this.m_strCurrentSQL = "";
			this.m_strCurrentYardDist="";
			this.m_strCurrentYardDist2="";
			this.m_intNumberOfOptimizerTablesLoadedIntoDatasets=0;

			
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_optimizer_scenario_filter));
            this.imgSize = new System.Windows.Forms.ImageList(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.grpboxYardDist = new System.Windows.Forms.GroupBox();
            this.btnYardDistDefault = new System.Windows.Forms.Button();
            this.lblHighSlope = new System.Windows.Forms.Label();
            this.txtYardDist2 = new System.Windows.Forms.TextBox();
            this.lblLowSlope = new System.Windows.Forms.Label();
            this.txtYardDist = new System.Windows.Forms.TextBox();
            this.grpboxSQL = new System.Windows.Forms.GroupBox();
            this.btnValidate = new System.Windows.Forms.Button();
            this.txtCurrentSQL = new System.Windows.Forms.TextBox();
            this.btnPrevSQL = new System.Windows.Forms.Button();
            this.btnCreateSQL = new System.Windows.Forms.Button();
            this.btnEditSQL = new System.Windows.Forms.Button();
            this.btnExecuteSQL = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.grpboxYardDist.SuspendLayout();
            this.grpboxSQL.SuspendLayout();
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
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.grpboxYardDist);
            this.groupBox1.Controls.Add(this.grpboxSQL);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(656, 408);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Paint += new System.Windows.Forms.PaintEventHandler(this.groupBox1_Paint);
            this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(650, 32);
            this.lblTitle.TabIndex = 29;
            this.lblTitle.Text = "Plot Filter";
            // 
            // grpboxYardDist
            // 
            this.grpboxYardDist.Controls.Add(this.btnYardDistDefault);
            this.grpboxYardDist.Controls.Add(this.lblHighSlope);
            this.grpboxYardDist.Controls.Add(this.txtYardDist2);
            this.grpboxYardDist.Controls.Add(this.lblLowSlope);
            this.grpboxYardDist.Controls.Add(this.txtYardDist);
            this.grpboxYardDist.ForeColor = System.Drawing.Color.Black;
            this.grpboxYardDist.Location = new System.Drawing.Point(24, 51);
            this.grpboxYardDist.Name = "grpboxYardDist";
            this.grpboxYardDist.Size = new System.Drawing.Size(417, 93);
            this.grpboxYardDist.TabIndex = 12;
            this.grpboxYardDist.TabStop = false;
            this.grpboxYardDist.Text = "Maximum Yarding Distance Allowed (Feet)";
            // 
            // btnYardDistDefault
            // 
            this.btnYardDistDefault.BackColor = System.Drawing.SystemColors.Control;
            this.btnYardDistDefault.ForeColor = System.Drawing.Color.Black;
            this.btnYardDistDefault.Location = new System.Drawing.Point(8, 58);
            this.btnYardDistDefault.Name = "btnYardDistDefault";
            this.btnYardDistDefault.Size = new System.Drawing.Size(104, 24);
            this.btnYardDistDefault.TabIndex = 17;
            this.btnYardDistDefault.Text = "Default Values";
            this.btnYardDistDefault.UseVisualStyleBackColor = false;
            this.btnYardDistDefault.Click += new System.EventHandler(this.btnYardDistDefault_Click);
            // 
            // lblHighSlope
            // 
            this.lblHighSlope.Location = new System.Drawing.Point(145, 16);
            this.lblHighSlope.Name = "lblHighSlope";
            this.lblHighSlope.Size = new System.Drawing.Size(134, 16);
            this.lblHighSlope.TabIndex = 16;
            this.lblHighSlope.Text = "Steep Slope";
            // 
            // txtYardDist2
            // 
            this.txtYardDist2.Location = new System.Drawing.Point(148, 32);
            this.txtYardDist2.Name = "txtYardDist2";
            this.txtYardDist2.Size = new System.Drawing.Size(104, 20);
            this.txtYardDist2.TabIndex = 15;
            this.txtYardDist2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYardDist2_KeyPress);
            // 
            // lblLowSlope
            // 
            this.lblLowSlope.Location = new System.Drawing.Point(12, 16);
            this.lblLowSlope.Name = "lblLowSlope";
            this.lblLowSlope.Size = new System.Drawing.Size(116, 16);
            this.lblLowSlope.TabIndex = 14;
            this.lblLowSlope.Text = "Low Slope";
            // 
            // txtYardDist
            // 
            this.txtYardDist.Location = new System.Drawing.Point(8, 32);
            this.txtYardDist.Name = "txtYardDist";
            this.txtYardDist.Size = new System.Drawing.Size(104, 20);
            this.txtYardDist.TabIndex = 13;
            this.txtYardDist.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYardDist_KeyPress);
            // 
            // grpboxSQL
            // 
            this.grpboxSQL.Controls.Add(this.btnValidate);
            this.grpboxSQL.Controls.Add(this.txtCurrentSQL);
            this.grpboxSQL.Controls.Add(this.btnPrevSQL);
            this.grpboxSQL.Controls.Add(this.btnCreateSQL);
            this.grpboxSQL.Controls.Add(this.btnEditSQL);
            this.grpboxSQL.Controls.Add(this.btnExecuteSQL);
            this.grpboxSQL.ForeColor = System.Drawing.Color.Black;
            this.grpboxSQL.Location = new System.Drawing.Point(24, 150);
            this.grpboxSQL.Name = "grpboxSQL";
            this.grpboxSQL.Size = new System.Drawing.Size(600, 225);
            this.grpboxSQL.TabIndex = 11;
            this.grpboxSQL.TabStop = false;
            this.grpboxSQL.Text = "Attribute Filter";
            // 
            // btnValidate
            // 
            this.btnValidate.BackColor = System.Drawing.SystemColors.Control;
            this.btnValidate.ForeColor = System.Drawing.Color.Black;
            this.btnValidate.Location = new System.Drawing.Point(224, 188);
            this.btnValidate.Name = "btnValidate";
            this.btnValidate.Size = new System.Drawing.Size(120, 24);
            this.btnValidate.TabIndex = 11;
            this.btnValidate.Text = "Validate";
            this.btnValidate.UseVisualStyleBackColor = false;
            this.btnValidate.Click += new System.EventHandler(this.btnValidate_Click);
            // 
            // txtCurrentSQL
            // 
            this.txtCurrentSQL.Location = new System.Drawing.Point(16, 24);
            this.txtCurrentSQL.Multiline = true;
            this.txtCurrentSQL.Name = "txtCurrentSQL";
            this.txtCurrentSQL.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtCurrentSQL.Size = new System.Drawing.Size(568, 152);
            this.txtCurrentSQL.TabIndex = 10;
            this.txtCurrentSQL.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCurrentSQL_KeyPress);
            // 
            // btnPrevSQL
            // 
            this.btnPrevSQL.BackColor = System.Drawing.SystemColors.Control;
            this.btnPrevSQL.ForeColor = System.Drawing.Color.Black;
            this.btnPrevSQL.Location = new System.Drawing.Point(344, 188);
            this.btnPrevSQL.Name = "btnPrevSQL";
            this.btnPrevSQL.Size = new System.Drawing.Size(120, 24);
            this.btnPrevSQL.TabIndex = 6;
            this.btnPrevSQL.Text = "View Previous SQL";
            this.btnPrevSQL.UseVisualStyleBackColor = false;
            this.btnPrevSQL.Click += new System.EventHandler(this.btnPrevSQL_Click);
            // 
            // btnCreateSQL
            // 
            this.btnCreateSQL.BackColor = System.Drawing.SystemColors.Control;
            this.btnCreateSQL.ForeColor = System.Drawing.Color.Black;
            this.btnCreateSQL.Location = new System.Drawing.Point(464, 188);
            this.btnCreateSQL.Name = "btnCreateSQL";
            this.btnCreateSQL.Size = new System.Drawing.Size(120, 24);
            this.btnCreateSQL.TabIndex = 8;
            this.btnCreateSQL.Text = "Create New SQL";
            this.btnCreateSQL.UseVisualStyleBackColor = false;
            this.btnCreateSQL.Click += new System.EventHandler(this.btnCreateSQL_Click);
            // 
            // btnEditSQL
            // 
            this.btnEditSQL.BackColor = System.Drawing.SystemColors.Control;
            this.btnEditSQL.ForeColor = System.Drawing.Color.Black;
            this.btnEditSQL.Location = new System.Drawing.Point(104, 188);
            this.btnEditSQL.Name = "btnEditSQL";
            this.btnEditSQL.Size = new System.Drawing.Size(120, 24);
            this.btnEditSQL.TabIndex = 5;
            this.btnEditSQL.Text = "Edit Current SQL";
            this.btnEditSQL.UseVisualStyleBackColor = false;
            this.btnEditSQL.Click += new System.EventHandler(this.btnEditSQL_Click);
            // 
            // btnExecuteSQL
            // 
            this.btnExecuteSQL.BackColor = System.Drawing.SystemColors.Control;
            this.btnExecuteSQL.ForeColor = System.Drawing.Color.Black;
            this.btnExecuteSQL.Location = new System.Drawing.Point(16, 188);
            this.btnExecuteSQL.Name = "btnExecuteSQL";
            this.btnExecuteSQL.Size = new System.Drawing.Size(88, 24);
            this.btnExecuteSQL.TabIndex = 7;
            this.btnExecuteSQL.Text = "Execute SQL";
            this.btnExecuteSQL.UseVisualStyleBackColor = false;
            this.btnExecuteSQL.Click += new System.EventHandler(this.btnExecuteSQL_Click);
            // 
            // uc_optimizer_scenario_filter
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_optimizer_scenario_filter";
            this.Size = new System.Drawing.Size(656, 408);
            this.groupBox1.ResumeLayout(false);
            this.grpboxYardDist.ResumeLayout(false);
            this.grpboxYardDist.PerformLayout();
            this.grpboxSQL.ResumeLayout(false);
            this.grpboxSQL.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void groupBox1_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{

		}

		private void grpboxFilter_Resize(object sender, System.EventArgs e)
		{
			try
			{
			}
			catch
			{
			}
		}
		public void loadvalues(bool p_bScenarioCopy)
		{
            this.txtCurrentSQL.Text = "";
            if (this.FilterType == "PLOT")
            {
                if (this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.PlotTableSQLFilter.Trim().Length == 0)
                {
                    this.txtCurrentSQL.Text = "SELECT @@PlotTable@@.* FROM @@PlotTable@@ ";
                }
                else
                {
                    this.txtCurrentSQL.Text = this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.PlotTableSQLFilter;
                }
                this.grpboxYardDist.Hide();
                this.grpboxSQL.Text = "Plot Attribute Filter";
            }
            else
            {
                this.txtYardDist.Text = "";
                this.txtYardDist2.Text = "";
                if (this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oCondTableSQLFilter.SQL.Trim().Length == 0)
                {
                    this.txtCurrentSQL.Text = "SELECT * FROM @@CondTable@@";
                }
                else
                {
                    this.txtCurrentSQL.Text = this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oCondTableSQLFilter.SQL;
                }
                this.txtYardDist.Text = ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oCondTableSQLFilter.LowSlopeMaximumYardingDistanceFeet;
                this.txtYardDist2.Text = ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oCondTableSQLFilter.SteepSlopeMaximumYardingDistanceFeet;
                if (p_bScenarioCopy == false)
                {
                    this.m_strCurrentYardDist = this.txtYardDist.Text;
                    this.m_strCurrentYardDist2 = this.txtYardDist2.Text;
                }
                this.grpboxSQL.Text = "Condition Attribute Filter";
            }
            if (p_bScenarioCopy==false)
                this.m_strCurrentSQL = this.txtCurrentSQL.Text;

            if (p_bScenarioCopy==false)
                ((frmOptimizerScenario)this.ParentForm).m_bSave = false;
		}
		
		public int savevalues()
		{
			int x = 0;

			string str = "";
			string strNewSQL = "";
			System.Text.StringBuilder sb = new System.Text.StringBuilder();

			if (this.FilterType == "PLOT")
			{
				if (this.m_strCurrentSQL.Trim().ToUpper() == this.txtCurrentSQL.Text.Trim().ToUpper())
					return 0;

				m_strScenarioTable = "scenario_plot_filter";
			}
			else
			{
				if (this.m_strCurrentSQL.Trim().ToUpper() == this.txtCurrentSQL.Text.Trim().ToUpper() &&
					this.m_strCurrentYardDist.Trim() == this.txtYardDist.Text.Trim() &&
					this.m_strCurrentYardDist2.Trim() == this.txtYardDist2.Text.Trim())
					return 0;
				m_strScenarioTable = "scenario_cond_filter";
			}

			DataMgr oDataMgr = new DataMgr();
			string strScenarioId = this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.ScenarioId.Trim();
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

				if (this.m_strCurrentSQL.Trim().ToUpper() != this.txtCurrentSQL.Text.Trim().ToUpper())
				{
					strNewSQL = oDataMgr.FixString(this.txtCurrentSQL.Text, "'", "''");

					//delete any duplicates
					oDataMgr.m_strSQL = "DELETE FROM " + this.m_strScenarioTable + " WHERE " +
						" TRIM(UPPER(scenario_id))='" + strScenarioId.Trim().ToUpper() + "' AND " +
						"TRIM(UPPER(sql_command)) = '" + strNewSQL.Trim().ToUpper() + "'";
					oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

					oDataMgr.m_strSQL = "SELECT * FROM " + this.m_strScenarioTable + " WHERE " +
						" TRIM(UPPER(scenario_id)) = '" + strScenarioId.Trim().ToUpper() + "' AND current_yn = 'Y';";
					oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);

					if (oDataMgr.m_intError == 0)
					{
						//update the current sql as NOT being current
						if (oDataMgr.m_DataReader.HasRows)
						{
							oDataMgr.m_DataReader.Close();
							oDataMgr.m_strSQL = "UPDATE " + this.m_strScenarioTable + " SET current_yn = 'N'" +
								" WHERE TRIM(UPPER(scenario_id)) = '" + strScenarioId.Trim().ToUpper() + "' AND current_yn = 'Y';";
							oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
						}
						else
						{
							oDataMgr.m_DataReader.Close();
						}

						for (x = 0; x <= this.m_intNumberOfOptimizerTablesLoadedIntoDatasets - 1; x++)
						{
							if (m_strOptimizerTables[x].Trim().Length > 0)
								str = str + m_strOptimizerTables[x] + ",";
						}
						//remove the last comma
						if (str.Trim().Length > 0) str = str.Substring(0, str.Length - 1);

						//strNewSQL = p_ado.FixString(this.txtCurrentSQL.Text,"'","''");
						sb.Append("INSERT INTO " + this.m_strScenarioTable + " (scenario_id,sql_command,current_yn,table_list) Values('");
						sb.Append(strScenarioId);
						sb.Append("','");
						sb.Append(strNewSQL);
						sb.Append("','Y','");
						sb.Append(str);
						sb.Append("');");
						oDataMgr.SqlNonQuery(conn, sb.ToString());
					}
				}

				if (this.FilterType == "COND")
				{
					if (this.m_strCurrentYardDist.Trim() != this.txtYardDist.Text.Trim() ||
						this.m_strCurrentYardDist2.Trim() != this.txtYardDist2.Text.Trim())
					{
						//delete all records from the scenario wind speed class table
						oDataMgr.m_strSQL = "DELETE FROM scenario_cond_filter_misc WHERE " +
							" TRIM(UPPER(scenario_id)) = '" + strScenarioId.Trim().ToUpper() + "';";

						oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
						if (oDataMgr.m_intError < 0)
						{
							conn.Close();
							x = oDataMgr.m_intError;
							oDataMgr = null;
							return x;
						}
						if (this.txtYardDist.Text.Trim().Length == 0)
						{
							this.txtYardDist.Text = "1524";
						}
						if (this.txtYardDist2.Text.Trim().Length == 0)
						{
							this.txtYardDist2.Text = "610";
						}

						oDataMgr.m_strSQL = "INSERT INTO scenario_cond_filter_misc (scenario_id,yard_dist,yard_dist2)" +
							" VALUES ('" + strScenarioId.Trim() + "'," +
							this.txtYardDist.Text.Trim() + "," + this.txtYardDist2.Text.Trim() + ");";
						oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
						if (oDataMgr.m_intError < 0)
						{
							x = oDataMgr.m_intError;
							oDataMgr = null;
							return x;
						}
					}

				}
			}
			return 0;
		}
		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.grpboxSQL.Width = this.groupBox1.Width - 8;
				this.txtCurrentSQL.Width = this.grpboxSQL.Width - 8;
				this.btnEditSQL.Left = (int)(this.grpboxSQL.Width * .50) - (int)(this.btnEditSQL.Width*2); //* .50);
				this.btnExecuteSQL.Left = this.btnEditSQL.Left - this.btnExecuteSQL.Width;
				this.btnValidate.Left = this.btnEditSQL.Left + this.btnValidate.Width;
				this.btnPrevSQL.Left = this.btnValidate.Left + this.btnEditSQL.Width;
				this.btnCreateSQL.Left = this.btnPrevSQL.Left + this.btnPrevSQL.Width;
				this.grpboxSQL.Height = this.groupBox1.Height - this.grpboxSQL.Top;
				this.btnEditSQL.Top = this.grpboxSQL.Height - this.btnEditSQL.Height - 3;
				this.btnExecuteSQL.Top = this.btnEditSQL.Top;
				this.btnPrevSQL.Top = this.btnEditSQL.Top;
				this.btnCreateSQL.Top= this.btnEditSQL.Top;
				this.btnValidate.Top = this.btnEditSQL.Top;
				this.txtCurrentSQL.Height = this.grpboxSQL.Height - this.txtCurrentSQL.Top - this.btnEditSQL.Height - 10;
			}
			catch
			{
			}
		}


		private void btnCreateSQL_Click(object sender, System.EventArgs e)
        {
			int x = 0;
			int y = 0;
			DialogResult result;

			x = this.ReferenceOptimizerScenarioForm.uc_datasource1.getDataSourceTableNameRow("PLOT");
			if (x < 0)
			{
				MessageBox.Show("!!Plot table cannot be found!!");
				return;
			}

			frmSqlBuilder frmSQL = new frmSqlBuilder(ReferenceOptimizerScenarioForm.uc_datasource1.lstRequiredTables, "");

			if (this.FilterType == "PLOT") frmSQL.ClientId = "OPTIMIZER SCENARIO PLOT FILTER";
			else frmSQL.ClientId = "OPTIMIZER SCENARIO COND FILTER";

			frmSQL.ReferenceOptimizerScenarioForm = this.ReferenceOptimizerScenarioForm;
			frmSQL.txtSQLCommand.Text = this.txtCurrentSQL.Text;

			result = frmSQL.ShowDialog(this);

			if (result == DialogResult.OK)
			{
				if (frmSQL.txtSQLCommand.Text.Trim().Length > 0)
				{
					result = MessageBox.Show("REPLACE \n" + "----------------\n\n" + "'" + this.txtCurrentSQL.Text + "' \n\n\n WITH  \n----------------\n\n'" + frmSQL.txtSQLCommand.Text + "'", "Plot Filter", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
					if (result == DialogResult.Yes)
					{
						for (x = 0; x <= ReferenceOptimizerScenarioForm.uc_datasource1.lstRequiredTables.Items.Count - 1; x++)
						{
							m_strOptimizerTables[x] = "";
						}

						for (x = 0; x <= frmSQL.lstTables.Items.Count - 1; x++)
						{
							m_strOptimizerTables[x] = frmSQL.lstTables.Items[x].ToString();
						}
						this.txtCurrentSQL.Text = frmSQL.txtSQLCommand.Text;
						((frmOptimizerScenario)this.ParentForm).m_bSave = true;
					}
				}
			}

			frmSQL.Close();
			frmSQL = null;
		}

		/****************************************************************
		 ** edit the sql in sql builder form
		 ****************************************************************/
		private void btnEditSQL_Click(object sender, System.EventArgs e)
		{
			int x=0;
			int y=0;
			DialogResult result;
            
			x = this.ReferenceOptimizerScenarioForm.uc_datasource1.getDataSourceTableNameRow("PLOT");
			if (x < 0) 
			{
				MessageBox.Show("!!Plot table cannot be found!!");
				return;
			}

			frmSqlBuilder frmSQL = new frmSqlBuilder(ReferenceOptimizerScenarioForm.uc_datasource1.lstRequiredTables,"");
			
		    if (this.FilterType=="PLOT") frmSQL.ClientId="OPTIMIZER SCENARIO PLOT FILTER";
			else frmSQL.ClientId = "OPTIMIZER SCENARIO COND FILTER";
			frmSQL.ReferenceOptimizerScenarioForm=this.ReferenceOptimizerScenarioForm;
			frmSQL.txtSQLCommand.Text = this.txtCurrentSQL.Text;

			result = frmSQL.ShowDialog(this);
			//frmSQL.uc_sql_builder2.frmGridView1.Close();
			if (result == DialogResult.OK)
			{
				if (frmSQL.txtSQLCommand.Text.Trim().Length > 0)
				{
					result = MessageBox.Show("REPLACE \n" + "----------------\n\n" + "'" + this.txtCurrentSQL.Text + "' \n\n\n WITH  \n----------------\n\n'" + frmSQL.txtSQLCommand.Text + "'", "Plot Filter",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
					if (result == DialogResult.Yes)
					{
						for(x=0;x <= ReferenceOptimizerScenarioForm.uc_datasource1.lstRequiredTables.Items.Count-1;x++)
						{
							 m_strOptimizerTables[x] = "";
						}

						for(x=0;x <= frmSQL.lstTables.Items.Count-1;x++)
						{
						   m_strOptimizerTables[x] = frmSQL.lstTables.Items[x].ToString();
						}
						this.txtCurrentSQL.Text = frmSQL.txtSQLCommand.Text;
						((frmOptimizerScenario)this.ParentForm).m_bSave=true;
					}
				}
			}
			frmSQL.Close();
			frmSQL = null;

		}


		private void btnPrevSQL_Click(object sender, System.EventArgs e)
        {
			DataMgr p_dataMgr = new DataMgr();

			DialogResult result;

			string strScenarioDB = ((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\" +
				Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;

			frmDialog frmPrevExp = new frmDialog();

			frmPrevExp.Width = frmPrevExp.uc_previous_expressions1.m_intFullWd;
			frmPrevExp.Height = frmPrevExp.uc_previous_expressions1.m_intFullHt;
			frmPrevExp.Text = "Treatment Optimizer: Previous SQL Expressions";

			frmPrevExp.uc_previous_expressions1.Visible = true;

			string strConn = p_dataMgr.GetConnectionString(strScenarioDB);

			if (this.FilterType == "PLOT") m_strScenarioTable = "scenario_plot_filter";
			else m_strScenarioTable = "scenario_cond_filter";

			frmPrevExp.uc_previous_expressions1.loadvalues(strConn, "SELECT * FROM " + this.m_strScenarioTable + ";", "SQL_COMMAND", "SQL_COMMAND", this.m_strScenarioTable);

			result = frmPrevExp.ShowDialog(this);
			if (result == DialogResult.OK)
			{
				result = MessageBox.Show("REPLACE \n" + "----------------\n\n" + "'" + this.txtCurrentSQL.Text + "' \n\n\n WITH  \n----------------\n\n'" + frmPrevExp.uc_previous_expressions1.lblSQL.Text + "'", "Previous SQL", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
				if (result == DialogResult.Yes)
				{
					this.txtCurrentSQL.Text = frmPrevExp.uc_previous_expressions1.listView1.SelectedItems[0].SubItems[frmPrevExp.uc_previous_expressions1.m_intSelectColumn + 1].Text;
					((frmOptimizerScenario)this.ParentForm).m_bSave = true;
				}
			}
			frmPrevExp.Close();
			frmPrevExp = null;
		}

		private void btnExecuteSQL_Click(object sender, System.EventArgs e)
        {
			DataMgr oDataMgr = new DataMgr();
			utils p_utils = new utils();
			macrosubst oMacroSubst = new macrosubst();
			oMacroSubst.ReferenceSQLMacroSubstitutionVariableCollection = frmMain.g_oSQLMacroSubstitutionVariable_Collection;
			FIA_Biosum_Manager.Datasource oDs = new Datasource();
			oDs.m_strDataSourceMDBFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
				Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
			oDs.m_strDataSourceTableName = "scenario_datasource";
			oDs.m_strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text;
			oDs.LoadTableColumnNamesAndDataTypes = false;
			oDs.LoadTableRecordCount = false;

			oDs.populate_datasource_array_sqlite_new();
			System.Collections.Generic.List<string> lstDataSourceDbs = oDs.getDataSourceDbsList();

			if (oDs.m_intError == 0)
			{
				string strTempDb = p_utils.getRandomFile(m_oEnv.strTempDir, "db");
				oDataMgr.CreateDbFile(strTempDb);

				// attach databases from datasource array
				using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strTempDb)))
				{
					conn.Open();

					for (int x = 0; x <= lstDataSourceDbs.Count - 1; x++)
					{
						if (!oDataMgr.DatabaseAttached(conn, lstDataSourceDbs[x]))
						{
							oDataMgr.m_strSQL = "ATTACH DATABASE '" + lstDataSourceDbs[x] + "' AS alias" + x.ToString();
							oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
						}
					}

					string strSQL = this.txtCurrentSQL.Text.Trim();
					if (strSQL.IndexOf("@@", 0) >= 0) strSQL = oMacroSubst.SQLTranslateVariableSubstitution(strSQL);
					((frmOptimizerScenario)this.ParentForm).frmGridView1.LoadDataSet(conn, strSQL);

					if (((frmOptimizerScenario)this.ParentForm).frmGridView1.Visible == false)
					{

						((frmOptimizerScenario)this.ParentForm).frmGridView1.MdiParent = this.ParentForm.ParentForm;
						((frmOptimizerScenario)this.ParentForm).frmGridView1.Text = "Treatment Optimizer: Plot Filter Browser" + " (" + ((frmOptimizerScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text + ")";

						try
						{
							((frmOptimizerScenario)this.ParentForm).frmGridView1.Show();
						}
						catch
						{
							((frmOptimizerScenario)this.ParentForm).frmGridView1 = new frmGridView();
							((frmOptimizerScenario)this.ParentForm).frmGridView1.MdiParent = this.ParentForm.ParentForm;
							((frmOptimizerScenario)this.ParentForm).frmGridView1.Visible = false;
							((frmOptimizerScenario)this.ParentForm).frmGridView1.LoadDataSet(conn, strSQL);
							((frmOptimizerScenario)this.ParentForm).frmGridView1.MdiParent = this.ParentForm.ParentForm;
							((frmOptimizerScenario)this.ParentForm).frmGridView1.Text = "Treatment Optimizer: Plot Filter Browser" + " (" + ((frmOptimizerScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text + ")";

							((frmOptimizerScenario)this.ParentForm).frmGridView1.Show();
						}
					}
					if (((frmOptimizerScenario)this.ParentForm).frmGridView1.WindowState ==
							System.Windows.Forms.FormWindowState.Minimized)
						((frmOptimizerScenario)this.ParentForm).frmGridView1.WindowState =
							System.Windows.Forms.FormWindowState.Normal;
					((frmOptimizerScenario)this.ParentForm).frmGridView1.Focus();
				}
			}
		}

		private void txtCurrentSQL_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void txtYardDist_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
				((frmOptimizerScenario)this.ParentForm).m_bSave=true;
			}
			else
			{
				if (e.KeyChar == '\b')
				{
					((frmOptimizerScenario)this.ParentForm).m_bSave=true;
				}
				else
				{
					e.Handled=true;
				}
			}
		}

		private void btnYardDistDefault_Click(object sender, System.EventArgs e)
		{
            this.txtYardDist.Text = this._strLowYardingDistanceDefault;
			this.txtYardDist2.Text = this._strSteepYardingDistanceDefault;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void txtYardDist2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
				((frmOptimizerScenario)this.ParentForm).m_bSave=true;
			}
			else
			{
				if (e.KeyChar == '\b')
				{
					((frmOptimizerScenario)this.ParentForm).m_bSave=true;
				}
				else
				{
					e.Handled=true;
				}
			}
		}

		private void cmdFilter_Click(object sender, System.EventArgs e)
		{
		}
		
		public int Val_PlotFilter(System.Data.SQLite.SQLiteConnection p_oConn, string p_strSql)
        {
			m_intError = 0;
			m_strError = "";
			FIA_Biosum_Manager.macrosubst oVarSub = new macrosubst();
			oVarSub.ReferenceSQLMacroSubstitutionVariableCollection = frmMain.g_oSQLMacroSubstitutionVariable_Collection;
			DataMgr oDataMgr = new DataMgr();

			//make sure no duplicates of biosum_plot_id
			oDataMgr.m_strSQL = "SELECT b.plotcount, a.biosum_plot_id " +
				"FROM @@PlotTable@@ AS a, " +
				"(SELECT COUNT(*) AS plotcount, biosum_plot_id " +
				"FROM ( " + p_strSql + ") " +
				"GROUP BY biosum_plot_id) AS b " +
				"WHERE b.biosum_plot_id = a.biosum_plot_id AND b.plotcount > 1";
			oDataMgr.m_strSQL = oVarSub.SQLTranslateVariableSubstitution(oDataMgr.m_strSQL);

			if ((int)oDataMgr.getRecordCount(p_oConn, oDataMgr.m_strSQL, "testforplotduplicates") > 0)
			{
				m_intError = -5;
				m_strError = "Invalid Sql: Duplicate biosum_plot_id's were detected. Biosum_plot_id must be unique in the Treatment Optimizer Plot Filter";
				return -5;
			}
			else if (oDataMgr.m_intError != 0)
			{
				return oDataMgr.m_intError;
			}
            else
            {
				return 0;
            }
		}
		
		public int Val_PlotFilter(string p_strSql)
        {
			int intError = 0;
			if (p_strSql.ToUpper().Contains("PLOT_ACCESSIBLE_YN"))
			{
				intError = -1;
				m_strError = "The Plot filter contains the 'plot_accessible_YN' field. " +
							 "This field no longer exists on the plot table and must be " +
							 "removed from the filter !!";
				return intError;
			}
			DataMgr oDataMgr = new DataMgr();
			utils p_utils = new utils();
			FIA_Biosum_Manager.Datasource oDs = new Datasource();
			oDs.m_strDataSourceMDBFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
				Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
			oDs.m_strDataSourceTableName = "scenario_datasource";
			oDs.m_strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text;
			oDs.LoadTableColumnNamesAndDataTypes = false;
			oDs.LoadTableRecordCount = false;
			oDs.populate_datasource_array_sqlite_new();

			System.Collections.Generic.List<string> lstDataSourceDbs = oDs.getDataSourceDbsList();

			if (oDs.m_intError == 0)
            {
				string strTempDb = p_utils.getRandomFile(m_oEnv.strTempDir, "db");
				oDataMgr.CreateDbFile(strTempDb);

				// attach databases from datasource array
				using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strTempDb)))
                {
					conn.Open();

					for (int x = 0; x <= lstDataSourceDbs.Count - 1; x++)
					{
						if (!oDataMgr.DatabaseAttached(conn, lstDataSourceDbs[x]))
						{
							oDataMgr.m_strSQL = "ATTACH DATABASE '" + lstDataSourceDbs[x] + "' AS alias" + x.ToString();
							oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
						}
					}

					intError = Val_PlotFilter(conn, p_strSql.Trim());
					if (intError != 0 && String.IsNullOrEmpty(m_strError))
					{
						m_strError = "An error occurred while validating the Plot filter!";
					}
				}
			}
			oDataMgr = null;
			return intError;
		}

		
		public int Val_CondFilter(System.Data.SQLite.SQLiteConnection p_oConn, string p_strSql)
        {

			m_intError = 0;
			m_strError = "";
			FIA_Biosum_Manager.macrosubst oVarSub = new macrosubst();
			oVarSub.ReferenceSQLMacroSubstitutionVariableCollection = frmMain.g_oSQLMacroSubstitutionVariable_Collection;
			DataMgr oDataMgr = new DataMgr();

			//make sure no duplicates of biosum_plot_id
			oDataMgr.m_strSQL = "SELECT b.condcount,a.biosum_cond_id FROM @@CondTable@@ AS a," +
				"(SELECT COUNT(*) as condcount, biosum_cond_id " +
				"FROM (" + p_strSql + ") " +
				"GROUP BY biosum_cond_id) AS b " +
				"WHERE b.biosum_cond_id=a.biosum_cond_id AND b.condcount > 1";
			oDataMgr.m_strSQL = oVarSub.SQLTranslateVariableSubstitution(oDataMgr.m_strSQL);

			if (oDataMgr.getRecordCount(p_oConn, oDataMgr.m_strSQL, "testforcondduplicates") > 0)
			{
				m_intError = -6;
				m_strError = "Invalid Sql: Duplicate biosum_cond_id's were detected. Biosum_cond_id must be unique in the Treatment Optimizer Scenario Cond Filter";
				return -6;
			}
			else if (oDataMgr.m_intError != 0)
			{
				return oDataMgr.m_intError;
			}
			else
			{
				if (this.txtYardDist.Text.Trim().Length == 0 || this.txtYardDist2.Text.Trim().Length == 0)
				{
					m_intError = -7;
					m_strError = "No value defined for yarding distance in the <Filter Cond Records> tab";
					return -7;

				}
				return 0;
			}
		}

		
		public int Val_CondFilter(string p_strSql)
        {
			DataMgr oDataMgr = new DataMgr();
			utils p_utils = new utils();
			FIA_Biosum_Manager.Datasource oDs = new Datasource();
			oDs.m_strDataSourceMDBFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
				Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
			oDs.m_strDataSourceTableName = "scenario_datasource";
			oDs.m_strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text;
			oDs.LoadTableColumnNamesAndDataTypes = false;
			oDs.LoadTableRecordCount = false;
			oDs.populate_datasource_array_sqlite_new();

			System.Collections.Generic.List<string> lstDataSourceDbs = oDs.getDataSourceDbsList();

			int intError = 0;
			if (oDs.m_intError == 0)
			{
				string strTempDb = p_utils.getRandomFile(m_oEnv.strTempDir, "db");
				oDataMgr.CreateDbFile(strTempDb);

				using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strTempDb)))
				{
					conn.Open();

					for (int x = 0; x <= lstDataSourceDbs.Count - 1; x++)
                    {
						if (!oDataMgr.DatabaseAttached(conn, lstDataSourceDbs[x]))
                        {
							oDataMgr.m_strSQL = "ATTACH DATABASE '" + lstDataSourceDbs[x] + "' AS alias" + x.ToString(); ;
							oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                        }
                    }

					intError = Val_CondFilter(conn, p_strSql.Trim());
					if (intError != 0 && String.IsNullOrEmpty(m_strError))
					{
						m_strError = "An error occurred while validating the Condition filter!";
					}
				}
			}
			oDataMgr = null;
			return intError;
		}

		private void btnValidate_Click(object sender, System.EventArgs e)
		{
			int intError=0;

			if (this.FilterType=="PLOT") intError = Val_PlotFilter(this.txtCurrentSQL.Text);
			else intError = Val_CondFilter(this.txtCurrentSQL.Text);

			if (intError!=0)
			{
				MessageBox.Show(m_strError,"FIA Biosum");
			}
			else if (intError==0)
			{
				MessageBox.Show("Valid Syntax");
			}
			else
			{
				MessageBox.Show("Error");
			}
		}
	
		public string strNonSteepYardingDistance
		{
			set	{ this.txtYardDist.Text  = value; }
			get { return this.txtYardDist.Text.ToString(); }
		}
		public string strSteepYardingDistance
		{
			set	{ this.txtYardDist2.Text  = value; }
			get { return this.txtYardDist2.Text.ToString(); }
		}
		public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
		public string FilterType
		{
			get {return this._strFilterType;}
			set {this._strFilterType=value;}
		}
        public string strSteepSlope
        {
            set { this.lblHighSlope.Text = "Steep Slope > " + value; }
        }
        public string strLowSlope
        {
            set { this.lblLowSlope.Text = "Low Slope <= " + value; }
        }

		
	}

}

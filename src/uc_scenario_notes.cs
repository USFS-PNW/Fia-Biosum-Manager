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
	/// Summary description for uc_scenario_notes.
	/// </summary>
	public class uc_scenario_notes : System.Windows.Forms.UserControl
	{
		public System.Windows.Forms.TextBox txtNotes;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario=null;
		private FIA_Biosum_Manager.frmProcessorScenario _frmProcessorScenario=null;
		private string _strScenarioType="optimizer";
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_scenario_notes()
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
			this.txtNotes = new System.Windows.Forms.TextBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblTitle = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// txtNotes
			// 
			this.txtNotes.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNotes.Location = new System.Drawing.Point(16, 56);
			this.txtNotes.Multiline = true;
			this.txtNotes.Name = "txtNotes";
			this.txtNotes.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtNotes.Size = new System.Drawing.Size(632, 344);
			this.txtNotes.TabIndex = 0;
			this.txtNotes.Text = "";
			this.txtNotes.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotes_KeyPress);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Controls.Add(this.txtNotes);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(664, 424);
			this.groupBox1.TabIndex = 1;
			this.groupBox1.TabStop = false;
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(658, 32);
			this.lblTitle.TabIndex = 26;
			this.lblTitle.Text = "Scenario Notes";
			// 
			// uc_scenario_notes
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_scenario_notes";
			this.Size = new System.Drawing.Size(664, 424);
			this.Resize += new System.EventHandler(this.uc_scenario_notes_Resize);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void uc_scenario_notes_Resize(object sender, System.EventArgs e)
		{
			try
			{
                this.txtNotes.Left = 5;
                this.txtNotes.Width = this.Width - 10;
                this.txtNotes.Height = this.ClientSize.Height - this.txtNotes.Top - 5;
                //this.txtNotes.Left = (int) (this.groupBox1.Width * .50) - (int) (this.txtNotes.Width * .50);
            }
			catch
			{
			}
			
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.Visible=false;
			if (this.ScenarioType.Trim().ToUpper() == "OPTIMIZER") ((frmOptimizerScenario)this.ParentForm).Height = 0 ; //((frmScenario)this.ParentForm).grpboxMenu.Height * 2;
			else this.ReferenceProcessorScenarioForm.Height=0;
		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
			this.SaveScenarioNotes();
		}
		public void SaveScenarioNotes()
		{
            ado_data_access p_ado = null;
            SQLite.ADO.DataMgr oDataMgr = null;
            string strNotes = this.txtNotes.Text;
            string strProjDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
            string strScenarioDir = strProjDir + "\\" + ScenarioType + "\\db";
            if ((ReferenceProcessorScenarioForm != null && !ReferenceProcessorScenarioForm.m_bUsingSqlite) 
                 || ScenarioType.Trim().ToUpper() == "OPTIMIZER")
            {
                p_ado = new ado_data_access();
                strNotes = p_ado.FixString(strNotes, "'", "''");
                string strSQL = "";
                if (ScenarioType.Trim().ToUpper() == "OPTIMIZER")
                {
                    strSQL = "UPDATE scenario SET notes = '" +
                        strNotes +
                        "' WHERE trim(lcase(scenario_id)) = '" + ((frmOptimizerScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text.Trim().ToLower() + "';";
                }
                else
                {
                    strSQL = "UPDATE scenario SET notes = '" +
                        strNotes +
                        "' WHERE trim(lcase(scenario_id)) = '" + this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower() + "';";
                }
                System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
                string strFile = "scenario_" + ScenarioType + "_rule_definitions.mdb";
                StringBuilder strFullPath = new StringBuilder(strScenarioDir);
                strFullPath.Append("\\");
                strFullPath.Append(strFile);
                string strConn = p_ado.getMDBConnString(strFullPath.ToString(), "admin", "");
                p_ado.SqlNonQuery(strConn, strSQL);
                p_ado = null;
            }
            else
            {
                oDataMgr = new SQLite.ADO.DataMgr();
                strNotes = oDataMgr.FixString(strNotes, "'", "''");
                //@ToDo: Only support Processor at this time
                string strSQL = "UPDATE scenario SET notes = '" +
                    strNotes +
                    "' WHERE trim(lower(scenario_id)) = '" + this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower() + "';";
                string strFile = "scenario_" + ScenarioType + "_rule_definitions.db";
                StringBuilder strFullPath = new StringBuilder(strScenarioDir);
                strFullPath.Append("\\");
                strFullPath.Append(strFile);
                string strConn = oDataMgr.GetConnectionString(strFullPath.ToString());
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    oDataMgr.SqlNonQuery(conn, strSQL);
                }
                oDataMgr = null;
            }		
		}
		public void LoadValues()
		{
			string strNotes="";
			FIA_Biosum_Manager.ado_data_access oAdo = new ado_data_access();
			string strScenarioDBDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\db";
			string strScenarioFile = "scenario_" + ScenarioType + "_rule_definitions.mdb"; 
			StringBuilder strScenarioFullPath = new StringBuilder(strScenarioDBDir);
			strScenarioFullPath.Append("\\");
			strScenarioFullPath.Append(strScenarioFile);
			string strScenarioConn = oAdo.getMDBConnString(strScenarioFullPath.ToString(),"admin","");
			oAdo.OpenConnection(strScenarioConn);
			if (ScenarioType.Trim().ToUpper()=="OPTIMIZER")
			{
				oAdo.m_strSQL = "SELECT notes FROM scenario WHERE TRIM(scenario_id)='" + this.ReferenceOptimizerScenarioForm.uc_scenario1.strScenarioId.Trim() + "'";
			}
			else
			{
				oAdo.m_strSQL = "SELECT notes FROM scenario WHERE TRIM(scenario_id)='" + this.ReferenceProcessorScenarioForm.uc_scenario1.strScenarioId.Trim() + "'";
			}
			strNotes=oAdo.getSingleStringValueFromSQLQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL,"scenario");
			oAdo.m_OleDbConnection.Close();
			while (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
			{
				oAdo.m_OleDbConnection.Close();
				System.Threading.Thread.Sleep(1000);
			}
			oAdo=null;
			this.txtNotes.Text = strNotes;
		}

        public void LoadValuesSqlite()
        {
            string strNotes = "";
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strScenarioDBDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\db";
            string strScenarioFile = "scenario_" + ScenarioType + "_rule_definitions.db";
            StringBuilder strScenarioFullPath = new StringBuilder(strScenarioDBDir);
            strScenarioFullPath.Append("\\");
            strScenarioFullPath.Append(strScenarioFile);
            string strScenarioConn = dataMgr.GetConnectionString(strScenarioFullPath.ToString());
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strScenarioConn))
            {
                con.Open();
                if (ScenarioType.Trim().ToUpper() == "OPTIMIZER")
                {
                    dataMgr.m_strSQL = "SELECT notes FROM scenario WHERE TRIM(scenario_id)='" + this.ReferenceOptimizerScenarioForm.uc_scenario1.strScenarioId.Trim() + "'";
                }
                else
                {
                    dataMgr.m_strSQL = "SELECT notes FROM scenario WHERE TRIM(scenario_id)='" + this.ReferenceProcessorScenarioForm.uc_scenario1.strScenarioId.Trim() + "'";
                }
                strNotes = dataMgr.getSingleStringValueFromSQLQuery(con, dataMgr.m_strSQL, "scenario");
            }
            this.txtNotes.Text = strNotes;
        }

        public void loadvalues_FromProperties()
        {
            if (ScenarioType.Trim().ToUpper() == "OPTIMIZER")
            {
                FIA_Biosum_Manager.OptimizerScenarioItem oItem = ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem;
                this.txtNotes.Text = oItem.Notes.Trim();
            }
            else
            {
                if (!String.IsNullOrEmpty(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.Notes))
                {
                    FIA_Biosum_Manager.ProcessorScenarioItem oItem = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem;
                    this.txtNotes.Text = oItem.Notes.Trim();
                }
            }


        }


		private void txtNotes_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (ScenarioType.Trim().ToUpper()=="OPTIMIZER") ReferenceOptimizerScenarioForm.m_bSave=true;
			else ReferenceProcessorScenarioForm.m_bSave=true;
		}
		public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
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
			get {return _strScenarioType;}
			set {_strScenarioType=value;}
		}
	}
}

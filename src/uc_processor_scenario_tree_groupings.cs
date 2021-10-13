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
	/// Summary description for uc_scenario_tree_groupings.
	/// </summary>
	public class uc_processor_scenario_tree_groupings : System.Windows.Forms.UserControl
    {
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmProcessorScenario _frmProcessorScenario=null;
		private string _strScenarioType="processor";
        private Button BtnTreeDiameterGroups;
        private FIA_Biosum_Manager.frmDialog m_frmTreeDiamGroups;     //processor tree diameter form
        private Button BtnTreeSpeciesGroups;       
        private FIA_Biosum_Manager.frmDialog m_frmTreeSpeciesGroups;     //processor tree species groups form

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

        public uc_processor_scenario_tree_groupings()
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.BtnTreeSpeciesGroups = new System.Windows.Forms.Button();
            this.BtnTreeDiameterGroups = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.BtnTreeSpeciesGroups);
            this.groupBox1.Controls.Add(this.BtnTreeDiameterGroups);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(664, 424);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            // 
            // BtnTreeSpeciesGroups
            // 
            this.BtnTreeSpeciesGroups.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnTreeSpeciesGroups.Location = new System.Drawing.Point(185, 160);
            this.BtnTreeSpeciesGroups.Name = "BtnTreeSpeciesGroups";
            this.BtnTreeSpeciesGroups.Size = new System.Drawing.Size(198, 33);
            this.BtnTreeSpeciesGroups.TabIndex = 28;
            this.BtnTreeSpeciesGroups.Text = "TREE SPECIES GROUPS";
            this.BtnTreeSpeciesGroups.UseVisualStyleBackColor = true;
            this.BtnTreeSpeciesGroups.Click += new System.EventHandler(this.BtnTreeSpeciesGroups_Click);
            // 
            // BtnTreeDiameterGroups
            // 
            this.BtnTreeDiameterGroups.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnTreeDiameterGroups.Location = new System.Drawing.Point(185, 99);
            this.BtnTreeDiameterGroups.Name = "BtnTreeDiameterGroups";
            this.BtnTreeDiameterGroups.Size = new System.Drawing.Size(198, 33);
            this.BtnTreeDiameterGroups.TabIndex = 27;
            this.BtnTreeDiameterGroups.Text = "TREE DIAMETER GROUPS";
            this.BtnTreeDiameterGroups.UseVisualStyleBackColor = true;
            this.BtnTreeDiameterGroups.Click += new System.EventHandler(this.BtnTreeDiameterGroups_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(658, 32);
            this.lblTitle.TabIndex = 26;
            this.lblTitle.Text = "Tree Groupings";
            // 
            // uc_processor_scenario_tree_groupings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_scenario_tree_groupings";
            this.Size = new System.Drawing.Size(664, 424);
            this.Resize += new System.EventHandler(this.uc_scenario_tree_groupings_Resize);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void uc_scenario_tree_groupings_Resize(object sender, System.EventArgs e)
		{
			
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.Visible=false;
			if (this.ScenarioType.Trim().ToUpper() == "CORE") ((frmOptimizerScenario)this.ParentForm).Height = 0 ; //((frmScenario)this.ParentForm).grpboxMenu.Height * 2;
			else this.ReferenceProcessorScenarioForm.Height=0;
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
        public FIA_Biosum_Manager.uc_processor_scenario_tree_diam_groups_list uc_processor_scenario_tree_diam_groups_list1
        {
            get 
            { 
                return m_frmTreeDiamGroups.uc_processor_scenario_tree_diam_groups_list1;  
            }
        }
        public FIA_Biosum_Manager.uc_processor_scenario_tree_spc_groups uc_processor_scenario_tree_spc_groups1
        {
            get
            {
                return m_frmTreeSpeciesGroups.uc_processor_scenario_tree_spc_groups1;
            }
        }

        private void BtnTreeDiameterGroups_Click(object sender, EventArgs e)
        {

            frmMain.g_oFrmMain.ActivateStandByAnimation(
                frmMain.g_oFrmMain.WindowState,
                frmMain.g_oFrmMain.Left,
                frmMain.g_oFrmMain.Height,
                frmMain.g_oFrmMain.Width,
                frmMain.g_oFrmMain.Top);

            // Initialize Tree Diameter form
            this.m_frmTreeDiamGroups = new frmDialog(_frmProcessorScenario, frmMain.g_oFrmMain);
            this.m_frmTreeDiamGroups.MaximizeBox = false;
            this.m_frmTreeDiamGroups.BackColor = System.Drawing.SystemColors.Control;
            this.m_frmTreeDiamGroups.Text = "Processor: Tree Diameter Groups";
            // @ToDo: Not sure if we need this
            //this.m_frmTreeDiam.MdiParent = ofrmProcessorScenario;
            this.m_frmTreeDiamGroups.Initialize_Plot_Tree_Diam_User_Control();

            this.m_frmTreeDiamGroups.Height = 0;
            this.m_frmTreeDiamGroups.Width = 0;    
            
            if (this.m_frmTreeDiamGroups.uc_processor_scenario_tree_diam_groups_list1.Top + this.m_frmTreeDiamGroups.uc_processor_scenario_tree_diam_groups_list1.Height > this.m_frmTreeDiamGroups.ClientSize.Height + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmTreeDiamGroups.Height = x;
                        if (this.m_frmTreeDiamGroups.uc_processor_scenario_tree_diam_groups_list1.Top +
                            this.m_frmTreeDiamGroups.uc_processor_scenario_tree_diam_groups_list1.Height <
                            this.m_frmTreeDiamGroups.ClientSize.Height)
                        {
                            break;
                        }
                    }

                }
                if (this.m_frmTreeDiamGroups.uc_processor_scenario_tree_diam_groups_list1.Left + this.m_frmTreeDiamGroups.uc_processor_scenario_tree_diam_groups_list1.Width > this.m_frmTreeDiamGroups.ClientSize.Width + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmTreeDiamGroups.Width = x;
                        if (this.m_frmTreeDiamGroups.uc_processor_scenario_tree_diam_groups_list1.Left +
                            this.m_frmTreeDiamGroups.uc_processor_scenario_tree_diam_groups_list1.Width <
                            this.m_frmTreeDiamGroups.ClientSize.Width)
                        {
                            break;
                        }
                    }

                }

            if (ReferenceProcessorScenarioForm.m_bTreeGroupsFirstTime == true &&
                ReferenceProcessorScenarioForm.m_bTreeGroupsCopied == false)
            {
                ReferenceProcessorScenarioForm.LoadTreeGroupings();
            }
            m_frmTreeDiamGroups.uc_processor_scenario_tree_diam_groups_list1.loadvalues();

            frmMain.g_oFrmMain.DeactivateStandByAnimation();
                
            this.m_frmTreeDiamGroups.Show();
            this.m_frmTreeDiamGroups.Left = 0;
            this.m_frmTreeDiamGroups.Top = 0;
            this.m_frmTreeDiamGroups.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
        }

        private void BtnTreeSpeciesGroups_Click(object sender, EventArgs e)
        {

            frmMain.g_oFrmMain.ActivateStandByAnimation(
                frmMain.g_oFrmMain.WindowState,
                frmMain.g_oFrmMain.Left,
                frmMain.g_oFrmMain.Height,
                frmMain.g_oFrmMain.Width,
                frmMain.g_oFrmMain.Top);

            // Initialize Tree Species Group form
            this.m_frmTreeSpeciesGroups = new frmDialog(_frmProcessorScenario, frmMain.g_oFrmMain);
            this.m_frmTreeSpeciesGroups.MaximizeBox = false;
            this.m_frmTreeSpeciesGroups.BackColor = System.Drawing.SystemColors.Control;
            this.m_frmTreeSpeciesGroups.Text = "Processor: Tree Species Groups";
            this.m_frmTreeSpeciesGroups.Initialize_Processor_Tree_Species_Groups();
            this.m_frmTreeSpeciesGroups.Height = 0;
            this.m_frmTreeSpeciesGroups.Width = 0;
            if (this.m_frmTreeSpeciesGroups.uc_processor_scenario_tree_spc_groups1.Top + m_frmTreeSpeciesGroups.uc_processor_scenario_tree_spc_groups1.Height > this.m_frmTreeSpeciesGroups.ClientSize.Height + 2)
            {
                for (int x = 1; ; x++)
                {
                    this.m_frmTreeSpeciesGroups.Height = x;
                    if (m_frmTreeSpeciesGroups.uc_processor_scenario_tree_spc_groups1.Top +
                        m_frmTreeSpeciesGroups.uc_processor_scenario_tree_spc_groups1.Height <
                        this.m_frmTreeSpeciesGroups.ClientSize.Height)
                    {
                        break;
                    }
                }

            }
            if (m_frmTreeSpeciesGroups.uc_processor_scenario_tree_spc_groups1.Left + m_frmTreeSpeciesGroups.uc_processor_scenario_tree_spc_groups1.Width > this.m_frmTreeSpeciesGroups.ClientSize.Width + 2)
            {
                for (int x = 1; ; x++)
                {
                    this.m_frmTreeSpeciesGroups.Width = x;
                    if (m_frmTreeSpeciesGroups.uc_processor_scenario_tree_spc_groups1.Left +
                        m_frmTreeSpeciesGroups.uc_processor_scenario_tree_spc_groups1.Width <
                        this.m_frmTreeSpeciesGroups.ClientSize.Width)
                    {
                        break;
                    }
                }

            }

            if (ReferenceProcessorScenarioForm.m_bTreeGroupsFirstTime == true &&
                ReferenceProcessorScenarioForm.m_bTreeGroupsCopied == false)
            {
                ReferenceProcessorScenarioForm.LoadTreeGroupings();
            }
            m_frmTreeSpeciesGroups.uc_processor_scenario_tree_spc_groups1.loadvalues();

            frmMain.g_oFrmMain.DeactivateStandByAnimation();
            this.m_frmTreeSpeciesGroups.Left = 0;
            this.m_frmTreeSpeciesGroups.Top = 0;
            this.m_frmTreeSpeciesGroups.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            this.m_frmTreeSpeciesGroups.Show();
        }

        public void saveTreeGroupings_FromProperties()
        {
            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                ado_data_access _objAdo = new ado_data_access();
                string strDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsDbFile;
                _objAdo.OpenConnection(_objAdo.getMDBConnString(strDbFile, "", ""));

                if (_objAdo.m_intError == 0)
                {
                    string strScenarioId = this.ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.ScenarioId;
                    //delete the current records
                    _objAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName +
                            " WHERE TRIM(UCASE(scenario_id)) = '" + strScenarioId.Trim() + "'";
                    _objAdo.SqlNonQuery(_objAdo.m_OleDbConnection, _objAdo.m_strSQL);

                    // saving tree diameter groups
                    if (_objAdo.m_intError == 0)
                    {
                        string strMin;
                        string strMax;
                        string strDef;
                        string strId;
                        for (int x = 0; x <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Count - 1; x++)
                        {
                            FIA_Biosum_Manager.ProcessorScenarioItem.TreeDiamGroupsItem oItem =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Item(x);
                            strId = oItem.DiamGroup;
                            strMin = oItem.MinDiam;
                            strMax = oItem.MaxDiam;
                            strDef = oItem.DiamClass;

                            _objAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName + " " +
                                "(diam_group,diam_class,min_diam,max_diam,scenario_id) VALUES " +
                                "(" + strId + ",'" + strDef.Trim() + "'," +
                                strMin + "," + strMax + ",'" + strScenarioId.Trim() + "');";
                            _objAdo.SqlNonQuery(_objAdo.m_OleDbConnection, _objAdo.m_strSQL);
                            if (_objAdo.m_intError != 0)
                            {
                                break;
                            }
                        }
                    }
                    // saving tree species groups
                    if (_objAdo.m_intError == 0)
                    {
                        string strCommonName;
                        int intSpCd;
                        int intSpcGrp;
                        string strGrpLabel;
                        int x;

                        //delete all records from the tree species group table
                        _objAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName +
                                           " WHERE TRIM(UCASE(scenario_id))='" + strScenarioId.Trim().ToUpper() + " '";
                        _objAdo.SqlNonQuery(_objAdo.m_OleDbConnection, _objAdo.m_strSQL);
                        if (_objAdo.m_intError != 0) return;

                        //delete all records from the tree species group list table
                        _objAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName +
                                          " WHERE TRIM(UCASE(scenario_id))='" + strScenarioId.Trim().ToUpper() + " '";
                        _objAdo.SqlNonQuery(_objAdo.m_OleDbConnection, _objAdo.m_strSQL);

                        if (_objAdo.m_intError == 0)
                        {
                            for (x = 0; x <= this.ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Count - 1; x++)
                            {
                                FIA_Biosum_Manager.ProcessorScenarioItem.SpcGroupItem oItem =
                                    ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Item(x);
                                intSpcGrp = oItem.SpeciesGroup;
                                strGrpLabel = oItem.SpeciesGroupLabel;
                                _objAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName + " " +
                                                   "(SPECIES_GROUP,SPECIES_LABEL,SCENARIO_ID) VALUES " +
                                                   "(" + Convert.ToString(intSpcGrp).Trim() + ",'" + strGrpLabel.Trim() + "','" + strScenarioId.Trim() + "');";
                                _objAdo.SqlNonQuery(_objAdo.m_OleDbConnection, _objAdo.m_strSQL);
                            }
                        }
                        if (_objAdo.m_intError == 0)
                        {
                            for (x = 0; x <= this.ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Count - 1; x++)
                            {
                                FIA_Biosum_Manager.ProcessorScenarioItem.SpcGroupListItem oItem =
                                  ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Item(x);
                                intSpcGrp = oItem.SpeciesGroup;
                                strCommonName = oItem.CommonName;
                                strCommonName = _objAdo.FixString(strCommonName.Trim(), "'", "''");
                                intSpCd = oItem.SpeciesCode;

                                _objAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName + " " +
                                    "(SPECIES_GROUP,common_name,SCENARIO_ID,SPCD) VALUES " +
                                    "(" + Convert.ToString(intSpcGrp).Trim() + ",'" + strCommonName + "','" + strScenarioId.Trim() + "', " +
                                    intSpCd + " );";
                                _objAdo.SqlNonQuery(_objAdo.m_OleDbConnection, _objAdo.m_strSQL);
                            }
                        }
                    }

                    ReferenceProcessorScenarioForm.m_bTreeGroupsCopied = false;
                }
                _objAdo.CloseConnection(_objAdo.m_OleDbConnection);
                _objAdo = null;
            }
            else
            {
                saveTreeGroupings_FromPropertiesSqlite();
            }
        }
        public void saveTreeGroupings_FromPropertiesSqlite()
        {
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            string strDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDbFile)))
            {
                conn.Open();
                if (oDataMgr.m_intError ==0)
                {
                    string strScenarioId = this.ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.ScenarioId;
                    //delete the current records
                    oDataMgr.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName +
                            " WHERE TRIM(UPPER(scenario_id)) = '" + strScenarioId.ToUpper().Trim() + "'";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                    // saving tree diameter groups
                    if (oDataMgr.m_intError == 0)
                    {
                        string strMin;
                        string strMax;
                        string strDef;
                        string strId;
                        for (int x = 0; x <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Count - 1; x++)
                        {
                            ProcessorScenarioItem.TreeDiamGroupsItem oItem =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Item(x);
                            strId = oItem.DiamGroup;
                            strMin = oItem.MinDiam;
                            strMax = oItem.MaxDiam;
                            strDef = oItem.DiamClass;

                            oDataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName + " " +
                                "(diam_group,diam_class,min_diam,max_diam,scenario_id) VALUES " +
                                "(" + strId + ",'" + strDef.Trim() + "'," +
                                strMin + "," + strMax + ",'" + strScenarioId.Trim() + "');";
                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                            if (oDataMgr.m_intError != 0)
                            {
                                break;
                            }
                        }
                    }
                    // saving tree species groups
                    if (oDataMgr.m_intError == 0)
                    {
                        string strCommonName;
                        int intSpCd;
                        int intSpcGrp;
                        string strGrpLabel;
                        int x;

                        //delete all records from the tree species group table
                        oDataMgr.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName +
                                           " WHERE TRIM(UPPER(scenario_id))='" + strScenarioId.Trim().ToUpper() + "'";
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                        if (oDataMgr.m_intError != 0) return;

                        //delete all records from the tree species group list table
                        oDataMgr.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName +
                                          " WHERE TRIM(UPPER(scenario_id))='" + strScenarioId.Trim().ToUpper() + "'";
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                        if (oDataMgr.m_intError == 0)
                        {
                            for (x = 0; x <= this.ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Count - 1; x++)
                            {
                                FIA_Biosum_Manager.ProcessorScenarioItem.SpcGroupItem oItem =
                                    ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Item(x);
                                intSpcGrp = oItem.SpeciesGroup;
                                strGrpLabel = oItem.SpeciesGroupLabel;
                                oDataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName + " " +
                                                   "(SPECIES_GROUP,SPECIES_LABEL,SCENARIO_ID) VALUES " +
                                                   "(" + Convert.ToString(intSpcGrp).Trim() + ",'" + strGrpLabel.Trim() + "','" + strScenarioId.Trim() + "');";
                                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                            }
                        }
                        if (oDataMgr.m_intError == 0)
                        {
                            for (x = 0; x <= this.ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Count - 1; x++)
                            {
                                ProcessorScenarioItem.SpcGroupListItem oItem =
                                  ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Item(x);
                                intSpcGrp = oItem.SpeciesGroup;
                                strCommonName = oItem.CommonName;
                                strCommonName = oDataMgr.FixString(strCommonName.Trim(), "'", "''");
                                intSpCd = oItem.SpeciesCode;

                                oDataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName + " " +
                                    "(SPECIES_GROUP,common_name,SCENARIO_ID,SPCD) VALUES " +
                                    "(" + Convert.ToString(intSpcGrp).Trim() + ",'" + strCommonName + "','" + strScenarioId.Trim() + "', " +
                                    intSpCd + " );";
                                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                            }
                        }
                    }
                    ReferenceProcessorScenarioForm.m_bTreeGroupsCopied = false;
                }

            }
            oDataMgr = null;
        }
    }
}

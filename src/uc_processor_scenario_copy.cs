using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    public partial class uc_processor_scenario_copy : UserControl
    {
        private FIA_Biosum_Manager.ProcessorScenarioItem_Collection m_oProcessorScenarioItem_Collection = new ProcessorScenarioItem_Collection();
        private FIA_Biosum_Manager.ProcessorScenarioItem m_oProcessorScenarioItem;
        private Queries m_oQueries = new Queries();
        private FIA_Biosum_Manager.ProcessorScenarioItem _oProcessorScenarioItem;
        private FIA_Biosum_Manager.frmProcessorScenario _frmProcessorScenario;
        private FIA_Biosum_Manager.ProcessorScenarioTools m_oProcessorScenarioTools = new ProcessorScenarioTools();
        private ListViewAlternateBackgroundColors m_oLvAlternateColors = new ListViewAlternateBackgroundColors();

        const int COL_CHECKBOX = 0;
        const int COL_SCENARIOID = 1;
        const int COL_DESC = 2;
        private bool m_bSuppressCheckEvents = false;
        
        FIA_Biosum_Manager.frmDialog _frmDialog = null;
        public uc_processor_scenario_copy()
        {
            InitializeComponent();
        }
        public FIA_Biosum_Manager.frmDialog ReferenceDialogForm
        {
            get { return _frmDialog; }
            set { _frmDialog = value; }
        }
        public FIA_Biosum_Manager.ProcessorScenarioItem ReferenceCurrentScenarioItem
        {
            get { return this._oProcessorScenarioItem; }
            set { this._oProcessorScenarioItem = value; }
        }
        public FIA_Biosum_Manager.frmProcessorScenario ReferenceProcessorScenarioForm
        {
            get { return _frmProcessorScenario; }
            set { _frmProcessorScenario = value; }
        }
        public void loadvalues()
        {
            lvProcessorScenario.Items.Clear();
            this.m_oLvAlternateColors.InitializeRowCollection();
            this.m_oLvAlternateColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceListView = this.lvProcessorScenario;
            this.m_oLvAlternateColors.CustomFullRowSelect = true;
            if (frmMain.g_oGridViewFont != null) this.lvProcessorScenario.Font = frmMain.g_oGridViewFont;

            if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
            {
                loadgrid();
            }
            else
            {
                loadgridSqlite();
            }
        }
        private void loadgrid()
        {
            //
            //OPEN CONNECTION TO DB FILE CONTAINING Processor Scenario TABLE
            //
            //scenario mdb connection
            string strProcessorScenarioMDB =
              frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
              "\\processor\\db\\scenario_processor_rule_definitions.mdb";
            //
            //get a list of all the scenarios
            //
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(strProcessorScenarioMDB, "", ""));
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection,
                       "SELECT scenario_id,description " +
                       "FROM scenario " +
                       "WHERE scenario_id IS NOT NULL AND " +
                                         "LEN(TRIM(scenario_id)) > 0");

            int x = 0;
            ListViewItem entryListItem = null;
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    if (oAdo.m_OleDbDataReader["scenario_id"] != DBNull.Value &&
                        oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim().Length > 0 &&
                        ReferenceCurrentScenarioItem.ScenarioId.Trim().ToUpper() !=
                        oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim().ToUpper())
                    {
                        entryListItem = lvProcessorScenario.Items.Add(" ");

                        entryListItem.UseItemStyleForSubItems = false;
                        this.m_oLvAlternateColors.AddRow();
                        this.m_oLvAlternateColors.AddColumns(x, lvProcessorScenario.Columns.Count);


                        entryListItem.SubItems.Add(oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim());

                        if (oAdo.m_OleDbDataReader["description"] != DBNull.Value &&
                            oAdo.m_OleDbDataReader["description"].ToString().Trim().Length > 0)
                        {
                            entryListItem.SubItems.Add(oAdo.m_OleDbDataReader["description"].ToString().Trim());
                        }
                        else
                        {
                            entryListItem.SubItems.Add(" ");
                        }
                        x = x + 1;
                    }
                }
                this.m_oLvAlternateColors.ListView();
            }
            else
            {
                MessageBox.Show("!!No Scenarios To Copy!!", "FIA Bisoum");
                btnCopy.Enabled = false;

            }
            oAdo.m_OleDbDataReader.Close();
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
        }
        private void loadgridSqlite()
        {
            //
            //OPEN CONNECTION TO DB FILE CONTAINING Processor Scenario TABLE
            //
            //scenario mdb connection
            string strProcessorScenarioDB =
              frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
              "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;

            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();            
            using (System.Data.SQLite.SQLiteConnection conn = 
                new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strProcessorScenarioDB)))
            {
                conn.Open();
                //
                //get a list of all the scenarios
                //
                oDataMgr.SqlQueryReader(conn,
                           "SELECT scenario_id,description " +
                           "FROM scenario " +
                           "WHERE scenario_id IS NOT NULL AND " +
                                             "LENGTH(TRIM(scenario_id)) > 0");
                int x = 0;
                ListViewItem entryListItem = null;
                if (oDataMgr.m_DataReader.HasRows)
                {
                    while (oDataMgr.m_DataReader.Read())

                    {
                        if (oDataMgr.m_DataReader["scenario_id"] != DBNull.Value &&
                            oDataMgr.m_DataReader["scenario_id"].ToString().Trim().Length > 0 &&
                            ReferenceCurrentScenarioItem.ScenarioId.Trim().ToUpper() !=
                            oDataMgr.m_DataReader["scenario_id"].ToString().Trim().ToUpper())
                        {
                            entryListItem = lvProcessorScenario.Items.Add(" ");

                            entryListItem.UseItemStyleForSubItems = false;
                            this.m_oLvAlternateColors.AddRow();
                            this.m_oLvAlternateColors.AddColumns(x, lvProcessorScenario.Columns.Count);


                            entryListItem.SubItems.Add(oDataMgr.m_DataReader["scenario_id"].ToString().Trim());

                            if (oDataMgr.m_DataReader["description"] != DBNull.Value &&
                                oDataMgr.m_DataReader["description"].ToString().Trim().Length > 0)
                            {
                                entryListItem.SubItems.Add(oDataMgr.m_DataReader["description"].ToString().Trim());
                            }
                            else
                            {
                                entryListItem.SubItems.Add(" ");
                            }
                            x = x + 1;
                        }
                    }
                    oDataMgr.m_DataReader.Close();
                    this.m_oLvAlternateColors.ListView();
                }
                else
                {
                    MessageBox.Show("!!No Scenarios To Copy!!", "FIA Bisoum");
                    btnCopy.Enabled = false;
                }
            }
        }

        private void lvProcessorScenario_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvProcessorScenario.SelectedItems.Count > 0)
            {
                m_oLvAlternateColors.DelegateListViewItem(lvProcessorScenario.SelectedItems[0]);
                if (chkFullDetails.Checked)
                {
                    FullDetails();
                }
                else
                {
                    txtDetails.Text = lvProcessorScenario.SelectedItems[0].SubItems[COL_DESC].Text;
                }
            }
        }

        private void lvProcessorScenario_MouseUp(object sender, MouseEventArgs e)
        {
            
            try
            {
                if (e.Button == MouseButtons.Left)
                {
                    int intRowHt = lvProcessorScenario.Items[0].Bounds.Height;
                    double dblRow = (double)(e.Y / intRowHt);
                    lvProcessorScenario.Items[lvProcessorScenario.TopItem.Index + (int)dblRow - 1].Selected = true;
                    this.m_oLvAlternateColors.DelegateListViewItem(lvProcessorScenario.Items[lvProcessorScenario.TopItem.Index + (int)dblRow - 1]);


                }
            }
            catch
            {
            }
        }

        private void lvProcessorScenario_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            
        }

        private void lvProcessorScenario_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (m_bSuppressCheckEvents == true) return;
            if (this.lvProcessorScenario.SelectedItems.Count == 0) return;
            m_bSuppressCheckEvents = true;
            for (int x = 0; x <= this.lvProcessorScenario.Items.Count - 1; x++)
            {
                if (this.lvProcessorScenario.Items[x].Index !=
                    this.lvProcessorScenario.SelectedItems[0].Index)
                {
                    lvProcessorScenario.Items[x].Checked = false;
                }
                else
                {
                    if (e.NewValue == System.Windows.Forms.CheckState.Checked)
                    {

                       

                    }
                }

            }
            m_bSuppressCheckEvents = false;

        }

        private void panel1_Resize(object sender, EventArgs e)
        {

            panel1_Resize();

        }
        public void panel1_Resize()
        {
            this.txtDetails.Width = this.panel1.Width - (int)(txtDetails.Left * 2);
            lblMsg.Width = this.txtDetails.Width;
            lblMsg.Left = txtDetails.Left;
            this.lvProcessorScenario.Width = txtDetails.Width;
            this.btnCopy.Top = this.panel1.ClientSize.Height - btnCopy.Height - 5;
            this.btnCancel.Top = btnCopy.Top;
            this.lblMsg.Top = btnCopy.Top - lblMsg.Height - 2;
            btnCopy.Left = (int)(txtDetails.Width * .5) - btnCopy.Width - 10;
            btnCancel.Left = (int)(txtDetails.Width * .5) + 10;
            this.txtDetails.Top = lblMsg.Top - txtDetails.Height - 2;
            this.chkFullDetails.Top = txtDetails.Top - chkFullDetails.Height - 2;
            this.lvProcessorScenario.Height = chkFullDetails.Top - lvProcessorScenario.Top;
        }

        private void chkFullDetails_CheckedChanged(object sender, EventArgs e)
        {
            if (chkFullDetails.Checked && lvProcessorScenario.SelectedItems.Count > 0)
                   FullDetails();
        }
        private void FullDetails()
        {
            int x=0;

            this.txtDetails.Text  = "";

            CheckIfScenarioLoaded(lvProcessorScenario.SelectedItems[0].SubItems[1].Text.Trim(),out x);
            
            this.m_oProcessorScenarioItem = m_oProcessorScenarioItem_Collection.Item(x);

            this.txtDetails.Text = m_oProcessorScenarioTools.ScenarioProperties(m_oProcessorScenarioItem);
        }
        public int val_processorscenario()
        {
            if (this.lvProcessorScenario.Items.Count == 0)
            {
                MessageBox.Show("Run Scenario Failed: No Processor Scenarios exist. At least one Processor Scenario must exist to run a Processor Scenario", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return -1;
            }
            if (this.lvProcessorScenario.CheckedItems.Count == 0)
            {
                MessageBox.Show("Run Scenario Failed: Select at least one Processor Scenario in <Processor Scenario>", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return -1;
            }
            this.m_oProcessorScenarioItem = this.m_oProcessorScenarioItem_Collection.Item(lvProcessorScenario.CheckedItems[0].Index);

            return 0;

        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            if (lvProcessorScenario.SelectedItems.Count == 0) return;
            int x = 0;
            string strMsg = "Copy scenario properties\r\n\r\nFROM\r\n-------\r\n" + lvProcessorScenario.SelectedItems[0].SubItems[1].Text.Trim() + "\r\n\r\nTO\r\n-------\r\n" + ReferenceCurrentScenarioItem.ScenarioId;
            DialogResult result = MessageBox.Show(strMsg, "FIA Biosum", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                CheckIfScenarioLoaded(lvProcessorScenario.SelectedItems[0].SubItems[1].Text.Trim(), out x);

                this.m_oProcessorScenarioItem = m_oProcessorScenarioItem_Collection.Item(x);

                ReferenceCurrentScenarioItem.Description = m_oProcessorScenarioItem.Description;
                ReferenceCurrentScenarioItem.Notes = m_oProcessorScenarioItem.Notes;
                ReferenceCurrentScenarioItem.SourceScenarioId = m_oProcessorScenarioItem.ScenarioId;
                
                ReferenceCurrentScenarioItem.m_oTreeDiamGroupsItem_Collection.Copy(m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection,
                    ref ReferenceCurrentScenarioItem.m_oTreeDiamGroupsItem_Collection, true);

                ReferenceCurrentScenarioItem.m_oEscalators.Copy(m_oProcessorScenarioItem.m_oEscalators, ReferenceCurrentScenarioItem.m_oEscalators);

                ReferenceCurrentScenarioItem.m_oHarvestCostItem_Collection.Copy(
                    m_oProcessorScenarioItem.m_oHarvestCostItem_Collection,
                ref ReferenceCurrentScenarioItem.m_oHarvestCostItem_Collection, true);

                ReferenceCurrentScenarioItem.m_oHarvestMethod.Copy(m_oProcessorScenarioItem.m_oHarvestMethod, ReferenceCurrentScenarioItem.m_oHarvestMethod);

                ReferenceCurrentScenarioItem.m_oMoveInCosts.Copy(m_oProcessorScenarioItem.m_oMoveInCosts, ReferenceCurrentScenarioItem.m_oMoveInCosts);

                ReferenceCurrentScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Copy(
                    m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection,
                ref ReferenceCurrentScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection, true);

                ReferenceCurrentScenarioItem.m_oTreeDiamGroupsItem_Collection.Copy(
                    m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection,
                ref ReferenceCurrentScenarioItem.m_oTreeDiamGroupsItem_Collection, true);

                ReferenceCurrentScenarioItem.m_oSpcGroupItem_Collection.Copy(
                    m_oProcessorScenarioItem.m_oSpcGroupItem_Collection,
                ref ReferenceCurrentScenarioItem.m_oSpcGroupItem_Collection, true);

                ReferenceCurrentScenarioItem.m_oSpcGroupListItem_Collection.Copy(
                    m_oProcessorScenarioItem.m_oSpcGroupListItem_Collection,
                ref ReferenceCurrentScenarioItem.m_oSpcGroupListItem_Collection, true);
               
                _frmDialog.DialogResult = DialogResult.OK;
                _frmDialog.Close();
                
            }
        }
        private void CheckIfScenarioLoaded(string p_strScenarioId,out int x)
        {
            
            //search to see if this scenario was loaded into the collection
            for (x = 0; x <= m_oProcessorScenarioItem_Collection.Count - 1; x++)
            {
                if (m_oProcessorScenarioItem_Collection.Item(x).ScenarioId.Trim().ToUpper() ==
                    p_strScenarioId.Trim().ToUpper()) break;
            }
            if (x > m_oProcessorScenarioItem_Collection.Count - 1)
            {

                lblMsg.Text = "Loading Processor Scenario " + p_strScenarioId.Trim() + "...Stand By";
                lblMsg.Show();
                lblMsg.Refresh();
                //load the scenario into the collection
                m_oProcessorScenarioTools.LoadScenario(p_strScenarioId.Trim(), m_oQueries, ReferenceProcessorScenarioForm.m_bUsingSqlite,
                    m_oProcessorScenarioItem_Collection);
                lblMsg.Hide();
            }
        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            _frmDialog.DialogResult = DialogResult.Cancel;
            _frmDialog.Close();
        }
        
    }
}

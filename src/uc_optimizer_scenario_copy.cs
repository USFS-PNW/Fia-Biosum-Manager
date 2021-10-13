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
    public partial class uc_optimizer_scenario_copy : UserControl
    {
        private FIA_Biosum_Manager.OptimizerScenarioItem_Collection m_oOptimizerScenarioItem_Collection = new OptimizerScenarioItem_Collection();
        private FIA_Biosum_Manager.OptimizerScenarioItem m_oOptimizerScenarioItem;
        private Queries m_oQueries = new Queries();
        private FIA_Biosum_Manager.OptimizerScenarioItem _oOptimizerScenarioItem;
        private FIA_Biosum_Manager.frmOptimizerScenario _frmOptimizerScenario;
        private FIA_Biosum_Manager.OptimizerScenarioTools m_oOptimizerScenarioTools = new OptimizerScenarioTools();
        private ListViewAlternateBackgroundColors m_oLvAlternateColors = new ListViewAlternateBackgroundColors();

        const int COL_CHECKBOX = 0;
        const int COL_SCENARIOID = 1;
        const int COL_DESC = 2;
        private bool m_bSuppressCheckEvents = false;
        
        FIA_Biosum_Manager.frmDialog _frmDialog = null;
        public uc_optimizer_scenario_copy()
        {
            InitializeComponent();
        }
        public FIA_Biosum_Manager.frmDialog ReferenceDialogForm
        {
            get { return _frmDialog; }
            set { _frmDialog = value; }
        }
        public FIA_Biosum_Manager.OptimizerScenarioItem ReferenceCurrentScenarioItem
        {
            get { return this._oOptimizerScenarioItem; }
            set { this._oOptimizerScenarioItem = value; }
        }
        public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
        {
            get { return _frmOptimizerScenario; }
            set { _frmOptimizerScenario = value; }
        }
        public void loadvalues()
        {
            int x;
            
            lvOptimizerScenario.Items.Clear();
            System.Windows.Forms.ListViewItem entryListItem = null;
            this.m_oLvAlternateColors.InitializeRowCollection();
            this.m_oLvAlternateColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceListView = this.lvOptimizerScenario;
            this.m_oLvAlternateColors.CustomFullRowSelect = true;
            if (frmMain.g_oGridViewFont != null) this.lvOptimizerScenario.Font = frmMain.g_oGridViewFont;
            //
            //OPEN CONNECTION TO DB FILE CONTAINING Optimizer Analysis Scenario TABLE
            //
            //scenario mdb connection
            string strOptimizerScenarioMDB =
              frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
              Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;
            //
            //get a list of all the scenarios
            //
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(strOptimizerScenarioMDB,"",""));
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection,
                       "SELECT scenario_id,description " +
                       "FROM scenario " +
                       "WHERE scenario_id IS NOT NULL AND " +
                                         "LEN(TRIM(scenario_id)) > 0");

            x = 0;
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    if (oAdo.m_OleDbDataReader["scenario_id"] != DBNull.Value &&
                        oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim().Length > 0 &&
                        ReferenceCurrentScenarioItem.ScenarioId.Trim().ToUpper() !=
                        oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim().ToUpper())
                    {
                        entryListItem = lvOptimizerScenario.Items.Add(" ");

                        entryListItem.UseItemStyleForSubItems = false;
                        this.m_oLvAlternateColors.AddRow();
                        this.m_oLvAlternateColors.AddColumns(x, lvOptimizerScenario.Columns.Count);


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
        public void savevalues()
        {
           
        }

        private void lvCoreAnalysisScenario_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvOptimizerScenario.SelectedItems.Count > 0)
            {
                m_oLvAlternateColors.DelegateListViewItem(lvOptimizerScenario.SelectedItems[0]);
                if (chkFullDetails.Checked)
                {
                    FullDetails();
                }
                else
                {
                    txtDetails.Text = lvOptimizerScenario.SelectedItems[0].SubItems[COL_DESC].Text;
                }
            }
        }

        private void lvCoreAnalysisScenario_MouseUp(object sender, MouseEventArgs e)
        {
            
            try
            {
                if (e.Button == MouseButtons.Left)
                {
                    int intRowHt = lvOptimizerScenario.Items[0].Bounds.Height;
                    double dblRow = (double)(e.Y / intRowHt);
                    lvOptimizerScenario.Items[lvOptimizerScenario.TopItem.Index + (int)dblRow - 1].Selected = true;
                    this.m_oLvAlternateColors.DelegateListViewItem(lvOptimizerScenario.Items[lvOptimizerScenario.TopItem.Index + (int)dblRow - 1]);


                }
            }
            catch
            {
            }
        }

        private void lvCoreAnalysisScenario_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            
        }

        private void lvCoreAnalysisScenario_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (m_bSuppressCheckEvents == true) return;
            if (this.lvOptimizerScenario.SelectedItems.Count == 0) return;
            m_bSuppressCheckEvents = true;
            for (int x = 0; x <= this.lvOptimizerScenario.Items.Count - 1; x++)
            {
                if (this.lvOptimizerScenario.Items[x].Index !=
                    this.lvOptimizerScenario.SelectedItems[0].Index)
                {
                    lvOptimizerScenario.Items[x].Checked = false;
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
            this.lvOptimizerScenario.Width = txtDetails.Width;
            this.btnCopy.Top = this.panel1.ClientSize.Height - btnCopy.Height - 5;
            this.btnCancel.Top = btnCopy.Top;
            this.lblMsg.Top = btnCopy.Top - lblMsg.Height - 2;
            btnCopy.Left = (int)(txtDetails.Width * .5) - btnCopy.Width - 10;
            btnCancel.Left = (int)(txtDetails.Width * .5) + 10;
            this.txtDetails.Top = lblMsg.Top - txtDetails.Height - 2;
            this.chkFullDetails.Top = txtDetails.Top - chkFullDetails.Height - 2;
            this.lvOptimizerScenario.Height = chkFullDetails.Top - lvOptimizerScenario.Top;
        }

        private void chkFullDetails_CheckedChanged(object sender, EventArgs e)
        {
            if (chkFullDetails.Checked && lvOptimizerScenario.SelectedItems.Count > 0)
                   FullDetails();
        }
        private void FullDetails()
        {
            int x=0;

            this.txtDetails.Text  = "";

            CheckIfScenarioLoaded(lvOptimizerScenario.SelectedItems[0].SubItems[1].Text.Trim(),out x);
            
            this.m_oOptimizerScenarioItem = m_oOptimizerScenarioItem_Collection.Item(x);

            this.txtDetails.Text = m_oOptimizerScenarioTools.ScenarioProperties(m_oOptimizerScenarioItem);
        }
        public int val_corescenario()
        {
            if (this.lvOptimizerScenario.Items.Count == 0)
            {
                MessageBox.Show("Run Scenario Failed: No Treatment Optimizer Scenarios exist. At least one Treatment Optimizer Scenario must exist to run a Treatment Optimizer Scenario", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return -1;
            }
            if (this.lvOptimizerScenario.CheckedItems.Count == 0)
            {
                MessageBox.Show("Run Scenario Failed: Select at least one Treatment Optimizer Scenario in <Cost and Revenue><Treatment Optimizer Scenario>", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return -1;
            }
            this.m_oOptimizerScenarioItem = this.m_oOptimizerScenarioItem_Collection.Item(lvOptimizerScenario.CheckedItems[0].Index);

            return 0;

        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            if (lvOptimizerScenario.SelectedItems.Count == 0) return;
            int x = 0;
            string strMsg = "Copy scenario properties\r\n\r\nFROM\r\n-------\r\n" + lvOptimizerScenario.SelectedItems[0].SubItems[1].Text.Trim() + "\r\n\r\nTO\r\n-------\r\n" + ReferenceCurrentScenarioItem.ScenarioId;
            DialogResult result = MessageBox.Show(strMsg, "FIA Biosum", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                CheckIfScenarioLoaded(lvOptimizerScenario.SelectedItems[0].SubItems[1].Text.Trim(), out x);

                this.m_oOptimizerScenarioItem = m_oOptimizerScenarioItem_Collection.Item(x);

                ReferenceCurrentScenarioItem.Description = m_oOptimizerScenarioItem.Description;
                ReferenceCurrentScenarioItem.Notes = m_oOptimizerScenarioItem.Notes;
                ReferenceCurrentScenarioItem.m_oCondTableSQLFilter = m_oOptimizerScenarioItem.m_oCondTableSQLFilter;
                ReferenceCurrentScenarioItem.m_oEffectiveVariablesItem_Collection.Copy(m_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection, ref ReferenceCurrentScenarioItem.m_oEffectiveVariablesItem_Collection, true);
                ReferenceCurrentScenarioItem.m_oOptimizationVariableItem_Collection.Copy(m_oOptimizerScenarioItem.m_oOptimizationVariableItem_Collection, ref ReferenceCurrentScenarioItem.m_oOptimizationVariableItem_Collection, true);
                ReferenceCurrentScenarioItem.m_oProcessingSiteItem_Collection.Copy(m_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection, ref ReferenceCurrentScenarioItem.m_oProcessingSiteItem_Collection, true);
                ReferenceCurrentScenarioItem.m_oLastTieBreakRankItem_Collection.Copy(m_oOptimizerScenarioItem.m_oLastTieBreakRankItem_Collection, ref ReferenceCurrentScenarioItem.m_oLastTieBreakRankItem_Collection, true);
                ReferenceCurrentScenarioItem.m_oTieBreaker_Collection.Copy(m_oOptimizerScenarioItem.m_oTieBreaker_Collection, ref ReferenceCurrentScenarioItem.m_oTieBreaker_Collection, true);
                ReferenceCurrentScenarioItem.m_oTranCosts.Copy(m_oOptimizerScenarioItem.m_oTranCosts, ReferenceCurrentScenarioItem.m_oTranCosts);
                ReferenceCurrentScenarioItem.m_oProcessorScenarioItem_Collection.Copy(m_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection, ref ReferenceCurrentScenarioItem.m_oProcessorScenarioItem_Collection, true);
                ReferenceCurrentScenarioItem.OwnerGroupCodeList = m_oOptimizerScenarioItem.OwnerGroupCodeList;
                ReferenceCurrentScenarioItem.PlotTableSQLFilter = m_oOptimizerScenarioItem.PlotTableSQLFilter;
               
                _frmDialog.DialogResult = DialogResult.OK;
                _frmDialog.Close();
                
            }
        }
        private void CheckIfScenarioLoaded(string p_strScenarioId,out int x)
        {
            
            //search to see if this scenario was loaded into the collection
            for (x = 0; x <= m_oOptimizerScenarioItem_Collection.Count - 1; x++)
            {
                if (m_oOptimizerScenarioItem_Collection.Item(x).ScenarioId.Trim().ToUpper() ==
                    p_strScenarioId.Trim().ToUpper()) break;
            }
            if (x > m_oOptimizerScenarioItem_Collection.Count - 1)
            {

                lblMsg.Text = "Loading Treatment Optimizer Scenario " + p_strScenarioId.Trim() + "...Stand By";
                lblMsg.Show();
                lblMsg.Refresh();
                //load the scenario into the collection
                m_oOptimizerScenarioTools.LoadScenario(
                    p_strScenarioId.Trim(),
                    m_oQueries, ReferenceOptimizerScenarioForm.m_bProcessorUsingSqlite,
                    m_oOptimizerScenarioItem_Collection);
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

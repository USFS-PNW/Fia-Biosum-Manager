using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    public partial class uc_processor_scenario_additional_harvest_cost_column_item : UserControl
    {
        private uc_processor_scenario_additional_harvest_cost_columns _uc_processor_scenario_additional_harvest_cost_columns=null;
        private Queries _oQueries=null;
        private SQLite.ADO.DataMgr _DataMgr;
        private FIA_Biosum_Manager.ValidateNumericValues m_oValidate = new ValidateNumericValues();
        private string m_strCubicFootDollarValueSave = "";
        private frmProcessorScenario _oFrmProcessorScenario = null;

        public frmProcessorScenario ReferenceProcessorScenarioForm
        {
            get { return _oFrmProcessorScenario; }
            set { _oFrmProcessorScenario = value; }
        }
        public string ColumnName
        {
            get { return this.txtColumnName.Text.Trim(); }
            set { this.txtColumnName.Text = value; }
        }
        public string Description
        {
            get { return this.txtDesc.Text.Trim(); }
            set { this.txtDesc.Text = value; }
        }
        public string Type
        {
            get { return this.txtRxScenario.Text; }
            set { this.txtRxScenario.Text = value; }
        }
        public bool EnableColumnNameEditButton
        {
            set {btnColumnNameEdit.Enabled = value; }
        }
        public bool EnableColumnNameRemoveButton
        {
            set {btnColumnNameRemove.Enabled = value; }
        }
       // public bool ShowSeparator
       // {
       //     set { pictureBox1.Visible = value; }
       // }
       // public bool EnableColumnValuesEditButton
        //{
        //    set {btnColumnValuesEdit.Enabled = value; }
       // }
        //public bool EnableColumnValuesEditNullButton
        //{
          //  set { this.btnColumnValuesEditNull.Enabled = value; }
        //}
        public string DefaultCubicFootDollarValue
        {
            get { return this.txtCubicFootDollarValue.Text.Trim(); }
            set 
            { 
                this.txtCubicFootDollarValue.Text = value; 
                ValidateDefaultCubicFootDollarValue();
                this.m_strCubicFootDollarValueSave = this.txtCubicFootDollarValue.Text;
            }
        }
        public uc_processor_scenario_additional_harvest_cost_columns ReferenceAdditionalHarvestCostColumnsUserControl
        {
            set { this._uc_processor_scenario_additional_harvest_cost_columns = value; }
            get { return this._uc_processor_scenario_additional_harvest_cost_columns; }
        }
        public Queries ReferenceQueries
        {
            set { _oQueries = value; }
            get { return _oQueries; }
        }
        public SQLite.ADO.DataMgr ReferenceDataMgr
        {
            set { _DataMgr = value; }
            get { return _DataMgr; }
        }
        public string NullCount
        {
            get { return lblNullValueCount.Text.Trim(); }
            set
            {
                if (Convert.ToInt32(value) > 0)
                {
                    lblNullValueCount.ForeColor = Color.Red;
                }
                else
                {
                    lblNullValueCount.ForeColor = Color.Green;

                }
                lblNullValueCount.Text = value;
            }
        }

        public uc_processor_scenario_additional_harvest_cost_column_item()
        {
            InitializeComponent();
            ReferenceDataMgr = new SQLite.ADO.DataMgr();
        }

        private void txtRxScenario_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtColumnName_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void txtRxScenario_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void btnColumnNameEdit_Click(object sender, EventArgs e)
        {
            int y;
            FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
            frmTemp.MaximizeBox = false;
            frmTemp.MinimizeBox = false;
            frmTemp.BackColor = System.Drawing.SystemColors.Control;
            frmTemp.Initialize_Scenario_Harvest_Costs_Column_Edit_Control();
            

            frmTemp.Height = 0;
            frmTemp.Width = 0;
            if (frmTemp.Top + frmTemp.uc_scenario_harvest_cost_column_edit1.Height > frmTemp.ClientSize.Height + 2)
            {
                for (y = 1; ; y++)
                {
                    frmTemp.Height = y;
                    if (frmTemp.uc_scenario_harvest_cost_column_edit1.Top +
                        frmTemp.uc_scenario_harvest_cost_column_edit1.Height <
                        frmTemp.ClientSize.Height)
                    {
                        break;
                    }
                }

            }
            if (frmTemp.uc_scenario_harvest_cost_column_edit1.Left + frmTemp.uc_scenario_harvest_cost_column_edit1.Width > frmTemp.ClientSize.Width + 2)
            {
                for (y = 1; ; y++)
                {
                    frmTemp.Width = y;
                    if (frmTemp.uc_scenario_harvest_cost_column_edit1.Left +
                        frmTemp.uc_scenario_harvest_cost_column_edit1.Width <
                        frmTemp.ClientSize.Width)
                    {
                        break;
                    }
                }

            }
            frmTemp.Left = 0;
            frmTemp.Top = 0;

            frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            frmTemp.uc_scenario_harvest_cost_column_edit1.Dock = System.Windows.Forms.DockStyle.Fill;

            frmTemp.uc_scenario_harvest_cost_column_edit1.EditType = "Edit";


            frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnText = this.txtColumnName.Text.Trim();
            frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnDescription = this.txtDesc.Text.Trim();
            frmTemp.uc_scenario_harvest_cost_column_edit1.cmbCol.Enabled = false;
            frmTemp.uc_scenario_harvest_cost_column_edit1.lblEdit.Hide();
            

            frmTemp.Text = "Edit Additional CPA Component Column";
            System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                if (txtDesc.Text.Trim().ToUpper() !=
                    frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnDescription.Trim().ToUpper())
                    ReferenceProcessorScenarioForm.m_bSave = true;

                this.txtDesc.Text = frmTemp.uc_scenario_harvest_cost_column_edit1.ColumnDescription;
            }
            frmTemp.Dispose();
        }

        private void btnColumnNameRemove_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you wish to remove harvest cost component " + ColumnName + "?(Y/N)", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                this.ReferenceAdditionalHarvestCostColumnsUserControl.RemoveColumn(this);
                this.ReferenceProcessorScenarioForm.m_bSave = true;
            }
        }

        private void txtCubicFootDollarValue_Leave(object sender, EventArgs e)
        {
            ValidateDefaultCubicFootDollarValue();

        }
        private void ValidateDefaultCubicFootDollarValue()
        {
            m_oValidate.RoundDecimalLength = 2;
            m_oValidate.Money = true;
            m_oValidate.NullsAllowed = false;
            m_oValidate.TestForMaxMin = false;
            m_oValidate.TestForMin = true;
            m_oValidate.MinValue = 0;
            m_oValidate.ValidateDecimal(txtCubicFootDollarValue.Text);
            if (m_oValidate.m_intError == 0)
                txtCubicFootDollarValue.Text = m_oValidate.ReturnValue;
            else
            {
                this.txtCubicFootDollarValue.Text = this.m_strCubicFootDollarValueSave;
                this.txtCubicFootDollarValue.Focus();

            }
        }
        private void btnGo_Click(object sender, EventArgs e)
        {
            switch (this.cmbEdit.Text.Trim())
            {
                case "Assign default value to all  componenet occurances":
                    DefaultValueToAll();
                    break;
                case "Assign default value to all component NULL values":
                    DefaultValueToAllNulls();
                    break;
                case "Assign previously entered values":
                    this.ReferenceAdditionalHarvestCostColumnsUserControl.UpdateItemFromPrevious(this);
                    break;
                case "Edit all values":
                    this.ReferenceAdditionalHarvestCostColumnsUserControl.EditAll(this);
                    break;
                case "Edit NULL values":
                    this.ReferenceAdditionalHarvestCostColumnsUserControl.EditAllNulls(this);
                    break;
            }
        }
        private void DefaultValueToAll()
        {
            string strMsg="";
            if (Type.Trim().ToUpper() != "SCENARIO")
            {
                strMsg = "All rows for treatment " + Type + " component " + ColumnName + "\r\n" +
                         "will be assigned the value " + this.txtCubicFootDollarValue.Text.Trim() + ".\r\n" +
                         "Do you wish to update the table rows with the default value?(Y/N)";
            }
            else
            {
                strMsg = "All rows for component " + ColumnName + "\r\n" +
                        "will be assigned the value " + this.txtCubicFootDollarValue.Text.Trim() + ".\r\n" +
                        "Do you wish to update the table rows with the default value?(Y/N)";
            }
            DialogResult result = MessageBox.Show(strMsg, "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                frmMain.g_oFrmMain.ActivateStandByAnimation(
                    this.ParentForm.WindowState,
                    this.ParentForm.Left,
                    this.ParentForm.Height,
                    this.ParentForm.Width,
                    this.ParentForm.Top);
                frmMain.g_sbpInfo.Text = "Updating Harvest Cost Component $/A/C Values...Stand By";
                                             
                string strScenarioDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(ReferenceDataMgr.GetConnectionString(strScenarioDb)))
                {
                    conn.Open();
                    ReferenceDataMgr.m_strSQL = "UPDATE additional_harvest_costs_work_table " +
                                                "SET " + ColumnName.Trim() + "=" + this.txtCubicFootDollarValue.Text.Replace("$", "");
                    if (Type.Trim().ToUpper() != "SCENARIO")
                    {
                        ReferenceDataMgr.m_strSQL = ReferenceDataMgr.m_strSQL + " WHERE rx='" + Type.Trim() + "'";
                    }
                    ReferenceDataMgr.SqlNonQuery(conn, ReferenceDataMgr.m_strSQL);
                }
                frmMain.g_sbpInfo.Text = "Ready";
                if (ReferenceDataMgr.m_intError == 0)
                {                    
                    this.ReferenceAdditionalHarvestCostColumnsUserControl.UpdateNullCounts(ReferenceAdditionalHarvestCostColumnsUserControl.TempDb);
                    this.ReferenceAdditionalHarvestCostColumnsUserControl.ReferenceProcessorScenarioForm.m_bSave = true;
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();
                    MessageBox.Show("Done");
                }
                else
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();                                                                        
            }
            
        }
        private void DefaultValueToAllNulls()
        {
            string strMsg = "";
            if (Type.Trim().ToUpper() != "SCENARIO")
            {
                strMsg = "All rows for treatment " + Type + " component " + ColumnName + "\r\n" +
                         "will be assigned the value " + this.txtCubicFootDollarValue.Text.Trim() + " if the current value is null.\r\n" +
                         "Do you wish to update the table rows with the default value?(Y/N)";
            }
            else
            {
                strMsg = "All rows for component " + ColumnName + "\r\n" +
                        "will be assigned the value " + this.txtCubicFootDollarValue.Text.Trim() + " if the current value is null.\r\n" +
                        "Do you wish to update the table rows with the default value?(Y/N)";
            }
            DialogResult result = MessageBox.Show(strMsg, "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                frmMain.g_oFrmMain.ActivateStandByAnimation(
                   this.ParentForm.WindowState,
                   this.ParentForm.Left,
                   this.ParentForm.Height,
                   this.ParentForm.Width,
                   this.ParentForm.Top);
                frmMain.g_sbpInfo.Text = "Updating Harvest Cost Component $/A/C Values...Stand By";
                string strScenarioDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(ReferenceDataMgr.GetConnectionString(strScenarioDb)))
                {
                    conn.Open();
                    ReferenceDataMgr.m_strSQL = "UPDATE  additional_harvest_costs_work_table " +
                                                "SET " + ColumnName.Trim() + "=" + this.txtCubicFootDollarValue.Text.Replace("$", "");
                    if (Type.Trim().ToUpper() != "SCENARIO")
                    {
                        ReferenceDataMgr.m_strSQL = ReferenceDataMgr.m_strSQL + " WHERE rx='" + Type.Trim() + "' AND " + ColumnName.Trim() + " IS NULL";
                    }
                    else
                    {
                        ReferenceDataMgr.m_strSQL = ReferenceDataMgr.m_strSQL + " WHERE " + ColumnName.Trim() + " IS NULL";
                    }
                    ReferenceDataMgr.SqlNonQuery(conn, ReferenceDataMgr.m_strSQL);
                }
                frmMain.g_sbpInfo.Text = "Ready";
                if (ReferenceDataMgr.m_intError == 0)
                {
                    this.ReferenceAdditionalHarvestCostColumnsUserControl.UpdateNullCounts(ReferenceAdditionalHarvestCostColumnsUserControl.TempDb);
                    this.ReferenceAdditionalHarvestCostColumnsUserControl.ReferenceProcessorScenarioForm.m_bSave = true;
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();
                    MessageBox.Show("Done");
                }
                else
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();

            }

        }

        private void txtCubicFootDollarValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

        private void txtDesc_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }
       
    }
  
    public class uc_processor_scenario_additional_harvest_cost_column_collection : System.Collections.CollectionBase
    {
        public uc_processor_scenario_additional_harvest_cost_column_collection()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public void Add(uc_processor_scenario_additional_harvest_cost_column_item p_uc)
        {
            // vérify if object is not already in
            if (this.List.Contains(p_uc))
                throw new InvalidOperationException();

            // adding it
            this.List.Add(p_uc);

            // return collection
            //return this;
        }
        public void Remove(int index)
        {
            // Check to see if there is a widget at the supplied index.
            if (index > Count - 1 || index < 0)
            // If no widget exists, a messagebox is shown and the operation 
            // is canColumned.
            {
                System.Windows.Forms.MessageBox.Show("Index not valid!");
            }
            else
            {
                List.RemoveAt(index);
            }
        }
        public FIA_Biosum_Manager.uc_processor_scenario_additional_harvest_cost_column_item Item(int Index)
        {
            // The appropriate item is retrieved from the List object and
            // explicitly cast to the Widget type, then returned to the 
            // caller.
            return (FIA_Biosum_Manager.uc_processor_scenario_additional_harvest_cost_column_item)List[Index];
        }

    }
}

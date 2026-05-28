using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_scenario_merch_chip_value.
	/// </summary>
	public class uc_processor_scenario_merch_chip_value : System.Windows.Forms.UserControl
	{
		public int m_intError=0;
		public string m_strError="";
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Label lblSpcGrp;
		private System.Windows.Forms.Label lblDbhClass;
        private System.Windows.Forms.Label lblChipPct;
        private System.Windows.Forms.Label lblMerchValue;
        private System.Windows.Forms.Label lblMerchPct;
        private System.Windows.Forms.Label lblWood4Value;
        private System.Windows.Forms.Label lblWood4Pct;
        private System.Windows.Forms.Label lblWood5Value;
        private System.Windows.Forms.Label lblWood5Pct;
        private System.Windows.Forms.Label lblWood6Value;
        private System.Windows.Forms.Label lblWood6Pct;
        private System.Windows.Forms.Panel pnlMerchValues;
        private FIA_Biosum_Manager.uc_processor_scenario_spc_dbh_group_value_collection uc_processor_scenario_spc_dbh_group_value_collection1 = new uc_processor_scenario_spc_dbh_group_value_collection();
		private System.Windows.Forms.Label lblChipValue;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtChipValue;
		private RxTools m_oRxTools = new RxTools();
		private string _strScenarioId="";
		private frmProcessorScenario _frmProcessorScenario=null;
		public FIA_Biosum_Manager.ResizeFormUsingVisibleScrollBars m_oResizeForm = new ResizeFormUsingVisibleScrollBars();
		private FIA_Biosum_Manager.ValidateNumericValues m_oValidate = new ValidateNumericValues();
		private string m_strChipValueSave="";
        private Label lblChips;
        private uc_processor_scenario_spc_dbh_group_value uc_processor_scenario_spc_dbh_group_value1;
		
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_processor_scenario_merch_chip_value()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			m_oResizeForm.ScrollBarParentControl=panel1;
			m_oResizeForm.ResizeWidth=true;
			m_oResizeForm.ResizeHeight=true;
			m_oResizeForm.MaximumHeight = 650;
			this.m_oValidate.MinValue=0;
            this.m_oValidate.Money = true;
            this.m_oValidate.RoundDecimalLength=2;
			this.m_oValidate.NullsAllowed=false;
			this.m_oValidate.TestForMax=false;
			this.m_oValidate.TestForMin=true;


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
            //
            //SCENARIO ID
            //
            ScenarioId = this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
            //
            //SCENARIO DB
            //
            string strScenarioDB =
                frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultDbFile;
            ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadSpeciesAndDiameterGroupDollarValues(strScenarioDB,
                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);

                //REMOVE OLD CONTROLS FROM FORM IF THEY EXIST
                string strName = "uc_processor_scenario_spc_dbh_group_value2";
                if (this.pnlMerchValues.Controls[strName] != null)
                {
                    for (x = 2; x <= this.uc_processor_scenario_spc_dbh_group_value_collection1.Count; x++)
                    {
                        strName = "uc_processor_scenario_spc_dbh_group_value" + x;
                        uc_processor_scenario_spc_dbh_group_value oItem = (uc_processor_scenario_spc_dbh_group_value) this.pnlMerchValues.Controls[strName];
                        if (oItem != null)
                        {
                            this.pnlMerchValues.Controls.Remove(oItem);
                        }
                    }
                }

                //REMOVE OLD ITEMS FROM COLLECTION IF THEY EXIST
                if (this.uc_processor_scenario_spc_dbh_group_value_collection1.Count > 0)
                {
                    this.uc_processor_scenario_spc_dbh_group_value_collection1.Clear();
                }
                
                for (x = 0; x <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Count - 1; x++)
                {
                    if (this.uc_processor_scenario_spc_dbh_group_value_collection1.Count == 0)
                    {
                        if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim().Length > 0)
                        {
                          
                            //
                            //SPECIES GROUP
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.SpeciesGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim();
                            //
                            //DBH GROUP
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.DbhGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).DbhGroup.Trim();
                            //
                            //CHIP PERCENT
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.ChipPercent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).ChipPercent;
                            //
                            //MERCH CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchDollarPerCubicFootValue);
                            this.uc_processor_scenario_spc_dbh_group_value1.MerchDollarValue = this.m_oValidate.ReturnValue;
                            //
                            //MERCH PERCENT
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.MerchPercent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchPercent;
                            //
                            //WOOD4 CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood4DollarPerCubicFootValue);
                            this.uc_processor_scenario_spc_dbh_group_value1.Wood4DollarValue = this.m_oValidate.ReturnValue;
                            //
                            //WOOD4 PERCENT
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.Wood4Percent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood4Percent;
                            //
                            //WOOD5 CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood5DollarPerCubicFootValue);
                            this.uc_processor_scenario_spc_dbh_group_value1.Wood5DollarValue = this.m_oValidate.ReturnValue;
                            //
                            //WOOD5 PERCENT
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.Wood5Percent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood5Percent;
                            //
                            //WOOD6 CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood6DollarPerCubicFootValue);
                            this.uc_processor_scenario_spc_dbh_group_value1.Wood6DollarValue = this.m_oValidate.ReturnValue;
                            //
                            //WOOD6 PERCENT
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.Wood6Percent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood6Percent;
                            //
                            //USE AS ENERGY WOOD
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.EnergyWood =
                                  ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).UseAsEnergyWood;

                            this.uc_processor_scenario_spc_dbh_group_value1.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;

                            this.uc_processor_scenario_spc_dbh_group_value_collection1.Add(this.uc_processor_scenario_spc_dbh_group_value1);
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).ChipsDollarPerCubicFootValue);
                            this.txtChipValue.Text = this.m_oValidate.ReturnValue;
                            this.m_strChipValueSave = this.txtChipValue.Text;
                        }
                    }
                    else
                    {
                        if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim().Length > 0)
                        {
                            uc_processor_scenario_spc_dbh_group_value oItem = new uc_processor_scenario_spc_dbh_group_value();
                            //
                            //SPECIES GROUP
                            //
                            oItem.SpeciesGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim();
                            //
                            //DBH GROUP
                            //
                            oItem.DbhGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).DbhGroup.Trim();
                            //
                            //CHIP PERCENT
                            //
                            oItem.ChipPercent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).ChipPercent;
                            //
                            //MERCH CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchDollarPerCubicFootValue);
                            oItem.MerchDollarValue = this.m_oValidate.ReturnValue;
                            //
                            //MERCH PERCENT
                            //
                            oItem.MerchPercent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchPercent;
                            //
                            //WOOD4 CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood4DollarPerCubicFootValue);
                            oItem.Wood4DollarValue = this.m_oValidate.ReturnValue;
                            //
                            //WOOD4 PERCENT
                            //
                            oItem.Wood4Percent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood4Percent;
                            //
                            //WOOD5 CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood5DollarPerCubicFootValue);
                            oItem.Wood5DollarValue = this.m_oValidate.ReturnValue;
                            //
                            //WOOD5 PERCENT
                            //
                            oItem.Wood5Percent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood5Percent;
                            //
                            //WOOD6 CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood6DollarPerCubicFootValue);
                            oItem.Wood6DollarValue = this.m_oValidate.ReturnValue;
                            //
                            //WOOD6 PERCENT
                            //
                            oItem.Wood6Percent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood6Percent;
                            //
                            //USE AS ENERGY WOOD
                            oItem.EnergyWood =
                                  ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).UseAsEnergyWood;
                            oItem.Name = "uc_processor_scenario_spc_dbh_group_value" + Convert.ToString(uc_processor_scenario_spc_dbh_group_value_collection1.Count + 1).Trim();
                            this.pnlMerchValues.Controls.Add(oItem);
                            oItem.Top = uc_processor_scenario_spc_dbh_group_value_collection1.Item(uc_processor_scenario_spc_dbh_group_value_collection1.Count - 1).Top +
                                uc_processor_scenario_spc_dbh_group_value_collection1.Item(uc_processor_scenario_spc_dbh_group_value_collection1.Count - 1).Height;
                            oItem.Left = uc_processor_scenario_spc_dbh_group_value_collection1.Item(uc_processor_scenario_spc_dbh_group_value_collection1.Count - 1).Left;
                            oItem.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
                            oItem.Visible = true;
                            this.uc_processor_scenario_spc_dbh_group_value_collection1.Add(oItem);
                        }
                    }
                }
		}

        public void loadvalues_FromProperties()
        {
            int x;

            //REMOVE OLD CONTROLS FROM FORM IF THEY EXIST
            string strName = "uc_processor_scenario_spc_dbh_group_value2";
            if (this.pnlMerchValues.Controls[strName] != null)
            {
                for (x = 2; x <= this.uc_processor_scenario_spc_dbh_group_value_collection1.Count; x++)
                {
                    strName = "uc_processor_scenario_spc_dbh_group_value" + x;
                    uc_processor_scenario_spc_dbh_group_value oItem = (uc_processor_scenario_spc_dbh_group_value)this.pnlMerchValues.Controls[strName];
                    if (oItem != null)
                    {
                        this.pnlMerchValues.Controls.Remove(oItem);
                    }
                }
            }

            //REMOVE OLD ITEMS FROM COLLECTION IF THEY EXIST
            if (this.uc_processor_scenario_spc_dbh_group_value_collection1.Count > 0)
            {
                this.uc_processor_scenario_spc_dbh_group_value_collection1.Clear();
            }

            if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection != null)
            {
                for (x = 0; x <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Count - 1; x++)
                {
                    if (this.uc_processor_scenario_spc_dbh_group_value_collection1.Count == 0)
                    {
                        if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim().Length > 0)
                        {

                            //
                            //SPECIES GROUP
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.SpeciesGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim();
                            //
                            //DBH GROUP
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.DbhGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).DbhGroup.Trim();
                            //
                            //CHIP PERCENT
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.ChipPercent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).ChipPercent;
                            //
                            //MERCH CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchDollarPerCubicFootValue);
                            this.uc_processor_scenario_spc_dbh_group_value1.MerchDollarValue = this.m_oValidate.ReturnValue;
                            //
                            //MERCH PERCENT
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.MerchPercent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchPercent;
                            //
                            //WOOD4 CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood4DollarPerCubicFootValue);
                            this.uc_processor_scenario_spc_dbh_group_value1.Wood4DollarValue = this.m_oValidate.ReturnValue;
                            //
                            //WOOD4 PERCENT
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.Wood4Percent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood4Percent;
                            //
                            //WOOD5 CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood5DollarPerCubicFootValue);
                            this.uc_processor_scenario_spc_dbh_group_value1.Wood5DollarValue = this.m_oValidate.ReturnValue;
                            //
                            //WOOD5 PERCENT
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.Wood5Percent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood5Percent;
                            //
                            //WOOD6 CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood6DollarPerCubicFootValue);
                            this.uc_processor_scenario_spc_dbh_group_value1.Wood6DollarValue = this.m_oValidate.ReturnValue;
                            //
                            //WOOD6 PERCENT
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.Wood6Percent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood6Percent;
                            //
                            //USE AS ENERGY WOOD
                            //
                            this.uc_processor_scenario_spc_dbh_group_value1.EnergyWood =
                                  ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).UseAsEnergyWood;

                            this.uc_processor_scenario_spc_dbh_group_value1.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;

                            this.uc_processor_scenario_spc_dbh_group_value_collection1.Add(this.uc_processor_scenario_spc_dbh_group_value1);
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).ChipsDollarPerCubicFootValue);
                            this.txtChipValue.Text = this.m_oValidate.ReturnValue;
                            this.m_strChipValueSave = this.txtChipValue.Text;
                        }
                    }
                    else
                    {
                        if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim().Length > 0)
                        {
                            uc_processor_scenario_spc_dbh_group_value oItem = new uc_processor_scenario_spc_dbh_group_value();
                            //
                            //SPECIES GROUP
                            //
                            oItem.SpeciesGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim();
                            //
                            //DBH GROUP
                            //
                            oItem.DbhGroup =
                                ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).DbhGroup.Trim();
                            //
                            //CHIP PERCENT
                            //
                            oItem.ChipPercent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).ChipPercent;
                            //
                            //MERCH CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchDollarPerCubicFootValue);
                            oItem.MerchDollarValue = this.m_oValidate.ReturnValue;
                            //
                            //MERCH PERCENT
                            //
                            oItem.MerchPercent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchPercent;
                            //
                            //WOOD4 CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood4DollarPerCubicFootValue);
                            oItem.Wood4DollarValue = this.m_oValidate.ReturnValue;
                            //
                            //WOOD4 PERCENT
                            //
                            oItem.Wood4Percent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood4Percent;
                            //
                            //WOOD5 CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood5DollarPerCubicFootValue);
                            oItem.Wood5DollarValue = this.m_oValidate.ReturnValue;
                            //
                            //WOOD5 PERCENT
                            //
                            oItem.Wood5Percent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood5Percent;
                            //
                            //WOOD6 CUBIC FOOT VALUE
                            //
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood6DollarPerCubicFootValue);
                            oItem.Wood6DollarValue = this.m_oValidate.ReturnValue;
                            //
                            //WOOD6 PERCENT
                            //
                            oItem.Wood6Percent = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).Wood6Percent;
                            //
                            //USE AS ENERGY WOOD
                            //
                            oItem.EnergyWood =
                                  ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).UseAsEnergyWood;
                            oItem.Name = "uc_processor_scenario_spc_dbh_group_value" + Convert.ToString(uc_processor_scenario_spc_dbh_group_value_collection1.Count + 1).Trim();
                            this.pnlMerchValues.Controls.Add(oItem);
                            oItem.Top = uc_processor_scenario_spc_dbh_group_value_collection1.Item(uc_processor_scenario_spc_dbh_group_value_collection1.Count - 1).Top +
                                uc_processor_scenario_spc_dbh_group_value_collection1.Item(uc_processor_scenario_spc_dbh_group_value_collection1.Count - 1).Height;
                            oItem.Left = uc_processor_scenario_spc_dbh_group_value_collection1.Item(uc_processor_scenario_spc_dbh_group_value_collection1.Count - 1).Left;
                            oItem.ReferenceProcessorScenarioForm = ReferenceProcessorScenarioForm;
                            oItem.Visible = true;
                            this.uc_processor_scenario_spc_dbh_group_value_collection1.Add(oItem);
                            this.m_oValidate.ValidateDecimal(ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).ChipsDollarPerCubicFootValue);
                            this.txtChipValue.Text = this.m_oValidate.ReturnValue;
                            this.m_strChipValueSave = this.txtChipValue.Text;
                        }
                    }
                }
            }
        }

		public void savevalues()
		{
			int x;
			string strSpcGrp="";
			string strDbhGrp="";
            string strChipValue = "";
            string strChipPct = "";
            string strMerchValue="";
            string strMerchPct = "";
            string strWood4Value = "";
            string strWood4Pct = "";
            string strWood5Value = "";
            string strWood5Pct = "";
            string strWood6Value = "";
            string strWood6Pct = "";
            string strSql = "";

            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            string strScenarioDB =
                frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultDbFile;
            oDataMgr.OpenConnection(oDataMgr.GetConnectionString(strScenarioDB));
            if (oDataMgr.m_intError != 0)
            {
                m_intError = oDataMgr.m_intError;
                m_strError = oDataMgr.m_strError;
                oDataMgr = null;
                return;
            }
            m_intError = 0;
            m_strError = "";
            //
            //DELETE THE CURRENT SCENARIO RECORDS
            //
            oDataMgr.m_strSQL = "DELETE FROM scenario_tree_species_diam_dollar_values " +
                            "WHERE TRIM(scenario_id)='" + this.ScenarioId.Trim() + "'";
            oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
            //
            //DELETE THE WORK TABLE
            //
            if (oDataMgr.TableExist(oDataMgr.m_Connection, "spcgrp_dbhgrp"))
                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, "DROP TABLE spcgrp_dbhgrp");
            //
            //CREATE AND POPULATE WORK TABLE
            //
            oDataMgr.m_strSQL = "CREATE TABLE spcgrp_dbhgrp (" +
                    "species_group INTEGER," +
                    "species_label TEXT," +
                    "diam_group INTEGER," +
                    "diam_class TEXT)";
            oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);

            foreach (ProcessorScenarioItem.SpcGroupItem objSpcGroup in ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupItem_Collection)
            {
                foreach (ProcessorScenarioItem.TreeDiamGroupsItem objDiamGroup in ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection)
                {
                    // INITIALIZE RECORDS IN WORK TABLE
                    oDataMgr.m_strSQL = "INSERT INTO spcgrp_dbhgrp (species_group,species_label, diam_group, diam_class) " +
                        "VALUES (" + objSpcGroup.SpeciesGroup + ",'" + objSpcGroup.SpeciesGroupLabel + "'," +
                        objDiamGroup.DiamGroup + ",'" + objDiamGroup.DiamClass + "')";
                    oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                    //
                    //INSERT SCENARIO RECORDS
                    //
                    oDataMgr.m_strSQL = "INSERT INTO scenario_tree_species_diam_dollar_values (scenario_id,species_group,diam_group) " +
                                    "VALUES ('" + ScenarioId.Trim() + "'," + objSpcGroup.SpeciesGroup + "," +
                                    objDiamGroup.DiamGroup + ")";
                    oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                }
            }

			//
			//UPDATE SCENARIO RECORDS WITH MERCH AND CHIP VALUES
			//
			for (x=0;x<=this.uc_processor_scenario_spc_dbh_group_value_collection1.Count-1;x++)
			{
				strSpcGrp = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).SpeciesGroup.Trim();
				strDbhGrp = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).DbhGroup.Trim();
                strChipValue = this.txtChipValue.Text.Trim();
                strChipValue = strChipValue.Replace("$", "");
                strChipPct = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).ChipPercent.Trim();
                strMerchValue = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).MerchDollarValue.Trim();
				strMerchValue = strMerchValue.Replace("$","");
                strMerchPct = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).MerchPercent.Trim();
                strWood4Value = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).Wood4DollarValue.Trim();
                strWood4Value = strWood4Value.Replace("$", "");
                strWood4Pct = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).Wood4Percent.Trim();
                strWood5Value = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).Wood5DollarValue.Trim();
                strWood5Value = strWood5Value.Replace("$", "");
                strWood5Pct = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).Wood5Percent.Trim();
                strWood6Value = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).Wood6DollarValue.Trim();
                strWood6Value = strWood6Value.Replace("$", "");
                strWood6Pct = uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).Wood6Percent.Trim();


                strSql = "UPDATE scenario_tree_species_diam_dollar_values " +
                    "SET chip_value = " + strChipValue + "," +
                    "chip_pct = " + strChipPct + "," +
                    "merch_value = " + strMerchValue + "," +
                    "merch_pct = " + strMerchPct + "," +
                    "wood4_value = " + strWood4Value + "," +
                    "wood4_pct = " + strWood4Pct + "," +
                    "wood5_value = " + strWood5Value + "," +
                    "wood5_pct = " + strWood5Pct + "," +
                    "wood6_value = " + strWood6Value + "," +
                    "wood6_pct = " + strWood6Pct + " WHERE EXISTS (" +
                    "SELECT * FROM spcgrp_dbhgrp " +
                    "WHERE scenario_tree_species_diam_dollar_values.species_group = spcgrp_dbhgrp.species_group " +
                    "AND scenario_tree_species_diam_dollar_values.diam_group = spcgrp_dbhgrp.diam_group " +
                    "AND TRIM(scenario_tree_species_diam_dollar_values.scenario_id) = '" + ScenarioId.Trim() + "' " +
                    "AND TRIM(spcgrp_dbhgrp.species_label) = '" + strSpcGrp + "' " +
                    "AND TRIM(spcgrp_dbhgrp.diam_class) = '" + strDbhGrp + "')";

                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, strSql);				
				uc_processor_scenario_spc_dbh_group_value_collection1.Item(x).SaveValues();							
			}
			this.m_strChipValueSave= this.txtChipValue.Text;
            //
            //DELETE THE WORK TABLE AND CLOSE CONNECTION
            //
            if (oDataMgr.TableExist(oDataMgr.m_Connection, "spcgrp_dbhgrp"))
                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, "DROP TABLE spcgrp_dbhgrp");
            m_intError = oDataMgr.m_intError;

            oDataMgr.CloseConnection(oDataMgr.m_Connection);
            oDataMgr = null;			
		}
		public frmProcessorScenario ReferenceProcessorScenarioForm
		{
			get {return this._frmProcessorScenario;}
			set {this._frmProcessorScenario=value;}
		}
		public string ScenarioId
		{
			get {return _strScenarioId;}
			set {_strScenarioId=value;}
		}
        public uc_processor_scenario_spc_dbh_group_value_collection ReferenceUserControlMarketValueSpeciesDbhGroupCollection
        {
            get { return uc_processor_scenario_spc_dbh_group_value_collection1; }
        }
        public string MarketValueChips
        {
            get { return txtChipValue.Text.Trim(); }
        }


		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblChips = new System.Windows.Forms.Label();
            this.txtChipValue = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lblChipValue = new System.Windows.Forms.Label();
            this.pnlMerchValues = new System.Windows.Forms.Panel();
            this.uc_processor_scenario_spc_dbh_group_value1 = new FIA_Biosum_Manager.uc_processor_scenario_spc_dbh_group_value();
            this.lblChipPct = new System.Windows.Forms.Label();
            this.lblMerchValue = new System.Windows.Forms.Label();
            this.lblMerchPct = new System.Windows.Forms.Label();
            this.lblWood4Value = new System.Windows.Forms.Label();
            this.lblWood4Pct = new System.Windows.Forms.Label();
            this.lblWood5Value = new System.Windows.Forms.Label();
            this.lblWood5Pct = new System.Windows.Forms.Label();
            this.lblWood6Value = new System.Windows.Forms.Label();
            this.lblWood6Pct = new System.Windows.Forms.Label();
            this.lblDbhClass = new System.Windows.Forms.Label();
            this.lblSpcGrp = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.pnlMerchValues.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(696, 504);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.txtChipValue);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.lblChipValue);
            this.panel1.Controls.Add(this.pnlMerchValues);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 48);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(690, 453);
            this.panel1.TabIndex = 30;
            // 
            // txtChipValue
            // 
            this.txtChipValue.Location = new System.Drawing.Point(408, 80);
            this.txtChipValue.Name = "txtChipValue";
            this.txtChipValue.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtChipValue.Size = new System.Drawing.Size(72, 20);
            this.txtChipValue.TabIndex = 8;
            this.txtChipValue.Text = "$0.00";
            this.txtChipValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtChipValue.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtChipValue_KeyPress);
            this.txtChipValue.Leave += new System.EventHandler(this.txtChipValue_Leave);
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label2.Location = new System.Drawing.Point(15, 80);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(400, 24);
            this.label2.TabIndex = 7;
            this.label2.Text = " Hog fuel (chipped residues) value ($/green ton):";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label1.Location = new System.Drawing.Point(15, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(675, 32);
            this.label1.TabIndex = 6;
            this.label1.Text = "Enter % allocation and wood value  ($/ft3) per species group/diameter class and roundwood type";
            // 
            // lblChipValue
            // 
            this.lblChipValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblChipValue.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblChipValue.Location = new System.Drawing.Point(15, 48);
            this.lblChipValue.Name = "lblChipValue";
            this.lblChipValue.Size = new System.Drawing.Size(203, 24);
            this.lblChipValue.TabIndex = 5;
            this.lblChipValue.Text = "Chip (Energy) Wood Values";
            // 
            // pnlMerchValues
            // 
            this.pnlMerchValues.AutoScroll = true;
            this.pnlMerchValues.BackColor = System.Drawing.SystemColors.Control;
            this.pnlMerchValues.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pnlMerchValues.Controls.Add(this.uc_processor_scenario_spc_dbh_group_value1);
            this.pnlMerchValues.Controls.Add(this.lblSpcGrp);
            this.pnlMerchValues.Controls.Add(this.lblDbhClass);
            this.pnlMerchValues.Controls.Add(this.lblChips);
            this.pnlMerchValues.Controls.Add(this.lblChipPct);
            this.pnlMerchValues.Controls.Add(this.lblMerchValue);
            this.pnlMerchValues.Controls.Add(this.lblMerchPct);
            this.pnlMerchValues.Controls.Add(this.lblWood4Value);
            this.pnlMerchValues.Controls.Add(this.lblWood4Pct);
            this.pnlMerchValues.Controls.Add(this.lblWood5Value);
            this.pnlMerchValues.Controls.Add(this.lblWood5Pct);
            this.pnlMerchValues.Controls.Add(this.lblWood6Value);
            this.pnlMerchValues.Controls.Add(this.lblWood6Pct);
            this.pnlMerchValues.Location = new System.Drawing.Point(8, 122);
            this.pnlMerchValues.Name = "pnlMerchValues";
            this.pnlMerchValues.Size = new System.Drawing.Size(676, 317);
            this.pnlMerchValues.TabIndex = 3;
            // 
            // uc_processor_scenario_spc_dbh_group_value1
            // 
            this.uc_processor_scenario_spc_dbh_group_value1.ChipPercent = "0";
            this.uc_processor_scenario_spc_dbh_group_value1.MerchDollarValue = "$0.00";
            this.uc_processor_scenario_spc_dbh_group_value1.MerchPercent = "100";
            this.uc_processor_scenario_spc_dbh_group_value1.Wood4DollarValue = "$0.00";
            this.uc_processor_scenario_spc_dbh_group_value1.Wood4Percent = "0";
            this.uc_processor_scenario_spc_dbh_group_value1.Wood5DollarValue = "$0.00";
            this.uc_processor_scenario_spc_dbh_group_value1.Wood5Percent = "0";
            this.uc_processor_scenario_spc_dbh_group_value1.Wood4DollarValue = "$0.00";
            this.uc_processor_scenario_spc_dbh_group_value1.Wood6Percent = "0";
            this.uc_processor_scenario_spc_dbh_group_value1.DbhGroup = "";
            this.uc_processor_scenario_spc_dbh_group_value1.EnergyWood = false;
            this.uc_processor_scenario_spc_dbh_group_value1.Location = new System.Drawing.Point(-2, 70);
            this.uc_processor_scenario_spc_dbh_group_value1.Name = "uc_processor_scenario_spc_dbh_group_value1";
            this.uc_processor_scenario_spc_dbh_group_value1.ReferenceProcessorScenarioForm = null;
            this.uc_processor_scenario_spc_dbh_group_value1.Size = new System.Drawing.Size(915, 30);
            this.uc_processor_scenario_spc_dbh_group_value1.SpeciesGroup = "";
            this.uc_processor_scenario_spc_dbh_group_value1.TabIndex = 0;
            // 
            // lblDbhClass
            // 
            this.lblDbhClass.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDbhClass.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblDbhClass.Location = new System.Drawing.Point(164, 4);
            this.lblDbhClass.Name = "lblDbhClass";
            this.lblDbhClass.Size = new System.Drawing.Size(66, 65);
            this.lblDbhClass.TabIndex = 2;
            this.lblDbhClass.Text = "Tree DBH Class (inches)";
            this.lblDbhClass.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // lblSpcGrp
            // 
            this.lblSpcGrp.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSpcGrp.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblSpcGrp.Location = new System.Drawing.Point(29, 4);
            this.lblSpcGrp.Name = "lblSpcGrp";
            this.lblSpcGrp.Size = new System.Drawing.Size(104, 65);
            this.lblSpcGrp.TabIndex = 1;
            this.lblSpcGrp.Text = "Species Group";
            this.lblSpcGrp.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // lblChips
            // 
            this.lblChips.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblChips.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblChips.Location = new System.Drawing.Point(238, 4);
            this.lblChips.Name = "lblChips";
            this.lblChips.Size = new System.Drawing.Size(66, 65);
            this.lblChips.TabIndex = 3;
            this.lblChips.Text = "Allocate All to Energy Wood";
            this.lblChips.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            //
            // lblChipPct
            //
            this.lblChipPct.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblChipPct.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblChipPct.Location = new System.Drawing.Point(311, 4);
            this.lblChipPct.Name = "lblChipPct";
            this.lblChipPct.Size = new System.Drawing.Size(60, 65);
            this.lblChipPct.TabIndex = 4;
            this.lblChipPct.Text = "Pct Energy";
            this.lblChipPct.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // lblMerchValue
            // 
            this.lblMerchValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMerchValue.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblMerchValue.Location = new System.Drawing.Point(379, 4);
            this.lblMerchValue.Name = "lblMerchValue";
            this.lblMerchValue.Size = new System.Drawing.Size(60, 65);
            this.lblMerchValue.TabIndex = 5;
            this.lblMerchValue.Text = "Merch Value in $/ft3";
            this.lblMerchValue.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            //
            // lblMerchPct
            //
            this.lblMerchPct.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMerchPct.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblMerchPct.Location = new System.Drawing.Point(447, 4);
            this.lblMerchPct.Name = "lblMerchPct";
            this.lblMerchPct.Size = new System.Drawing.Size(60, 65);
            this.lblMerchPct.TabIndex = 6;
            this.lblMerchPct.Text = "Pct Merch";
            this.lblMerchPct.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            //
            // lblWood4Value
            //
            this.lblWood4Value.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWood4Value.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblWood4Value.Location = new System.Drawing.Point(515, 4);
            this.lblWood4Value.Name = "lblWood4Value";
            this.lblWood4Value.Size = new System.Drawing.Size(60, 65);
            this.lblWood4Value.TabIndex = 7;
            this.lblWood4Value.Text = "Wood4 Value in $/ft3";
            this.lblWood4Value.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            //
            // lblWood4Pct
            //
            this.lblWood4Pct.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWood4Pct.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblWood4Pct.Location = new System.Drawing.Point(583, 4);
            this.lblWood4Pct.Name = "lblWood4Pct";
            this.lblWood4Pct.Size = new System.Drawing.Size(60, 65);
            this.lblWood4Pct.TabIndex = 8;
            this.lblWood4Pct.Text = "Pct Wood4";
            this.lblWood4Pct.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            //
            // lblWood5Value
            //
            this.lblWood5Value.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWood5Value.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblWood5Value.Location = new System.Drawing.Point(651, 4);
            this.lblWood5Value.Name = "lblWood5Value";
            this.lblWood5Value.Size = new System.Drawing.Size(60, 65);
            this.lblWood5Value.TabIndex = 9;
            this.lblWood5Value.Text = "Wood5 Value in $/ft3";
            this.lblWood5Value.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            //
            // lblWood5Pct
            //
            this.lblWood5Pct.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWood5Pct.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblWood5Pct.Location = new System.Drawing.Point(719, 4);
            this.lblWood5Pct.Name = "lblWood5Pct";
            this.lblWood5Pct.Size = new System.Drawing.Size(60, 65);
            this.lblWood5Pct.TabIndex = 10;
            this.lblWood5Pct.Text = "Pct Wood5";
            this.lblWood5Pct.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            //
            // lblWood6Value
            //
            this.lblWood6Value.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWood6Value.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblWood6Value.Location = new System.Drawing.Point(787, 4);
            this.lblWood6Value.Name = "lblWood6Value";
            this.lblWood6Value.Size = new System.Drawing.Size(60, 65);
            this.lblWood6Value.TabIndex = 11;
            this.lblWood6Value.Text = "Wood6 Value in $/ft3";
            this.lblWood6Value.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            //
            // lblWood6Pct
            //
            this.lblWood6Pct.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWood6Pct.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblWood6Pct.Location = new System.Drawing.Point(855, 4);
            this.lblWood6Pct.Name = "lblWood6Pct";
            this.lblWood6Pct.Size = new System.Drawing.Size(60, 65);
            this.lblWood6Pct.TabIndex = 12;
            this.lblWood6Pct.Text = "Pct Wood6";
            this.lblWood6Pct.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(690, 32);
            this.lblTitle.TabIndex = 29;
            this.lblTitle.Text = "Values Assumed for Delivered Wood at the Mill or Processing Site gate";
            // 
            // uc_processor_scenario_merch_chip_value
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_scenario_merch_chip_value";
            this.Size = new System.Drawing.Size(696, 504);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.pnlMerchValues.ResumeLayout(false);
            this.ResumeLayout(false);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;

        }
		#endregion

		private void label2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtChipValue_Leave(object sender, System.EventArgs e)
		{
			m_oValidate.Money=true;
			m_oValidate.RoundDecimalLength=2;
			m_oValidate.NullsAllowed=false;
			m_oValidate.TestForMax=false;
			m_oValidate.TestForMin=true;
			m_oValidate.MinValue=0;
			m_oValidate.ValidateDecimal(this.txtChipValue.Text.Trim());
            if (m_oValidate.m_intError == 0)
            {
                
                this.txtChipValue.Text = m_oValidate.ReturnValue;
                
            }
            else
                this.txtChipValue.Focus();

		}

        private void txtChipValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            ReferenceProcessorScenarioForm.m_bSave = true;
        }

        public void resize_ValuePanel()
        {
            if (this.panel1.Width - this.pnlMerchValues.Width > 16)
            {
                this.pnlMerchValues.Width = this.panel1.Width - 16;
            }

            if (this.pnlMerchValues.Width > 950)
            {
                this.pnlMerchValues.Width = 950;
            }
        }
	}
}

using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_tree_spc_edit.
	/// </summary>
	public class uc_processor_tree_spc_edit : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Label lblID;
		private System.Windows.Forms.ComboBox cmbVariant;
		private System.Windows.Forms.TextBox txtSpCd;
		private System.Windows.Forms.TextBox txtCommon;
		private int m_intError=0;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label10;
		private System.Windows.Forms.TextBox txtGenus;
        private System.Windows.Forms.TextBox txtSpecies;
		private System.Windows.Forms.ComboBox cmbFvsSpCd;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label14;
		private string m_strFvsTreeSpcTable;
		private string m_strVariant;
        private int m_intConvertToSpCd;
        private string m_strFvsSpeciesCode;
        private System.Collections.Generic.IDictionary<String, String> m_dictFvsCommonName;
        private ado_data_access m_ado;
        private Label label8;
        private Label lblFvsCommonName;
        private string m_strTreeSpcTable;

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_processor_tree_spc_edit(ado_data_access p_ado,string p_strTreeSpcTable,string p_strFvsTreeSpcTable,string p_strVariant)
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

            m_ado = p_ado;                
            // Populate reference dictionary of spcd and fvs_common_name from fvs_tree_species
            m_dictFvsCommonName = new System.Collections.Generic.Dictionary<String, String>();
            m_strTreeSpcTable = p_strTreeSpcTable;

			if (p_strVariant.Trim().Length > 0)
			{
				p_ado.m_strSQL = "SELECT  spcd,fvs_species,common_name,fvs_common_name " + 
					"FROM " + p_strFvsTreeSpcTable + " " + 
					"WHERE LEN(TRIM(fvs_species)) > 0 AND " + 
					"LEN(TRIM(common_name)) > 0 AND " + 
					"TRIM(fvs_variant) = '" + p_strVariant.Trim() +  "' order by spcd;";
			
				p_ado.SqlQueryReader(p_ado.m_OleDbConnection,p_ado.m_OleDbTransaction,p_ado.m_strSQL);
				if (p_ado.m_OleDbDataReader.HasRows)
				{
					while (p_ado.m_OleDbDataReader.Read())
					{
                        string strMySpCd = Convert.ToString(p_ado.m_OleDbDataReader["spcd"]);
                        this.cmbFvsSpCd.Items.Add(strMySpCd + " - " + Convert.ToString(p_ado.m_OleDbDataReader["common_name"]) + " - " + p_ado.m_OleDbDataReader["fvs_species"]);

                        if (!m_dictFvsCommonName.ContainsKey(strMySpCd))
                        {
                            m_dictFvsCommonName.Add(strMySpCd, Convert.ToString(p_ado.m_OleDbDataReader["fvs_common_name"]));
                        }
					}

				}

                // Only show fvs variants on the form that exist in the fvs_tree_species table
                p_ado.m_strSQL = "SELECT distinct fvs_variant " +
                    "FROM " + p_strFvsTreeSpcTable + " ";
                p_ado.SqlQueryReader(p_ado.m_OleDbConnection, p_ado.m_OleDbTransaction, p_ado.m_strSQL);
                if (p_ado.m_OleDbDataReader.HasRows)
                {
                    // Copy all the combo box items into dictionary with variant as key
                    System.Collections.Generic.IDictionary<String, String> dictVariantItems = 
                        new System.Collections.Generic.Dictionary<String, String>();
                    for (int i = 0; i < this.cmbVariant.Items.Count; i++)
                    {
                        string strKey = this.cmbVariant.GetItemText(this.cmbVariant.Items[i]).Substring(0, 2);
                        if (!dictVariantItems.ContainsKey(strKey))
                        {
                            dictVariantItems.Add(strKey, this.cmbVariant.GetItemText(this.cmbVariant.Items[i]));
                        }
                    }
                    // Clear combo box items
                    this.cmbVariant.Items.Clear();
                    while (p_ado.m_OleDbDataReader.Read())
                    {
                        string strNextVariant = Convert.ToString(p_ado.m_OleDbDataReader["fvs_variant"]).Trim();
                        // Only add combo box items back that are in the fvs_species table
                        if (dictVariantItems.ContainsKey(strNextVariant))
                        {
                            this.cmbVariant.Items.Add(dictVariantItems[strNextVariant]);
                        }
                    }
                }

				p_ado.m_OleDbDataReader.Close();
			}
            this.m_strFvsTreeSpcTable = p_strFvsTreeSpcTable;
			this.m_strVariant = p_strVariant;

			


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
            this.lblFvsCommonName = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.cmbFvsSpCd = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtSpecies = new System.Windows.Forms.TextBox();
            this.txtGenus = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtCommon = new System.Windows.Forms.TextBox();
            this.txtSpCd = new System.Windows.Forms.TextBox();
            this.cmbVariant = new System.Windows.Forms.ComboBox();
            this.lblID = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblFvsCommonName);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label14);
            this.groupBox1.Controls.Add(this.cmbFvsSpCd);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txtSpecies);
            this.groupBox1.Controls.Add(this.txtGenus);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.txtCommon);
            this.groupBox1.Controls.Add(this.txtSpCd);
            this.groupBox1.Controls.Add(this.cmbVariant);
            this.groupBox1.Controls.Add(this.lblID);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnOK);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.groupBox1.Size = new System.Drawing.Size(624, 450);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // lblFvsCommonName
            // 
            this.lblFvsCommonName.BackColor = System.Drawing.SystemColors.Control;
            this.lblFvsCommonName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFvsCommonName.Location = new System.Drawing.Point(280, 332);
            this.lblFvsCommonName.Name = "lblFvsCommonName";
            this.lblFvsCommonName.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblFvsCommonName.Size = new System.Drawing.Size(296, 24);
            this.lblFvsCommonName.TabIndex = 39;
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(14, 332);
            this.label8.Name = "label8";
            this.label8.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label8.Size = new System.Drawing.Size(256, 16);
            this.label8.TabIndex = 37;
            this.label8.Text = "FVS Tree Species Common Name";
            // 
            // label14
            // 
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(15, 256);
            this.label14.Name = "label14";
            this.label14.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.label14.Size = new System.Drawing.Size(584, 37);
            this.label14.TabIndex = 36;
            this.label14.Text = "Optional: Enter an alternate FIA tree species code.  When the application creates" +
    " FVS input files, it will translate the tree species code to the tree species co" +
    "de defined below.";
            // 
            // cmbFvsSpCd
            // 
            this.cmbFvsSpCd.Location = new System.Drawing.Point(279, 300);
            this.cmbFvsSpCd.Name = "cmbFvsSpCd";
            this.cmbFvsSpCd.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmbFvsSpCd.Size = new System.Drawing.Size(296, 28);
            this.cmbFvsSpCd.TabIndex = 33;
            this.cmbFvsSpCd.SelectedIndexChanged += new System.EventHandler(this.cmbFvsSpCd_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(111, 300);
            this.label4.Name = "label4";
            this.label4.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label4.Size = new System.Drawing.Size(160, 16);
            this.label4.TabIndex = 32;
            this.label4.Text = "FVS Tree Species Code";
            // 
            // txtSpecies
            // 
            this.txtSpecies.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSpecies.Location = new System.Drawing.Point(280, 216);
            this.txtSpecies.MaxLength = 50;
            this.txtSpecies.Name = "txtSpecies";
            this.txtSpecies.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtSpecies.Size = new System.Drawing.Size(296, 26);
            this.txtSpecies.TabIndex = 26;
            // 
            // txtGenus
            // 
            this.txtGenus.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtGenus.Location = new System.Drawing.Point(280, 184);
            this.txtGenus.MaxLength = 50;
            this.txtGenus.Name = "txtGenus";
            this.txtGenus.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtGenus.Size = new System.Drawing.Size(296, 26);
            this.txtGenus.TabIndex = 25;
            // 
            // label10
            // 
            this.label10.Location = new System.Drawing.Point(80, 220);
            this.label10.Name = "label10";
            this.label10.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label10.Size = new System.Drawing.Size(184, 16);
            this.label10.TabIndex = 18;
            this.label10.Text = "Species";
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(32, 220);
            this.label7.Name = "label7";
            this.label7.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label7.Size = new System.Drawing.Size(232, 16);
            this.label7.TabIndex = 15;
            // 
            // txtCommon
            // 
            this.txtCommon.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCommon.Location = new System.Drawing.Point(280, 152);
            this.txtCommon.MaxLength = 50;
            this.txtCommon.Name = "txtCommon";
            this.txtCommon.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtCommon.Size = new System.Drawing.Size(296, 26);
            this.txtCommon.TabIndex = 2;
            // 
            // txtSpCd
            // 
            this.txtSpCd.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSpCd.Location = new System.Drawing.Point(280, 120);
            this.txtSpCd.MaxLength = 4;
            this.txtSpCd.Name = "txtSpCd";
            this.txtSpCd.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.txtSpCd.Size = new System.Drawing.Size(56, 26);
            this.txtSpCd.TabIndex = 1;
            this.txtSpCd.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSpCd_KeyDown);
            this.txtSpCd.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSpCd_KeyPress);
            this.txtSpCd.Validating += new System.ComponentModel.CancelEventHandler(this.txtSpCd_Validating);
            // 
            // cmbVariant
            // 
            this.cmbVariant.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbVariant.Items.AddRange(new object[] {
            "AK - SouthEast Alaska, Coastal BC",
            "BM - Blue Mountains",
            "CA - Inland CA, Southern Cascades ",
            "CI - Central Idaho",
            "CR - Central Rockies",
            "CS - Central States",
            "EC - Eastside Cascades",
            "EM - Eastern Montana",
            "IE - Inland Empire",
            "KT - Kootenai/Kaniksu/Tally Lake",
            "LS - Lake States",
            "NC - Northern California (Klamath Mountains)",
            "NE - Northeast",
            "NI - Northern Idaho (Inland Empire)",
            "PN - Pacific Northwest Coast",
            "SE - Southeast",
            "SN - Southhern",
            "SO - South Central Oregon, Northeast California",
            "TT - Tetons",
            "UT - Utah",
            "WC - Weside Cascades",
            "WS - Westside Sierra Nevada"});
            this.cmbVariant.Location = new System.Drawing.Point(280, 88);
            this.cmbVariant.Name = "cmbVariant";
            this.cmbVariant.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.cmbVariant.Size = new System.Drawing.Size(296, 28);
            this.cmbVariant.TabIndex = 0;
            this.cmbVariant.SelectedIndexChanged += new System.EventHandler(this.cmbVariant_SelectedIndexChanged);
            this.cmbVariant.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVariant_KeyDown);
            // 
            // lblID
            // 
            this.lblID.BackColor = System.Drawing.SystemColors.Control;
            this.lblID.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblID.Location = new System.Drawing.Point(280, 63);
            this.lblID.Name = "lblID";
            this.lblID.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblID.Size = new System.Drawing.Size(64, 24);
            this.lblID.TabIndex = 13;
            this.lblID.Text = "0";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(322, 385);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(88, 48);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(234, 385);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(88, 48);
            this.btnOK.TabIndex = 5;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // label6
            // 
            this.label6.Location = new System.Drawing.Point(104, 92);
            this.label6.Name = "label6";
            this.label6.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label6.Size = new System.Drawing.Size(160, 16);
            this.label6.TabIndex = 8;
            this.label6.Text = "FVS Variant";
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(104, 192);
            this.label5.Name = "label5";
            this.label5.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label5.Size = new System.Drawing.Size(160, 16);
            this.label5.TabIndex = 7;
            this.label5.Text = "Genus";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(112, 160);
            this.label3.Name = "label3";
            this.label3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label3.Size = new System.Drawing.Size(152, 16);
            this.label3.TabIndex = 5;
            this.label3.Text = "Common Name";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(104, 120);
            this.label2.Name = "label2";
            this.label2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label2.Size = new System.Drawing.Size(160, 16);
            this.label2.TabIndex = 4;
            this.label2.Text = "FIA Tree Species Code";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(232, 64);
            this.label1.Name = "label1";
            this.label1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.label1.Size = new System.Drawing.Size(32, 24);
            this.label1.TabIndex = 3;
            this.label1.Text = "ID";
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 22);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.lblTitle.Size = new System.Drawing.Size(618, 24);
            this.lblTitle.TabIndex = 2;
            this.lblTitle.Text = "Processor Tree Species Edit";
            // 
            // uc_processor_tree_spc_edit
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_tree_spc_edit";
            this.Size = new System.Drawing.Size(624, 450);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void txtSpCd_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (Char.IsDigit((char)e.KeyValue))
			{
			}
			else
			{
				if (e.KeyCode == Keys.Back)
				{
				}
				else
				{
					e.Handled=true;	
				}
			}
		}

		private void txtConvert_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (Char.IsDigit((char)e.KeyValue))
			{
			}
			else
			{
				if (e.KeyCode == Keys.Back)
				{
				}
				else
				{
					e.Handled=true;	
				}
			}
		}

		private void cmbVariant_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			e.Handled=true;
		}

		private void txtSpCd_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
			
		}

		private void txtConvert_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
			}
			else if (e.KeyChar == '\b')
			{
			}
			else
			{
				e.Handled=true;
			}
		}
		private void val_data()
		{
		
            this.m_intError = 0;
			if (this.cmbVariant.Text.Trim().Length == 0)
			{
				MessageBox.Show("!!Select A Variant!!","FIA Biosum",
					             System.Windows.Forms.MessageBoxButtons.OK,
					             System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				this.cmbVariant.Focus();
				return;
			}

			if (this.txtSpCd.Text.Trim().Length == 0)
			{
				MessageBox.Show("!!Enter A Tree Species!!","FIA Biosum",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				this.txtSpCd.Focus();
				return;
            }
            else
            {
                //check to make sure the spcd exists in fia_tree_species_ref before trying to add
                this.m_ado.m_strSQL = "SELECT common_name FROM " + Tables.ProcessorScenarioRun.DefaultFiaTreeSpeciesRefTableName +
                                      " WHERE spcd = " + this.txtSpCd.Text.Trim() + ";";
                this.m_ado.SqlQueryReader(this.m_ado.m_OleDbConnection, this.m_ado.m_OleDbTransaction, this.m_ado.m_strSQL);
                if (!this.m_ado.m_OleDbDataReader.HasRows)
                {
                    string strMessage = "!!spcd " + this.txtSpCd.Text.Trim() + " is missing from the " +
                                        Tables.ProcessorScenarioRun.DefaultFiaTreeSpeciesRefTableName + " table. If this is a " +
                                        "valid FIA species code, please contact the Biosum Administrator to have it added!!";
                    System.Windows.Forms.MessageBox.Show(strMessage, "FIA Biosum");
                    this.m_intError = -1;
                    this.txtSpCd.Focus();
                    return;
                }
                this.m_ado.m_OleDbDataReader.Close();
			}

            if (this.cmbFvsSpCd.Text.Trim().Length == 0)
            {
                MessageBox.Show("!!Select An FVS Tree Species!!", "FIA Biosum",
                                 System.Windows.Forms.MessageBoxButtons.OK,
                                 System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                this.cmbFvsSpCd.Focus();
                return;
            }
            if (String.IsNullOrEmpty(txtCommon.Text))
            {
                MessageBox.Show("!!Enter a common_name!!", "FIA Biosum",
                 System.Windows.Forms.MessageBoxButtons.OK,
                 System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                this.txtCommon.Focus();
                return;
            }
            else
            {
                //check to make sure the spcd exists in fia_tree_species_ref before trying to add
                string strCommonName = m_ado.FixString(txtCommon.Text.Trim(), "'", "''");
                this.m_ado.m_strSQL = "SELECT spcd FROM " + m_strTreeSpcTable +
                                      " WHERE spcd <> " + this.txtSpCd.Text.Trim() + " AND TRIM(common_name) = '" +
                                      strCommonName + "';";
                this.m_ado.SqlQueryReader(this.m_ado.m_OleDbConnection, this.m_ado.m_OleDbTransaction, this.m_ado.m_strSQL);
                if (this.m_ado.m_OleDbDataReader.HasRows)
                {
                    this.m_ado.m_OleDbDataReader.Read();
                    string strMessage = "!!common_name '" + this.txtCommon.Text + "' is already in use for spcd " +
                                         Convert.ToString(this.m_ado.m_OleDbDataReader["spcd"]).Trim() + ". " +
                                         "common_name + spcd combinations must be unique!!";
                    System.Windows.Forms.MessageBox.Show(strMessage, "FIA Biosum");
                    this.m_intError = -1;
                    this.txtCommon.Focus();
                    return;
                }
                this.m_ado.m_OleDbDataReader.Close();
            }
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.Cancel;
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			this.val_data();
			if (this.m_intError==0)
				((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.OK;
		}

		private void cmbVariant_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.cmbVariant.Text.Trim().Length > 0)
			{
				this.m_strVariant = this.cmbVariant.Text.Trim().Substring(0,2).ToString();

				if (m_strVariant.Trim().Length > 0)
				{

				    this.cmbFvsSpCd.Items.Clear();
                    // cache convertToSpcd
                    int intMyConvertToSpcd = -1;
                    if (m_intConvertToSpCd > 0)
                    {
                        intMyConvertToSpcd = m_intConvertToSpCd;
                    }
                    this.m_dictFvsCommonName.Clear();
					m_ado.m_strSQL = "SELECT  spcd,fvs_species,common_name,fvs_common_name " + 
						"FROM " + m_strFvsTreeSpcTable + " " + 
						"WHERE LEN(TRIM(fvs_species)) > 0 AND " + 
						"LEN(TRIM(common_name)) > 0 AND " + 
						"TRIM(fvs_variant) = '" + m_strVariant.Trim() +  "' order by spcd;";
			
					m_ado.SqlQueryReader(m_ado.m_OleDbConnection,m_ado.m_OleDbTransaction,m_ado.m_strSQL);
					if (m_ado.m_OleDbDataReader.HasRows)
					{
						while (m_ado.m_OleDbDataReader.Read())
						{
                            string strMySpCd = Convert.ToString(m_ado.m_OleDbDataReader["spcd"]);
                            this.cmbFvsSpCd.Items.Add(strMySpCd + " - " + Convert.ToString(m_ado.m_OleDbDataReader["common_name"]) + " - " + m_ado.m_OleDbDataReader["fvs_species"]);
                            if (!m_dictFvsCommonName.ContainsKey(strMySpCd))
                            {
                                m_dictFvsCommonName.Add(strMySpCd, Convert.ToString(m_ado.m_OleDbDataReader["fvs_common_name"]));
                            }
						}

                        if (intMyConvertToSpcd > 0)
                        {
                            this.intConvertToSpCd = intMyConvertToSpcd;
                        }
					}
					m_ado.m_OleDbDataReader.Close();
				}
			}
		}
	
		public string strId
		{
			set	{ this.lblID.Text = value; }
			get { return this.lblID.Text.ToString().Trim(); }
		}
		public string strSpCd
		{
			set	{ this.txtSpCd.Text = value; }
			get { return this.txtSpCd.Text.ToString(); }
		}
		public string strCommonName
		{
			set	{ this.txtCommon.Text = value; }
			get { return this.txtCommon.Text.ToString(); }
		}
		public string strVariant
		{
			set	
            {

                if (value != null)
                {
                    for (int i = 0; i < this.cmbVariant.Items.Count; i++)
                    {
                        if (value.Equals(this.cmbVariant.GetItemText(this.cmbVariant.Items[i]).Trim().Substring(0, 2).ToString()))
                        {
                            this.cmbVariant.SelectedIndex = i;
                            return;
                        }
                    }
                }
                else
                {
                    this.cmbVariant.SelectedIndex = -1;
                }
            }
			get { return this.cmbVariant.Text.Trim().Substring(0,2).ToString(); }
		}
        public int intConvertToSpCd
		{
			set	
            {
                m_intConvertToSpCd = value;
                this.cmbFvsSpCd.SelectedIndex = -1;
                if (value > 0)
                {
                    for (int i = 0; i < this.cmbFvsSpCd.Items.Count; i++)
                    {
                        string strFvsSpeciesText = this.cmbFvsSpCd.GetItemText(this.cmbFvsSpCd.Items[i]);
                        int intNextSpeciesCode = -1;
                        int intFirstDash = strFvsSpeciesText.IndexOf("-", 0);
                        Int32.TryParse(strFvsSpeciesText.Substring(0, intFirstDash - 1), out intNextSpeciesCode);
                        if (value == intNextSpeciesCode)
                        {
                            this.cmbFvsSpCd.SelectedIndex = i;
                            return;
                        }
                    }
                    // if we get this far, the intConvertToSpCd wasn't found in the list, set default
                    int intFiaSpcd = -1;
                    if (int.TryParse(txtSpCd.Text, out intFiaSpcd) == true)
                    {
                        if (intFiaSpcd < 300)
                        {
                            // Default to Other Softwood
                            m_intConvertToSpCd = 298;
                        }
                        else
                        {
                            // Default to Other Hardwood
                            m_intConvertToSpCd = 998;
                        }
                        for (int i = 0; i < this.cmbFvsSpCd.Items.Count; i++)
                        {
                            // Try to set the dropdown again
                            string strFvsSpeciesText = this.cmbFvsSpCd.GetItemText(this.cmbFvsSpCd.Items[i]);
                            int intNextSpeciesCode = -1;
                            int intFirstDash = strFvsSpeciesText.IndexOf("-", 0);
                            Int32.TryParse(strFvsSpeciesText.Substring(0, intFirstDash - 1), out intNextSpeciesCode);
                            if (m_intConvertToSpCd == intNextSpeciesCode)
                            {
                                this.cmbFvsSpCd.SelectedIndex = i;
                                return;
                            }
                        }
                    }
                }
            }
			get { return m_intConvertToSpCd; }
		}
		public string strTreeSpeciesGenus
		{
			set	{ this.txtGenus.Text  = value; }
			get { return this.txtGenus.Text.ToString(); }
		}
		public string strTreeSpecies
		{
			set	{ this.txtSpecies.Text  = value; }
			get { return this.txtSpecies.Text.ToString(); }
		}
        public string strFvsCommonName
        {
            set { this.lblFvsCommonName.Text = value; }
            get { return this.lblFvsCommonName.Text.ToString(); }
        }

        public string strFvsSpeciesCode
        {
            get
            {
                return m_strFvsSpeciesCode;
            }
        }

        private void cmbFvsSpCd_SelectedIndexChanged(object sender, EventArgs e)
        {
            m_strFvsSpeciesCode = "";
            m_intConvertToSpCd = -1;
            lblFvsCommonName.Text = "";
            if (this.cmbFvsSpCd.Text.Trim().Length > 0)
            {
                int intFirstDash = this.cmbFvsSpCd.Text.IndexOf("-", 0);
                Int32.TryParse(this.cmbFvsSpCd.Text.Trim().Substring(0, intFirstDash -1), out m_intConvertToSpCd);
                int intSecondDash = this.cmbFvsSpCd.Text.IndexOf("-", intFirstDash + 1);
                m_strFvsSpeciesCode = this.cmbFvsSpCd.Text.Substring(intSecondDash + 1, this.cmbFvsSpCd.Text.Trim().Length - intSecondDash - 1).Trim();
                if (m_dictFvsCommonName.ContainsKey(m_intConvertToSpCd.ToString()))
                {
                    lblFvsCommonName.Text = m_dictFvsCommonName[m_intConvertToSpCd.ToString()];
                }
            }
        }

        private void txtSpCd_Validating(object sender,
                                        System.ComponentModel.CancelEventArgs e)
        {
            intConvertToSpCd = -1;
            // Set the convertToSpCd to same as spcd user entered, if possible
            if (txtSpCd.Text.Length > 0)
            {
                bool bIsInt = false;
                int intMySpCd = -1;
                bIsInt = Int32.TryParse(txtSpCd.Text, out intMySpCd);
                if (bIsInt == true)
                {
                    intConvertToSpCd = intMySpCd;
                }
            }
        }

	}
}

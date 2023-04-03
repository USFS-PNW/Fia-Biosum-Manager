using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_rx_edit.
	/// </summary>
	public class uc_scenario_harvest_cost_column_edit : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		public System.Windows.Forms.TextBox txtDesc;
		public System.Windows.Forms.ComboBox cmbCol;
		private string _strColumnList="";
		private string _strEditType="";
		private FIA_Biosum_Manager.utils m_oUtils = new utils();
		private string[] m_strColumnArray;
		private string _strCurrentSelectedColumnList="";
		private string[] m_strCurrentSelectedColumnArray;
		private string _strHarvestCostTableColumnList="";
		private string[] m_strHarvestCostTableColumnArray;
		public System.Windows.Forms.Label lblEdit;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_scenario_harvest_cost_column_edit()
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
            this.lblEdit = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbCol = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtDesc = new System.Windows.Forms.TextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblEdit);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.cmbCol);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtDesc);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnOK);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(520, 376);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // lblEdit
            // 
            this.lblEdit.Location = new System.Drawing.Point(152, 59);
            this.lblEdit.Name = "lblEdit";
            this.lblEdit.Size = new System.Drawing.Size(329, 25);
            this.lblEdit.TabIndex = 2;
            this.lblEdit.Text = "Select or enter component name (up to 8 chars)";
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(18, 87);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(128, 32);
            this.label2.TabIndex = 3;
            this.label2.Text = "Component";
            // 
            // cmbCol
            // 
            this.cmbCol.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbCol.Location = new System.Drawing.Point(152, 87);
            this.cmbCol.MaxLength = 8;
            this.cmbCol.Name = "cmbCol";
            this.cmbCol.Size = new System.Drawing.Size(352, 33);
            this.cmbCol.TabIndex = 4;
            this.cmbCol.Validating += new System.ComponentModel.CancelEventHandler(this.cmbCol_Validating);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(18, 184);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(128, 113);
            this.label1.TabIndex = 5;
            this.label1.Text = "Description of cost represented by this component";
            // 
            // txtDesc
            // 
            this.txtDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDesc.Location = new System.Drawing.Point(152, 199);
            this.txtDesc.Multiline = true;
            this.txtDesc.Name = "txtDesc";
            this.txtDesc.Size = new System.Drawing.Size(344, 72);
            this.txtDesc.TabIndex = 6;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(264, 303);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(88, 48);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(176, 303);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(88, 48);
            this.btnOK.TabIndex = 7;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(514, 32);
            this.lblTitle.TabIndex = 1;
            this.lblTitle.Text = "Additional CPA Component";
            // 
            // uc_scenario_harvest_cost_column_edit
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_scenario_harvest_cost_column_edit";
            this.Size = new System.Drawing.Size(520, 376);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		public void loadvalues()
		{
			int x=0;
			int y=0;

			this.cmbCol.Items.Clear();
			if (EditType.Trim().ToUpper()=="MODIFY")
			{
				cmbCol.Enabled=false;
			}
			else
			{
				if (this._strCurrentSelectedColumnList.Trim().Length > 0)
				{
					this.m_strCurrentSelectedColumnArray = m_oUtils.ConvertListToArray(this._strCurrentSelectedColumnList,",");
				}
				if (this._strHarvestCostTableColumnList.Trim().Length > 0)
				{
					this.m_strHarvestCostTableColumnArray = m_oUtils.ConvertListToArray(this._strHarvestCostTableColumnList,",");
				}
				if (_strColumnList.Trim().Length > 0)
				{
					m_strColumnArray = m_oUtils.ConvertListToArray(this._strColumnList,",");

					for (x=0;x<=m_strColumnArray.Length-1;x++)
					{
						if (this._strCurrentSelectedColumnList.Trim().Length > 0)
						{
							for (y=0;y<=this.m_strCurrentSelectedColumnArray.Length - 1;y++)
							{
								if (this.m_strColumnArray[x].Trim().ToUpper() == 
									this.m_strCurrentSelectedColumnArray[y].Trim().ToUpper()) break;
							}
							if (y > this.m_strCurrentSelectedColumnArray.Length - 1)
								this.cmbCol.Items.Add(m_strColumnArray[x]);
						}
						else
							cmbCol.Items.Add(m_strColumnArray[x]);
					}
				}
                if (m_strHarvestCostTableColumnArray != null)
                {
                    for (y = 0; y <= this.m_strHarvestCostTableColumnArray.Length - 1; y++)
                    {

                        for (x = 0; x <= cmbCol.Items.Count - 1; x++)
                        {
                            if (this.m_strHarvestCostTableColumnArray[y].Trim().ToUpper() ==
                                cmbCol.Items[x].ToString().Trim().ToUpper())
                            {
                                break;
                            }

                        }
                        if (x > cmbCol.Items.Count - 1)
                        {
                            if (this._strCurrentSelectedColumnList.Trim().Length > 0)
                            {
                                for (x = 0; x <= this.m_strCurrentSelectedColumnArray.Length - 1; x++)
                                {
                                    if (m_strHarvestCostTableColumnArray[y].Trim().ToUpper() ==
                                        this.m_strCurrentSelectedColumnArray[x].Trim().ToUpper()) break;
                                }
                                if (x > this.m_strCurrentSelectedColumnArray.Length - 1)
                                    cmbCol.Items.Add(this.m_strHarvestCostTableColumnArray[y].Trim());
                            }
                            else
                                cmbCol.Items.Add(this.m_strHarvestCostTableColumnArray[y].Trim());
                        }
                    }
                }
				
			}
			
		}
		public string EditType
		{
			set {_strEditType=value;}
			get {return _strEditType;}
		}
		public string ColumnText
		{
			set {this.cmbCol.Text = value;}
			get {return this.cmbCol.Text;}
		}
		public string ColumnList
		{
			set {_strColumnList = value;}
			get {return _strColumnList;}
		}
		public string CurrentSelectedColumnList
		{
			set {this._strCurrentSelectedColumnList = value;}
			get {return this._strCurrentSelectedColumnList;}
		}
		public string HarvestCostTableColumnList
		{
			set {_strHarvestCostTableColumnList=value;}
			get {return  _strHarvestCostTableColumnList;}
		}
		public string ColumnDescription
		{
			set {this.txtDesc.Text=value;}
			get {return this.txtDesc.Text;}
		}
		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.Cancel;
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			int y;
			int x;
			if (this.cmbCol.Text.Trim().Length == 0)
			{
				MessageBox.Show("!!Enter Column Name!!","FIA Biosum");
				return;
			}

			if (EditType.Trim().ToUpper() == "NEW")
			{
				
				string strCol = ColumnText.Trim();
				//check for invalid characters
				if (strCol.IndexOf("]",0) >=0 || strCol.IndexOf("[",0) >=0 ||
					strCol.Substring(0,1)==" " || strCol.IndexOf("!",0) >=0  ||
					strCol.IndexOf(".",0) >=0 )
				{
					MessageBox.Show("(.)([])(!) and leading space are not allowed in the column name","FIA Biosum");
					return;
				}
				//remove spaces from the column name
				strCol=strCol.Replace(" ","_");

				if (m_strColumnArray!=null)
				

				for (x=0;x<=m_strColumnArray.Length -1;x++)
				{
					//make sure this is a unique column
					if (this._strCurrentSelectedColumnList.Trim().Length > 0)
					{
						for (y=0;y<=this.m_strCurrentSelectedColumnArray.Length - 1;y++)
						{
					
							if (this.m_strColumnArray[x].Trim().ToUpper() == 
								this.m_strCurrentSelectedColumnArray[y].Trim().ToUpper()) break;
						}
						if (y > this.m_strCurrentSelectedColumnArray.Length - 1)
							this.cmbCol.Items.Add(m_strColumnArray[x]);
					}
					else
						this.cmbCol.Items.Add(m_strColumnArray[x]);

				}

				//make sure column is not currently selected
				ColumnText = strCol;

			}

			((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.OK;
		}

        // Prevent invalid characters from column name
        private void cmbCol_Validating(object sender, CancelEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            if (String.IsNullOrEmpty(comboBox.Text))
            {
                return;
            }
            System.Text.RegularExpressions.Regex rx =
                new System.Text.RegularExpressions.Regex("^[a-zA-Z][a-zA-Z0-9_]*$");

            System.Text.RegularExpressions.MatchCollection matches = rx.Matches(comboBox.Text);
            if (matches.Count < 1)
            {
                MessageBox.Show("The component name contains an invalid character. Only letters, numbers and underscores are permitted. Numbers and underscores are not permitted in the first position.", "FIA Biosum");
                e.Cancel = true;
            }
        }

	}
}

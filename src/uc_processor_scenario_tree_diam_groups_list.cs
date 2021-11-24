using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_processor_scenario_tree_diam_groups_list.
	/// </summary>
	public class uc_processor_scenario_tree_diam_groups_list : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnEdit;
		private System.Windows.Forms.Button btnNew;
		public int m_intDialogHt;
		public int m_intDialogWd;
		private System.Windows.Forms.ListView lstTreeDiam;
		private int m_intError;
        private string m_strError;
		private System.Windows.Forms.Button btnDefault;
		public System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnClear;
		private System.Windows.Forms.Button btnHelp;
		private FIA_Biosum_Manager.ado_data_access m_ado;
		private System.Windows.Forms.Button btnDelete;

        // scenario-specific variables
        private string _strScenarioId = "";
        private frmProcessorScenario _frmProcessorScenario = null;
        // Help system variables
        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultProcessorXPSFile;
        private Button btnSelectFile;
        private TextBox txtImportFile;
        private TextBox textBox1;
        private Button BtnImport;


		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_processor_scenario_tree_diam_groups_list()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_intDialogHt = this.groupBox1.Top + this.btnClose.Top + this.btnClose.Height + 20;
			this.m_intDialogWd = this.groupBox1.Left + this.btnClose.Left + this.btnClose.Width + 20;
            this.m_oEnv = new env();

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

        public frmProcessorScenario ReferenceProcessorScenarioForm
        {
            get { return this._frmProcessorScenario; }
            set { this._frmProcessorScenario = value; }
        }
        public string ScenarioId
        {
            get { return _strScenarioId; }
            set { _strScenarioId = value; }
        }

        public void loadvalues()
        {
            string strId = "";
            string strMin = "";
            string strMax = "";
            string strDef = "";

            this.lstTreeDiam.Clear();
            this.lstTreeDiam.Columns.Add("Group ID", 60, HorizontalAlignment.Left);
            this.lstTreeDiam.Columns.Add("Minimum Diameter", 150, HorizontalAlignment.Left);
            this.lstTreeDiam.Columns.Add("Maximum Diameter", 150, HorizontalAlignment.Left);
            this.lstTreeDiam.Columns.Add("Definition", 200, HorizontalAlignment.Left);

            this._strScenarioId = this.ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.ScenarioId; 
            if (ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection != null)
            {
                try
                {
                    //load up each row from SCENARIO_TREE_DIAM_GROUPS table
                    int x;
                    for (x = 0; x <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Count - 1; x++)
                    {
                        ProcessorScenarioItem.TreeDiamGroupsItem p_oTreeGroupsItem = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Item(x);

                        strId = "";
                        strMax = "";
                        strMin = "";
                        strDef = "";

                        //make sure the row is not null values
                        if (p_oTreeGroupsItem.DiamClass.Trim().Length > 0)
                        {
                            strId = p_oTreeGroupsItem.DiamGroup;
                            strMin = p_oTreeGroupsItem.MinDiam;
                            strMax = p_oTreeGroupsItem.MaxDiam;
                            strDef = p_oTreeGroupsItem.DiamClass;
                            this.lstTreeDiam.BeginUpdate();
                            System.Windows.Forms.ListViewItem listItem = new ListViewItem();
                            listItem.Text = strId;
                            listItem.SubItems.Add(strMin);
                            listItem.SubItems.Add(strMax);
                            listItem.SubItems.Add(strDef);
                            this.lstTreeDiam.Items.Add(listItem);
                            this.lstTreeDiam.EndUpdate();
                        }

                    }
                    if (this.lstTreeDiam.Items.Count > 0)
                    {
                        if (this.lstTreeDiam.SelectedItems.Count == 0)
                        {
                            this.lstTreeDiam.Items[this.lstTreeDiam.Items.Count - 1].Selected = true;
                        }
                    }
                    ((frmDialog)this.ParentForm).Enabled = true;
                }
                catch (Exception caught)
                {
                    this.m_intError = -1;
                    MessageBox.Show(caught.Message);
                }
            }
        }

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_processor_scenario_tree_diam_groups_list));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.BtnImport = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.txtImportFile = new System.Windows.Forms.TextBox();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnHelp = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnDefault = new System.Windows.Forms.Button();
            this.btnNew = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.lstTreeDiam = new System.Windows.Forms.ListView();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.BtnImport);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.btnSelectFile);
            this.groupBox1.Controls.Add(this.txtImportFile);
            this.groupBox1.Controls.Add(this.btnDelete);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.btnClear);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnSave);
            this.groupBox1.Controls.Add(this.btnDefault);
            this.groupBox1.Controls.Add(this.btnNew);
            this.groupBox1.Controls.Add(this.btnEdit);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.lstTreeDiam);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(630, 346);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // BtnImport
            // 
            this.BtnImport.Enabled = false;
            this.BtnImport.Location = new System.Drawing.Point(549, 258);
            this.BtnImport.Name = "BtnImport";
            this.BtnImport.Size = new System.Drawing.Size(64, 32);
            this.BtnImport.TabIndex = 14;
            this.BtnImport.Text = "Import";
            this.BtnImport.Click += new System.EventHandler(this.BtnImport_Click);
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Menu;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(17, 267);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(155, 20);
            this.textBox1.TabIndex = 13;
            this.textBox1.Text = "Import groups from file";
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Image = ((System.Drawing.Image)(resources.GetObject("btnSelectFile.Image")));
            this.btnSelectFile.Location = new System.Drawing.Point(511, 258);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(32, 32);
            this.btnSelectFile.TabIndex = 12;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // txtImportFile
            // 
            this.txtImportFile.Enabled = false;
            this.txtImportFile.Location = new System.Drawing.Point(177, 265);
            this.txtImportFile.Name = "txtImportFile";
            this.txtImportFile.Size = new System.Drawing.Size(328, 20);
            this.txtImportFile.TabIndex = 11;
            // 
            // btnDelete
            // 
            this.btnDelete.Enabled = false;
            this.btnDelete.Location = new System.Drawing.Point(286, 216);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(64, 32);
            this.btnDelete.TabIndex = 10;
            this.btnDelete.Text = "Delete";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(8, 304);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(64, 32);
            this.btnHelp.TabIndex = 8;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(350, 216);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(64, 32);
            this.btnClear.TabIndex = 7;
            this.btnClear.Text = "Clear All";
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(478, 216);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(64, 32);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Enabled = false;
            this.btnSave.Location = new System.Drawing.Point(414, 216);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(64, 32);
            this.btnSave.TabIndex = 5;
            this.btnSave.Text = "Save";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnDefault
            // 
            this.btnDefault.Location = new System.Drawing.Point(28, 216);
            this.btnDefault.Name = "btnDefault";
            this.btnDefault.Size = new System.Drawing.Size(114, 32);
            this.btnDefault.TabIndex = 2;
            this.btnDefault.Text = "Use Default Values";
            this.btnDefault.Click += new System.EventHandler(this.btnDefault_Click);
            // 
            // btnNew
            // 
            this.btnNew.Location = new System.Drawing.Point(158, 216);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(64, 32);
            this.btnNew.TabIndex = 3;
            this.btnNew.Text = "New";
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(222, 216);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(64, 32);
            this.btnEdit.TabIndex = 4;
            this.btnEdit.Text = "Edit";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(517, 304);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 9;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lstTreeDiam
            // 
            this.lstTreeDiam.FullRowSelect = true;
            this.lstTreeDiam.GridLines = true;
            this.lstTreeDiam.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstTreeDiam.HideSelection = false;
            this.lstTreeDiam.Location = new System.Drawing.Point(16, 48);
            this.lstTreeDiam.MultiSelect = false;
            this.lstTreeDiam.Name = "lstTreeDiam";
            this.lstTreeDiam.Size = new System.Drawing.Size(536, 160);
            this.lstTreeDiam.TabIndex = 1;
            this.lstTreeDiam.UseCompatibleStateImageBehavior = false;
            this.lstTreeDiam.View = System.Windows.Forms.View.Details;
            this.lstTreeDiam.SelectedIndexChanged += new System.EventHandler(this.lstTreeDiam_SelectedIndexChanged);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(624, 24);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "Tree Diameter Groups List";
            // 
            // uc_processor_scenario_tree_diam_groups_list
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_scenario_tree_diam_groups_list";
            this.Size = new System.Drawing.Size(630, 346);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

        private void btnClose_Click(object sender, System.EventArgs e)
        {
            if (this.btnSave.Enabled == true)
            {
                DialogResult result = MessageBox.Show("Save Changes Y/N", "Tree Diameter Groups", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    this.savevalues();
                    // Force reload of components that use tree groups since they changed in db
                    ReferenceProcessorScenarioForm.m_bTreeGroupsFirstTime = true;
                }
            }
            this.ParentForm.Close();
            this.ParentForm.Dispose();
        }

		private void btnNew_Click(object sender, System.EventArgs e)
		{
			string strMax="";
			string strMin="";
			string strId = "";
            double dblMin = 0;

			//get the last id if there is one
			if (this.lstTreeDiam.Items.Count > 0)
			{
				strId = Convert.ToString(Convert.ToInt16(this.lstTreeDiam.Items[this.lstTreeDiam.Items.Count-1].Text) + 1);
				strMin = this.lstTreeDiam.Items[this.lstTreeDiam.Items.Count-1].SubItems[2].Text;
				if (strMin.IndexOf(".") < 0) strMin += ".0";

                dblMin = Convert.ToDouble(strMin) + 0.1;
                strMin = Convert.ToString(dblMin);
				
				strMax = strMin;
				//if (strMin.IndexOf(".") < 0) strMin += ".0";

			}
			else
			{
				strMax="0.0";
				strMin="0.0";
				strId = "1";
			}

			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
		    frmTemp.MaximizeBox = false;
			frmTemp.BackColor = System.Drawing.SystemColors.Control;
			frmTemp.Text = "Database: Tree Diameter Groups (New)";
			frmTemp.Initialize_Plot_Tree_Diam_Edit_User_Control();

			frmTemp.uc_processor_scenario_tree_diam_groups_edit1.txtGroupId = strId;
			frmTemp.uc_processor_scenario_tree_diam_groups_edit1.txtMaximumDiameter = strMax;
			frmTemp.uc_processor_scenario_tree_diam_groups_edit1.txtMinimumDiameter = strMin;


			//frmTemp.MdiParent = this;
			//frmTemp.Initialize_Plot_Tree_Diam_User_Control();


			frmTemp.Height=0;
			frmTemp.Width=0;
			if (frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Top + frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Height > frmTemp.ClientSize.Height + 2)
			{
				for (int x=1;;x++)
				{
					frmTemp.Height = x;
					if (frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Top + 
						frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Height < 
						frmTemp.ClientSize.Height)
					{
						break;
					}
				}

			}
			if (frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Left + frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Width > frmTemp.ClientSize.Width + 2)
			{
				for (int x=1;;x++)
				{
					frmTemp.Width = x;
					if (frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Left + 
						frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Width < 
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
            System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
			if (result==System.Windows.Forms.DialogResult.OK)
			{
				strMax = frmTemp.uc_processor_scenario_tree_diam_groups_edit1.txtMaximumDiameter.Trim();
				//remove the period if it is the last character
				if (strMax.IndexOf(".") == strMax.Length-1)
				{
					strMax = strMax.Replace(".","");
				}
				strMin = frmTemp.uc_processor_scenario_tree_diam_groups_edit1.txtMinimumDiameter.Trim();
				//remove the period if it is the last character
				if (strMin.IndexOf(".") == strMin.Length-1)
				{
					strMin = strMin.Replace(".","");
				}
				this.lstTreeDiam.BeginUpdate();
				System.Windows.Forms.ListViewItem listItem = new ListViewItem();
				listItem.Text=frmTemp.uc_processor_scenario_tree_diam_groups_edit1.txtGroupId;
				listItem.SubItems.Add(strMin);
				listItem.SubItems.Add(strMax);
				listItem.SubItems.Add(strMin + " - " + strMax);
				this.lstTreeDiam.Items.Add(listItem);
				this.lstTreeDiam.EndUpdate();
				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			}
			frmTemp.Close();
			frmTemp.Dispose();
            frmTemp=null;
		}

		private void btnDefault_Click(object sender, System.EventArgs e)
		{
			this.lstTreeDiam.Clear();
			this.lstTreeDiam.Columns.Add("Group ID", 60, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Minimum Diameter", 150, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Maximum Diameter", 150, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Definition", 200, HorizontalAlignment.Left);

			this.m_intError=0;
			this.lstTreeDiam.BeginUpdate();
			this.lstTreeDiam.Items.Add("1");
			this.lstTreeDiam.Items[0].SubItems.Add("1");
			this.lstTreeDiam.Items[0].SubItems.Add("7");
			this.lstTreeDiam.Items[0].SubItems.Add("1 - 7");

			this.lstTreeDiam.Items.Add("2");
			this.lstTreeDiam.Items[1].SubItems.Add("7.1");
			this.lstTreeDiam.Items[1].SubItems.Add("10");
			this.lstTreeDiam.Items[1].SubItems.Add("7.1 - 10");

			this.lstTreeDiam.Items.Add("3");
			this.lstTreeDiam.Items[2].SubItems.Add("10.1");
			this.lstTreeDiam.Items[2].SubItems.Add("16");
			this.lstTreeDiam.Items[2].SubItems.Add("10.1 - 16");

			this.lstTreeDiam.Items.Add("4");
			this.lstTreeDiam.Items[3].SubItems.Add("16.1");
			this.lstTreeDiam.Items[3].SubItems.Add("21");
			this.lstTreeDiam.Items[3].SubItems.Add("16.1 - 21");

			this.lstTreeDiam.Items.Add("5");
			this.lstTreeDiam.Items[4].SubItems.Add("21.1");
			this.lstTreeDiam.Items[4].SubItems.Add("999");
			this.lstTreeDiam.Items[4].SubItems.Add("21.1 - 999");




			this.lstTreeDiam.EndUpdate();
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;


		}

		private void btnClear_Click(object sender, System.EventArgs e)
		{
			if (this.lstTreeDiam.Items.Count > 0) this.btnSave.Enabled=true;
			this.lstTreeDiam.Clear();
			this.lstTreeDiam.Columns.Add("Group ID", 60, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Minimum Diameter", 150, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Maximum Diameter", 150, HorizontalAlignment.Left);
			this.lstTreeDiam.Columns.Add("Definition", 200, HorizontalAlignment.Left);
		}

		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			string strMax="";
			string strMin="";
			string strId = "";

			if (this.lstTreeDiam.SelectedItems.Count == 0)
				return;



			
			strId = this.lstTreeDiam.SelectedItems[0].Text;
			strMin = this.lstTreeDiam.SelectedItems[0].SubItems[1].Text;
            strMax = this.lstTreeDiam.SelectedItems[0].SubItems[2].Text;
  		    if (strMin.IndexOf(".") < 0) strMin += ".0";
			if (strMax.IndexOf(".") < 0) strMax += ".0";

			FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
			frmTemp.MaximizeBox = false;
			frmTemp.BackColor = System.Drawing.SystemColors.Control;
			frmTemp.Text = "Database: Tree Diameter Groups (Edit)";
			frmTemp.Initialize_Plot_Tree_Diam_Edit_User_Control();

			frmTemp.uc_processor_scenario_tree_diam_groups_edit1.txtGroupId = strId;
			frmTemp.uc_processor_scenario_tree_diam_groups_edit1.txtMaximumDiameter = strMax;
			frmTemp.uc_processor_scenario_tree_diam_groups_edit1.txtMinimumDiameter = strMin;


			frmTemp.Height=0;
			frmTemp.Width=0;
			if (frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Top + frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Height > frmTemp.ClientSize.Height + 2)
			{
				for (int x=1;;x++)
				{
					frmTemp.Height = x;
					if (frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Top + 
						frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Height < 
						frmTemp.ClientSize.Height)
					{
						break;
					}
				}

			}
			if (frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Left + frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Width > frmTemp.ClientSize.Width + 2)
			{
				for (int x=1;;x++)
				{
					frmTemp.Width = x;
					if (frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Left + 
						frmTemp.uc_processor_scenario_tree_diam_groups_edit1.Width < 
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
			System.Windows.Forms.DialogResult result = frmTemp.ShowDialog();
			if (result==System.Windows.Forms.DialogResult.OK)
			{
				strMax = frmTemp.uc_processor_scenario_tree_diam_groups_edit1.txtMaximumDiameter.Trim();
				//remove the period if it is the last character
				if (strMax.IndexOf(".") == strMax.Length-1)
				{
					strMax = strMax.Replace(".","");
				}
				strMin = frmTemp.uc_processor_scenario_tree_diam_groups_edit1.txtMinimumDiameter.Trim();
				//remove the period if it is the last character
				if (strMin.IndexOf(".") == strMin.Length-1)
				{
					strMin = strMin.Replace(".","");
				}
				this.lstTreeDiam.BeginUpdate();
				this.lstTreeDiam.SelectedItems[0].SubItems[1].Text = strMin;
				this.lstTreeDiam.SelectedItems[0].SubItems[2].Text = strMax;
				this.lstTreeDiam.SelectedItems[0].SubItems[3].Text = strMin + " - " + strMax;
				this.lstTreeDiam.EndUpdate();
				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			}
			frmTemp.Close();
			frmTemp.Dispose();
			frmTemp=null;

		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
		  this.savevalues();
          // Force reload of components that use tree groups since they changed in db
          ReferenceProcessorScenarioForm.m_bTreeGroupsFirstTime = true;
          // Copied values have been saved
          ReferenceProcessorScenarioForm.m_bTreeGroupsCopied = false;
		}
		private void val_data()
		{
			int x;
			//int y;
			double dblMin;
            double dblMax;
			string strId1;
			string strId2;
            this.m_intError=0;

			if (this.lstTreeDiam.Items.Count==0)
			{
				MessageBox.Show("No tree diameter groups to save","Tree Diameter Groups",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
				return;
			}

			for (x=0;x<=this.lstTreeDiam.Items.Count-2;x++)
			{
				//make sure the preceding max is less than the next minimum
			   dblMax=Convert.ToDouble(this.lstTreeDiam.Items[x].SubItems[2].Text);
			   dblMin=Convert.ToDouble(this.lstTreeDiam.Items[x+1].SubItems[1].Text);
				if (dblMin <= dblMax)
				{ 
					strId1 = this.lstTreeDiam.Items[x].Text;
                    strId2 = this.lstTreeDiam.Items[x+1].Text ;
					MessageBox.Show("Group ID " + strId2 + " minimum diameter must be greater than Group ID " + strId1 + " maximum diameter","Tree Diameter Groups",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					this.m_intError=-1;
					return;
				}
			}            
		}
		public void savevalues()
		{


			val_data();
			if (this.m_intError==0)
			{
				//string strSQL;
				int x;
                string strMin;
				string strMax;
				string strDef;
				string strId;

                if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
                {
                    //
                    //OPEN CONNECTION TO DB FILE CONTAINING PROCESSOR SCENARIO TABLES
                    //
                    //scenario mdb connection
                    string strScenarioMDB =
                        frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                        "\\processor\\db\\scenario_processor_rule_definitions.mdb";
                    ado_data_access oAdo = new ado_data_access();
                    oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB, "", ""));
                    if (oAdo.m_intError != 0)
                    {
                        m_intError = m_ado.m_intError;
                        m_strError = m_ado.m_strError;
                        oAdo = null;
                        return;
                    }

                    if (this.m_intError == 0)
                    {
                        //delete the current records
                        oAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName +
                            " WHERE TRIM(UCASE(scenario_id)) = '" + ScenarioId.Trim().ToUpper() + "'";
                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                        if (oAdo.m_intError == 0)
                        {
                            for (x = 0; x <= this.lstTreeDiam.Items.Count - 1; x++)
                            {
                                strId = this.lstTreeDiam.Items[x].Text;
                                strMin = this.lstTreeDiam.Items[x].SubItems[1].Text;
                                strMax = this.lstTreeDiam.Items[x].SubItems[2].Text;
                                strDef = this.lstTreeDiam.Items[x].SubItems[3].Text;

                                oAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName + " " +
                                    "(diam_group,diam_class,min_diam,max_diam,scenario_id) VALUES " +
                                    "(" + strId + ",'" + strDef.Trim() + "'," +
                                    strMin + "," + strMax + ",'" + ScenarioId + "');";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                if (oAdo.m_intError != 0)
                                {
                                    break;
                                }
                            }
                        }
                    }
                    oAdo.CloseConnection(oAdo.m_OleDbConnection);
                    this.m_intError = oAdo.m_intError;
                    oAdo = null;
                }
                else
                {
                    string strScenarioDB =
                        frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                        "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                    SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
                    using (System.Data.SQLite.SQLiteConnection conn =
                        new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strScenarioDB)))
                    {
                        conn.Open();
                        if (oDataMgr.m_intError != 0)
                        {
                            m_intError = oDataMgr.m_intError;
                            m_strError = oDataMgr.m_strError;
                            oDataMgr = null;
                            return;
                        }
                        if (this.m_intError == 0)
                        {
                            //delete the current records
                            oDataMgr.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName +
                                " WHERE TRIM(UPPER(scenario_id)) = '" + ScenarioId.Trim().ToUpper() + "'";
                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                            if (oDataMgr.m_intError == 0)
                            {
                                for (x = 0; x <= this.lstTreeDiam.Items.Count - 1; x++)
                                {
                                    strId = this.lstTreeDiam.Items[x].Text;
                                    strMin = this.lstTreeDiam.Items[x].SubItems[1].Text;
                                    strMax = this.lstTreeDiam.Items[x].SubItems[2].Text;
                                    strDef = this.lstTreeDiam.Items[x].SubItems[3].Text;
                                    oDataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName + " " +
                                        "(diam_group,diam_class,min_diam,max_diam,scenario_id) VALUES " +
                                        "(" + strId + ",'" + strDef.Trim() + "'," +
                                        strMin + "," + strMax + ",'" + ScenarioId + "');";
                                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                    if (oDataMgr.m_intError != 0)
                                    {
                                        break;
                                    }
                                }
                                this.m_intError = oDataMgr.m_intError;
                                oDataMgr = null;
                            }
                        }
                    }
                }
                if (this.m_intError == 0)
                {
                    this.btnSave.Enabled = false;
                }
            }
		}

		private void btnDelete_Click(object sender, System.EventArgs e)
		{
			this.lstTreeDiam.SelectedItems[0].Remove();
			if (this.lstTreeDiam.Items.Count > 0)
			{
				this.lstTreeDiam.Items[this.lstTreeDiam.Items.Count-1].Selected=true;
			}
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			this.lstTreeDiam.Focus();
		}

		private void lstTreeDiam_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.lstTreeDiam.SelectedItems.Count > 0)
			{
				if (this.lstTreeDiam.SelectedItems[0].Index ==
					this.lstTreeDiam.Items.Count-1)
				{
					if (this.btnDelete.Enabled==false) this.btnDelete.Enabled=true;
				}
				else
				{
					if (this.btnDelete.Enabled==true) this.btnDelete.Enabled=false;
				}
		
			}
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{

			this.ParentForm.Close();
			this.ParentForm.Dispose();
			
		}

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "PROCESSOR", "TREE_DIAMETER_GROUPS" });
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            var OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Title = "Text File With Maximum Diameter Group values";
            OpenFileDialog1.Filter = "Text File (*.TXT) |*.txt";
            var result = OpenFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                if (OpenFileDialog1.FileName.Trim().Length > 0)
                {
                    txtImportFile.Text = OpenFileDialog1.FileName.Trim();
                    BtnImport.Enabled = true;
                }
                OpenFileDialog1 = null;
            }
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(txtImportFile.Text))
            {
                DialogResult result = MessageBox.Show("Importing tree diameter groups will overwrite your current groups. Do you wish to continue ?", 
                    "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                if (result != System.Windows.Forms.DialogResult.Yes)
                {
                    return;
                }
                //Open the file with a stream reader.
                System.Collections.Generic.IList<String> lstRows = new System.Collections.Generic.List<String>();
                using (System.IO.StreamReader s = new System.IO.StreamReader(txtImportFile.Text, System.Text.Encoding.Default))
                {
                    string strNextLine = null;
                    double dblHighestValue = -99;
                    while ((strNextLine = s.ReadLine()) != null)
                    {
                        double dblNextValue = -1;
                        bool bIsNumeric = Double.TryParse(strNextLine, out dblNextValue);
                        if (bIsNumeric == false)
                        {
                            MessageBox.Show("!! All group maximum diameter threshold values must be numeric !!", "FIA Biosum",
                                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            return;
                        }
                        if (dblNextValue <= dblHighestValue)
                        {
                            MessageBox.Show("!! Group maximum diameter threshold values must be unique and in ascending order !!", "FIA Biosum",
                                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            return;
                        }
                        dblNextValue = Math.Round(dblNextValue, 1); //Round to the nearest tenth; We are precise only to that degree
                        lstRows.Add(Convert.ToString(dblNextValue));
                        dblHighestValue = dblNextValue;
                    }
                }
                if (lstRows.Count > 0)
                {
                    btnClear.PerformClick();
                    this.lstTreeDiam.BeginUpdate();
                    int intGroup = 1;
                    double dblMinimum = 1.0;
                    double dblMaximum = -1.0;
                    foreach (string strRow in lstRows)
                    {
                        if (!String.IsNullOrEmpty(strRow))
                        {
                            this.lstTreeDiam.Items.Add(Convert.ToString(intGroup));
                            this.lstTreeDiam.Items[intGroup - 1].SubItems.Add(Convert.ToString(dblMinimum));
                            this.lstTreeDiam.Items[intGroup - 1].SubItems.Add(strRow);
                            this.lstTreeDiam.Items[intGroup - 1].SubItems.Add(dblMinimum + " - " + strRow);
                            dblMaximum = Convert.ToDouble(strRow);
                            dblMinimum = dblMaximum + 0.1;
                            intGroup++;
                        }
                    }
                    this.lstTreeDiam.EndUpdate();
                    txtImportFile.Text = "";
                    BtnImport.Enabled = false;
                    if (this.btnSave.Enabled == false) this.btnSave.Enabled = true;
                    MessageBox.Show("!! Tree diameter group values have been imported. Click the save button to keep them !!", "FIA Biosum",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("!! The import file did not have any values to import !!", "FIA Biosum",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("!! A file must be selected to import groups !!", "FIA Biosum", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

            }
        }
	}
}

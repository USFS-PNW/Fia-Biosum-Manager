using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using SQLite.ADO;


namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_plot_add_edit.
	/// </summary>
	public class uc_plot_add_edit : System.Windows.Forms.UserControl
	{
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnEdit;
        private System.Windows.Forms.ToolBar tlbPlotAddEdit;
        private System.Windows.Forms.ToolBarButton tlbbtnAdd;
        //private int m_intError=0;
        public const int TABLETYPE = 0;
        public const int PATH = 1;
        public const int MDBFILE = 2;
        public const int FILESTATUS = 3;
        public const int TABLE = 4;
        public const int TABLESTATUS = 5;
        public const int RECORDCOUNT = 6;
        private System.Windows.Forms.ImageList imageList1;
        private ToolBarButton tblbtnDeleteConds;
        private ToolBarButton tlbbtnHelp;
        private System.ComponentModel.IContainer components;
        private env m_oEnv;
        private Help m_oHelp;
        private ToolBarButton tblbtnDeletePackages;
        private string m_xpsFile = Help.DefaultDatabaseXPSFile;


		public uc_plot_add_edit()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			
			// TODO: Add any initialization after the InitializeComponent call
            this.m_oEnv = new env();
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_plot_add_edit));
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.tlbPlotAddEdit = new System.Windows.Forms.ToolBar();
            this.tlbbtnAdd = new System.Windows.Forms.ToolBarButton();
            this.tblbtnDeleteConds = new System.Windows.Forms.ToolBarButton();
            this.tblbtnDeletePackages = new System.Windows.Forms.ToolBarButton();
            this.tlbbtnHelp = new System.Windows.Forms.ToolBarButton();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.SuspendLayout();
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(24, 72);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(80, 72);
            this.btnAdd.TabIndex = 0;
            this.btnAdd.Text = "Add Plot Data";
            this.btnAdd.Visible = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(168, 72);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(80, 72);
            this.btnEdit.TabIndex = 1;
            this.btnEdit.Text = "Edit Plot Data";
            this.btnEdit.Visible = false;
            // 
            // tlbPlotAddEdit
            // 
            this.tlbPlotAddEdit.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
            this.tlbbtnAdd,
            this.tblbtnDeleteConds,
            this.tblbtnDeletePackages,
            this.tlbbtnHelp});
            this.tlbPlotAddEdit.ButtonSize = new System.Drawing.Size(150, 55);
            this.tlbPlotAddEdit.Divider = false;
            this.tlbPlotAddEdit.Dock = System.Windows.Forms.DockStyle.None;
            this.tlbPlotAddEdit.DropDownArrows = true;
            this.tlbPlotAddEdit.ImageList = this.imageList1;
            this.tlbPlotAddEdit.Location = new System.Drawing.Point(5, 5);
            this.tlbPlotAddEdit.Name = "tlbPlotAddEdit";
            this.tlbPlotAddEdit.ShowToolTips = true;
            this.tlbPlotAddEdit.Size = new System.Drawing.Size(610, 62);
            this.tlbPlotAddEdit.TabIndex = 2;
            this.tlbPlotAddEdit.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.tlbPlotAddEdit_ButtonClick);
            // 
            // tlbbtnAdd
            // 
            this.tlbbtnAdd.ImageIndex = 0;
            this.tlbbtnAdd.Name = "tlbbtnAdd";
            this.tlbbtnAdd.Text = "Add Plot Data";
            // 
            // tblbtnDeleteConds
            // 
            this.tblbtnDeleteConds.ImageIndex = 1;
            this.tblbtnDeleteConds.Name = "tblbtnDeleteConds";
            this.tblbtnDeleteConds.Text = "Delete Conditions";
            // 
            // tblbtnDeletePackages
            // 
            this.tblbtnDeletePackages.ImageIndex = 1;
            this.tblbtnDeletePackages.Name = "tblbtnDeletePackages";
            this.tblbtnDeletePackages.Text = "Delete Packages";
            // 
            // tlbbtnHelp
            // 
            this.tlbbtnHelp.ImageIndex = 2;
            this.tlbbtnHelp.Name = "tlbbtnHelp";
            this.tlbbtnHelp.Text = "Help";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "");
            this.imageList1.Images.SetKeyName(1, "");
            this.imageList1.Images.SetKeyName(2, "HelpSystemBlue32.png");
            // 
            // uc_plot_add_edit
            // 
            this.Controls.Add(this.tlbPlotAddEdit);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.btnAdd);
            this.Name = "uc_plot_add_edit";
            this.Size = new System.Drawing.Size(615, 72);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		private void btnAdd_Click(object sender, System.EventArgs e)
		{
			frmDialog frmTemp = new frmDialog(((frmDialog)this.ParentForm).m_frmMain);
			frmTemp.Visible=false;
			//FIA_Biosum_Manager.uc_plot_input uc_plot_input1 = new uc_plot_input();
			//frmTemp.Controls.Add(uc_plot_input1);
			frmTemp.Initialize_Plot_Input_User_Control();
			frmTemp.MaximizeBox = false;
			frmTemp.MinimizeBox = false;
			frmTemp.Width = frmTemp.uc_plot_input1.m_DialogWd;
			frmTemp.Height = frmTemp.uc_plot_input1.m_DialogHt;
			frmTemp.Text = "Database: Add Plot Data";
			frmTemp.uc_plot_input1.Dock = System.Windows.Forms.DockStyle.Fill;
			frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			frmTemp.uc_plot_input1.Visible=true;
			frmTemp.DisposeOfFormWhenClosing=true;
            frmTemp.MinimizeMainForm = true;
            frmTemp.ParentControl = frmMain.g_oFrmMain;
            frmTemp.ParentControl.Enabled = false;
            frmTemp.Show();
		}

		private void tlbPlotAddEdit_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			switch (e.Button.Text.Trim().ToUpper())
			{
				case "ADD PLOT DATA":
					frmDialog frmTemp = new frmDialog(((frmDialog)this.ParentForm).m_frmMain);
					frmTemp.Visible=false;
					frmTemp.Initialize_Plot_Input_User_Control();
					frmTemp.MaximizeBox = false;
					frmTemp.MinimizeBox = true;
					frmTemp.Width = frmTemp.uc_plot_input1.m_DialogWd;
					frmTemp.Height = frmTemp.uc_plot_input1.m_DialogHt;
					frmTemp.Text = "Database: Add Plot Data";
					frmTemp.uc_plot_input1.Dock = System.Windows.Forms.DockStyle.Fill;
					frmTemp.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
					frmTemp.uc_plot_input1.Visible=true;
					frmTemp.DisposeOfFormWhenClosing=true;
                    frmTemp.uc_plot_input1.ReferenceFormDialog = frmTemp;
                    frmTemp.MinimizeMainForm = true;
                    frmTemp.ParentControl = frmMain.g_oFrmMain;
                    frmTemp.ParentControl.Enabled = false;
					frmTemp.Show();
					break;

				case "DELETE CONDITIONS":
					frmDialog frmTemp2 = new frmDialog(((frmDialog)this.ParentForm).m_frmMain);
					frmTemp2.Visible=false;
					frmTemp2.Initialize_Delete_Conditions_User_Control();
					frmTemp2.MaximizeBox = false;
					frmTemp2.MinimizeBox = true;
					frmTemp2.Width = frmTemp2.uc_delete_conditions.m_DialogWd;
					frmTemp2.Height = frmTemp2.uc_delete_conditions.m_DialogHt;
					frmTemp2.Text = "Database: Delete Conditions";
					frmTemp2.uc_delete_conditions.Dock = System.Windows.Forms.DockStyle.Fill;
					frmTemp2.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
					frmTemp2.uc_delete_conditions.Visible=true;
					frmTemp2.DisposeOfFormWhenClosing=true;
                    frmTemp2.uc_delete_conditions.ReferenceFormDialog = frmTemp2;
                    frmTemp2.MinimizeMainForm = true;
                    frmTemp2.ParentControl = frmMain.g_oFrmMain;
                    frmTemp2.ParentControl.Enabled = false;
					frmTemp2.Show();
					break;

                case "DELETE PACKAGES":
                    frmDialog frmTemp3 = new frmDialog(((frmDialog) this.ParentForm).m_frmMain);
                    frmTemp3.Visible = false;
                    frmTemp3.Initialize_Delete_Packages_User_Control();
                    frmTemp3.MaximizeBox = false;
                    frmTemp3.MinimizeBox = true;
                    frmTemp3.Width = frmTemp3.uc_delete_packages.m_DialogWd;
                    frmTemp3.Height = frmTemp3.uc_delete_packages.m_DialogHt;
                    frmTemp3.Text = "Database: Delete Packages";
                    frmTemp3.uc_delete_packages.Dock = System.Windows.Forms.DockStyle.Fill;
                    frmTemp3.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                    frmTemp3.uc_delete_packages.Visible = true;
                    frmTemp3.DisposeOfFormWhenClosing = true;
                    frmTemp3.uc_delete_packages.ReferenceFormDialog = frmTemp3;
                    frmTemp3.MinimizeMainForm = true;
                    frmTemp3.ParentControl = frmMain.g_oFrmMain;
                    frmTemp3.ParentControl.Enabled = false;
                    frmTemp3.Show();
                    break;
                case "HELP":
                    if (m_oHelp == null)
                    {
                        m_oHelp = new Help(m_xpsFile, m_oEnv);
                    }
                    m_oHelp.ShowHelp(new string[] { "DATABASE", "PLOT_DATA_MENU" });
                    break;
			}
		}
		
	}
}

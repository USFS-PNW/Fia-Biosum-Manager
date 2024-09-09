using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;


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
        private System.Windows.Forms.ContextMenu contextMenu1;
        private System.Windows.Forms.MenuItem mnuEditDeleteAll;
        private System.Windows.Forms.MenuItem mnuEditBrowse;
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
        private ToolBarButton tlbbtnEdit;
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
            this.tlbbtnEdit = new System.Windows.Forms.ToolBarButton();
            this.contextMenu1 = new System.Windows.Forms.ContextMenu();
            this.mnuEditDeleteAll = new System.Windows.Forms.MenuItem();
            this.mnuEditBrowse = new System.Windows.Forms.MenuItem();
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
            this.tlbbtnEdit,
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
            // Hide Condition zapper until we can convert it to sqlite
            this.tblbtnDeleteConds.Visible = false;
            // 
            // tblbtnDeletePackages
            // 
            this.tblbtnDeletePackages.ImageIndex = 1;
            this.tblbtnDeletePackages.Name = "tblbtnDeletePackages";
            this.tblbtnDeletePackages.Text = "Delete Packages";
            // Hide Package zapper until we can convert it to sqlite
            this.tblbtnDeletePackages.Visible = false;
            // 
            // tlbbtnEdit
            // 
            this.tlbbtnEdit.DropDownMenu = this.contextMenu1;
            this.tlbbtnEdit.Enabled = false;
            this.tlbbtnEdit.ImageIndex = 1;
            this.tlbbtnEdit.Name = "tlbbtnEdit";
            this.tlbbtnEdit.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton;
            this.tlbbtnEdit.Text = "Delete Plot Data";
            this.tlbbtnEdit.Visible = false;
            // 
            // contextMenu1
            // 
            this.contextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuEditDeleteAll,
            this.mnuEditBrowse});
            // 
            // mnuEditDeleteAll
            // 
            this.mnuEditDeleteAll.Index = 0;
            this.mnuEditDeleteAll.Text = "Delete All Plot Records";
            this.mnuEditDeleteAll.Click += new System.EventHandler(this.mnuEditDeleteAll_Click);
            // 
            // mnuEditBrowse
            // 
            this.mnuEditBrowse.Index = 1;
            this.mnuEditBrowse.Text = "Browse And Delete Selected Plot Records";
            this.mnuEditBrowse.Click += new System.EventHandler(this.mnuEditBrowse_Click);
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
					//FIA_Biosum_Manager.uc_plot_input uc_plot_input1 = new uc_plot_input();
					//frmTemp.Controls.Add(uc_plot_input1);
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

                case "BROWSE AND DELETE SELECTED PLOT RECORDS":
					//instantiate the datasource class
					FIA_Biosum_Manager.Datasource p_datasource = new Datasource(((frmDialog)this.ParentForm).m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim());
	              
                    int intPlot = p_datasource.getTableNameRow("PLOT");
					//see if datasource found the PLOT table type 
					if (intPlot >= 0)
					{
						//see if the MDB file is found and the plot table is found
						if (p_datasource.m_strDataSource[intPlot,FILESTATUS].Trim() == "F" &&
							p_datasource.m_strDataSource[intPlot,TABLESTATUS].Trim() == "F")
						{
							//see if there are records in the plot table
							if (p_datasource.m_strDataSource[intPlot,RECORDCOUNT].Trim() != "0")
							{
								string strConn="";
								string strFile = p_datasource.m_strDataSource[intPlot,MDBFILE].Trim();

								if (strFile.Substring(strFile.Trim().Length - 6,6).ToUpper()==".ACCDB")
									strConn = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + 
										strFile + ";User Id=admin;Password=;";
								else
									strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + 
										strFile + ";User Id=admin;Password=;";

								//string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + 
								//	              p_datasource.m_strDataSource[intPlot,MDBFILE].Trim() + 
                                //                ";User Id=admin;Password=;";

								FIA_Biosum_Manager.frmGridView frmGV = new frmGridView();

								frmGV.Text = "Database: Browse And Delete Plot Records";
								frmGV.LoadDataSetToDeleteOnly(strConn,"select * from " + p_datasource.m_strDataSource[intPlot,TABLE].Trim() + ";",p_datasource.m_strDataSource[intPlot,TABLE].Trim());
								frmGV.Show();
								frmGV.Focus();
							}
							else
							{
							  MessageBox.Show("!!No Records In The PLOT table!!","BROWSE AND DELETE PLOT RECORDS",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
							}
						}
						else
						{
						   MessageBox.Show("!!Could not find either the MDB File named \n" + 
							                p_datasource.m_strDataSource[intPlot,MDBFILE] + "\n" + 
							                " or the plot table named \n" + 
							                p_datasource.m_strDataSource[intPlot,TABLE].Trim() + "!!",
                                           "BROWSE AND DELETE PLOT RECORDS",
							                System.Windows.Forms.MessageBoxButtons.OK,
							               System.Windows.Forms.MessageBoxIcon.Exclamation);
						}
					}
					else
					{
						MessageBox.Show("!!Could not find the table type PLOT in the datasource table!!","BROWSE AND DELETE PLOT RECORDS",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
					}
					break;

				case "EDIT PLOT DATA":
					break;
                case "DELETE ALL PLOT RECORDS":
					this.DeleteAllPlotRecords();
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
		/// <summary>
		/// Delete all plot records and the plot's tree and condition records
		/// </summary>
		private void DeleteAllPlotRecords()
		{
			string strMsg = "All condition and tree records for the plots deleted will also be deleted. \n Are you sure you want to remove all plot data from the project?";
			DialogResult result = MessageBox.Show(strMsg,"Delete All Plot Data", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
			switch (result) 
			{
				case DialogResult.Yes:
					break;
				case DialogResult.No:
					return;
			    case DialogResult.Cancel:
					return;
			}
            frmMain.g_oFrmMain.ActivateStandByAnimation(frmMain.g_oFrmMain.WindowState,
                               this.Left, this.Height, this.Width, this.Top);
			//instantiate the ado data access class
			FIA_Biosum_Manager.ado_data_access  p_ado = new ado_data_access();

            //instantiate the datasource class
			FIA_Biosum_Manager.Datasource p_datasource = new Datasource(((frmDialog)this.ParentForm).m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim());
            
			string strSQL="";
			/********************************************************************
			 **have the data source class create an mdb file in the user's temp
			 **directory, create links to the plot,tree, and cond table and
			 **return the name of the mdb path and file
			 ********************************************************************/
			string strMDBFile = p_datasource.CreateMDBAndTableDataSourceLinks();

			//have the ado data access class return the connection string
			string strConn = p_ado.getMDBConnString(strMDBFile,"","");

			//get the location in the array for each of the table type information
			int intTree = p_datasource.getValidTableNameRow("TREE");
			int intPlot = p_datasource.getValidTableNameRow("PLOT");
			int intCond = p_datasource.getValidTableNameRow("CONDITION");
			int intTreeRegionalBiomass = p_datasource.getValidTableNameRow("TREE REGIONAL BIOMASS");
			int intSiteTree = p_datasource.getValidTableNameRow("SITE TREE");

			//check to see if we found all the table information
			if (intTree >= 0 && intPlot >=0 && intCond >=0 )
			{
				//check if we have tree regional biomass records
				if (p_datasource.m_strDataSource[intTreeRegionalBiomass,RECORDCOUNT].Trim() != "0")
				{
					
					//delete the tree regional records that are related to a plot
					strSQL =  "DELETE FROM " + 
						p_datasource.m_strDataSource[intTreeRegionalBiomass,TABLE];// + " tr " + 
//						"WHERE EXISTS (SELECT * FROM " + p_datasource.m_strDataSource[intCond,TABLE] + " c, " +
//						              p_datasource.m_strDataSource[intPlot,TABLE] + " p, " + 
//						              p_datasource.m_strDataSource[intTree,TABLE] + " t " + 
//									 "WHERE c.biosum_plot_id=p.biosum_plot_id AND " +
//						             "t.biosum_cond_id=c.biosum_cond_id AND " + 
//									 "tr.tre_cn = t.cn);";
					p_ado.SqlNonQuery(strConn,strSQL);
				}

				//check if we have site tree records
				if (p_datasource.m_strDataSource[intSiteTree,RECORDCOUNT].Trim() != "0")
				{
					strSQL =  "DELETE FROM " + 
						p_datasource.m_strDataSource[intSiteTree,TABLE]; 
					p_ado.SqlNonQuery(strConn,strSQL);
				}

                //check if we have tree records
				if (p_datasource.m_strDataSource[intTree,RECORDCOUNT].Trim() != "0")
				{
					
                   //delete the tree records that are related to a plot
					strSQL =  "DELETE FROM " + 
						p_datasource.m_strDataSource[intTree,TABLE]; // + " t " + 
					//	"WHERE EXISTS (SELECT * FROM " + p_datasource.m_strDataSource[intCond,TABLE] + " c " +
					//	" INNER JOIN " + p_datasource.m_strDataSource[intPlot,TABLE] + " p " + 
					//	"ON c.biosum_plot_id = p.biosum_plot_id " + 
					//	" WHERE t.biosum_cond_id = c.biosum_cond_id);";
					
					p_ado.SqlNonQuery(strConn,strSQL);
				}
				//check if we have cond records
				if (p_ado.m_intError == 0 && p_datasource.m_strDataSource[intCond,RECORDCOUNT].Trim() != "0")
				{
					
                   //delete the cond records that are related to a plot
					strSQL =  "DELETE FROM " + 
						p_datasource.m_strDataSource[intCond,TABLE]; // + " c " + 
//						"WHERE EXISTS (SELECT * FROM " + p_datasource.m_strDataSource[intPlot,TABLE] + " p " +
//						" WHERE c.biosum_plot_id = p.biosum_plot_id);";

					
					p_ado.SqlNonQuery(strConn,strSQL);
				}

				//check if we have plot records
				if (p_ado.m_intError == 0 && p_datasource.m_strDataSource[intPlot,RECORDCOUNT].Trim() != "0")
				{
					
                    //delete all the plot records                   
					strSQL =  "DELETE FROM " + 
						p_datasource.m_strDataSource[intPlot,TABLE]; 

					
					p_ado.SqlNonQuery(strConn,strSQL);
				}
                frmMain.g_oFrmMain.DeactivateStandByAnimation();
				MessageBox.Show("Done","DELETE ALL PLOT DATA",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.None);
			}
			else
			{
                frmMain.g_oFrmMain.DeactivateStandByAnimation();
				strSQL="Delete Failed - Could not locate these designated tables:\n";
				if (intPlot==-1)
					strSQL+= "Plot table\n";
				if (intCond==-1)
					strSQL+= "Cond table\n";
				if (intTree==-1)
					strSQL+= "Tree table\n";
				
					
				MessageBox.Show(strSQL,"DELETE ALL PLOT DATA",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			}
		}

		private void mnuEditDeleteAll_Click(object sender, System.EventArgs e)
		{
		    this.DeleteAllPlotRecords();
		}

		private void mnuEditBrowse_Click(object sender, System.EventArgs e)
		{
			//instantiate the datasource class
			//FIA_Biosum_Manager.Datasource p_datasource = new Datasource(((frmDialog)this.ParentForm).m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim());
			FIA_Biosum_Manager.Datasource p_datasource = new Datasource();
			p_datasource.LoadTableColumnNamesAndDataTypes=false;
			p_datasource.LoadTableRecordCount=false;
			p_datasource.m_strDataSourceMDBFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\project.mdb";
			p_datasource.m_strDataSourceTableName = "datasource";
			p_datasource.m_strScenarioId="";
			p_datasource.populate_datasource_array();
			
	              
			int intPlot = p_datasource.getTableNameRow("PLOT");
			//see if datasource found the PLOT table type 
			if (intPlot >= 0)
			{
				string strFile=p_datasource.m_strDataSource[intPlot,PATH].Trim() + "\\" +
					p_datasource.m_strDataSource[intPlot,MDBFILE].Trim();

				ado_data_access oAdo = new ado_data_access();
				string strConn = oAdo.getMDBConnString(strFile,"","");
				oAdo.OpenConnection(strConn);
				if (oAdo.m_intError==0)
				{
					//see if the MDB file is found and the plot table is found
					if (p_datasource.m_strDataSource[intPlot,FILESTATUS].Trim() == "F" &&
						p_datasource.m_strDataSource[intPlot,TABLESTATUS].Trim() == "F")
					{
						//see if there are records in the plot table
						if (Convert.ToInt32(oAdo.getRecordCount(oAdo.m_OleDbConnection,"SELECT COUNT(*) FROM " + p_datasource.m_strDataSource[intPlot,TABLE].Trim(),"temp")) > 0)
						{
							//string strConn="";
							

							oAdo.CloseConnection(oAdo.m_OleDbConnection);
						
							FIA_Biosum_Manager.frmGridView frmGV = new frmGridView();
							frmGV.strProjectDirectory = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();

							//frmGV.MdiParent = (frmMain)this.ParentForm.ParentForm;
							frmGV.Text = "Database: Browse And Delete Plot Records";
							frmGV.LoadDataSetToDeleteOnly(strConn,"select * from " + p_datasource.m_strDataSource[intPlot,TABLE].Trim() + ";",p_datasource.m_strDataSource[intPlot,TABLE].Trim());
                            frmMain.g_oFrmMain.Enabled = false;
                            frmGV.MinimizeMainForm = true;
							frmGV.Show();
							//frmGV.Show();
							//frmGV.Focus();
						}
						else
						{
							oAdo.CloseConnection(oAdo.m_OleDbConnection);
							MessageBox.Show("!!No Records In The PLOT table!!","BROWSE AND DELETE PLOT RECORDS",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
						}
					}
					else
					{
						oAdo.CloseConnection(oAdo.m_OleDbConnection);
						MessageBox.Show("!!Could not find either the MDB File named \n" + 
							p_datasource.m_strDataSource[intPlot,MDBFILE] + "\n" + 
							" or the plot table named \n" + 
							p_datasource.m_strDataSource[intPlot,TABLE].Trim() + "!!",
							"BROWSE AND DELETE PLOT RECORDS",
							System.Windows.Forms.MessageBoxButtons.OK,
							System.Windows.Forms.MessageBoxIcon.Exclamation);
					}
				}
			}
			else
			{
				MessageBox.Show("!!Could not find the table type PLOT in the datasource table!!","BROWSE AND DELETE PLOT RECORDS",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			}
		
		}

		
	}
}

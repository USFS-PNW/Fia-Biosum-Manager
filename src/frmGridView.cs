using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for frmGridView.
	/// </summary>
	public class frmGridView : System.Windows.Forms.Form
	{
		public System.Windows.Forms.ToolBar toolBar1;
		private FIA_Biosum_Manager.uc_gridview_collection uc_gridview_collection1;
		private FIA_Biosum_Manager.uc_gridview uc_gridview1;
		
		private System.Windows.Forms.ImageList imageList1;
		public int m_intError=0;
		public string m_strError="";
		public int m_intNumberOfGridViews=0;
		private System.Windows.Forms.ToolBarButton btnMaxSize;
		private System.Windows.Forms.ToolBarButton btnMultPane;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.ContextMenu m_ContextMenu;
		private System.Windows.Forms.ContextMenu m_ContextMenu2;
		
		private int m_intArrayCount=0;
		public System.Windows.Forms.VScrollBar m_vScrollBar;
		private int m_intTopPos;
		private int m_intCurrBtn=0;
		private int m_intRowCount=0;
		private string _strProjectDirectory;
		private bool _bHarvestCostColumns=false;
		private System.Windows.Forms.ToolBarButton Separator;
		private System.Windows.Forms.ToolBarButton btnFont;
		private System.Windows.Forms.ToolBarButton btnAlternateRowColor;
		private System.Windows.Forms.ToolBarButton btnBackground;
		private System.Windows.Forms.ToolBarButton btnRowColor;
		private System.Windows.Forms.ToolBarButton btnSelectedRowBackgroundColor;
		FIA_Biosum_Manager.frmOptimizerScenario _frmCoreScenario=null;
        FIA_Biosum_Manager.frmProcessorScenario _frmProcessorScenario = null;

        private Control _oParentControl = null;
        private bool _bMinimizeMainForm = false;
        private string _strCallingClient = "";
        private uc_processor_scenario_additional_harvest_cost_columns _uc_processor_scenario_additional_harvest_cost_columns = null;
        private bool _bUsingSQLite = false;


        public frmGridView()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			Initialize();
			

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmGridView));
            this.toolBar1 = new System.Windows.Forms.ToolBar();
            this.btnMaxSize = new System.Windows.Forms.ToolBarButton();
            this.btnMultPane = new System.Windows.Forms.ToolBarButton();
            this.Separator = new System.Windows.Forms.ToolBarButton();
            this.btnFont = new System.Windows.Forms.ToolBarButton();
            this.btnBackground = new System.Windows.Forms.ToolBarButton();
            this.btnRowColor = new System.Windows.Forms.ToolBarButton();
            this.btnAlternateRowColor = new System.Windows.Forms.ToolBarButton();
            this.btnSelectedRowBackgroundColor = new System.Windows.Forms.ToolBarButton();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.SuspendLayout();
            // 
            // toolBar1
            // 
            this.toolBar1.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
            this.btnMaxSize,
            this.btnMultPane,
            this.Separator,
            this.btnFont,
            this.btnBackground,
            this.btnRowColor,
            this.btnAlternateRowColor,
            this.btnSelectedRowBackgroundColor});
            this.toolBar1.DropDownArrows = true;
            this.toolBar1.ImageList = this.imageList1;
            this.toolBar1.Location = new System.Drawing.Point(0, 0);
            this.toolBar1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.toolBar1.Name = "toolBar1";
            this.toolBar1.ShowToolTips = true;
            this.toolBar1.Size = new System.Drawing.Size(453, 36);
            this.toolBar1.TabIndex = 0;
            this.toolBar1.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.toolBar1_ButtonClick);
            // 
            // btnMaxSize
            // 
            this.btnMaxSize.ImageIndex = 0;
            this.btnMaxSize.Name = "btnMaxSize";
            this.btnMaxSize.Pushed = true;
            this.btnMaxSize.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton;
            this.btnMaxSize.ToolTipText = "Single Pane";
            // 
            // btnMultPane
            // 
            this.btnMultPane.ImageIndex = 1;
            this.btnMultPane.Name = "btnMultPane";
            this.btnMultPane.Style = System.Windows.Forms.ToolBarButtonStyle.DropDownButton;
            this.btnMultPane.ToolTipText = "Multiple Panes";
            // 
            // Separator
            // 
            this.Separator.Name = "Separator";
            this.Separator.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
            // 
            // btnFont
            // 
            this.btnFont.ImageIndex = 2;
            this.btnFont.Name = "btnFont";
            this.btnFont.ToolTipText = "Grid Font";
            // 
            // btnBackground
            // 
            this.btnBackground.ImageIndex = 3;
            this.btnBackground.Name = "btnBackground";
            this.btnBackground.ToolTipText = "Grid Background Color";
            // 
            // btnRowColor
            // 
            this.btnRowColor.ImageIndex = 4;
            this.btnRowColor.Name = "btnRowColor";
            this.btnRowColor.ToolTipText = "Grid Row Background Color";
            // 
            // btnAlternateRowColor
            // 
            this.btnAlternateRowColor.ImageIndex = 5;
            this.btnAlternateRowColor.Name = "btnAlternateRowColor";
            this.btnAlternateRowColor.ToolTipText = "Grid Alternate Row Background Color";
            // 
            // btnSelectedRowBackgroundColor
            // 
            this.btnSelectedRowBackgroundColor.ImageIndex = 6;
            this.btnSelectedRowBackgroundColor.Name = "btnSelectedRowBackgroundColor";
            this.btnSelectedRowBackgroundColor.ToolTipText = "Selected Row Background Color";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "");
            this.imageList1.Images.SetKeyName(1, "");
            this.imageList1.Images.SetKeyName(2, "");
            this.imageList1.Images.SetKeyName(3, "");
            this.imageList1.Images.SetKeyName(4, "");
            this.imageList1.Images.SetKeyName(5, "");
            this.imageList1.Images.SetKeyName(6, "");
            // 
            // frmGridView
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(453, 303);
            this.Controls.Add(this.toolBar1);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "frmGridView";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "frmGridView";
            this.Closing += new System.ComponentModel.CancelEventHandler(this.frmGridView_Closing);
            this.Load += new System.EventHandler(this.frmGridView_Load);
            this.Resize += new System.EventHandler(this.frmGridView_Resize);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion
		public void SaveAll()
		{
			int intCount;
			intCount = this.uc_gridview_collection1.Count;
			uc_gridview p_gridview;
			for (int a=this.uc_gridview_collection1.Count-1; a > 0; a--)
			{
				//get a reference to the gridview object and kill it
				p_gridview = new uc_gridview();
				p_gridview.ReferenceGridViewForm=this;
				p_gridview = this.uc_gridview_collection1.Item(a).getGridViewObject;
				if (p_gridview.btnSave.Enabled==true)
				{
					p_gridview.savevalues();
				}
				
			}

		}
		private void frmGridView_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{

			int intCount;
			intCount = this.uc_gridview_collection1.Count;
			uc_gridview p_gridview;
			for (int a=this.uc_gridview_collection1.Count-1; a > 0; a--)
			{
				//get a reference to the gridview object and kill it
				p_gridview = new uc_gridview();
				p_gridview.ReferenceGridViewForm=this;
				p_gridview = this.uc_gridview_collection1.Item(a).getGridViewObject;
				p_gridview.CloseGridView();

			}
            if (CallingClient.Trim() == "ProcessorScenario:HarvesCostColumns_EditAll")
                ReferenceProcessorScenarioAdditionalHarvestCostColumns.UpdateNullCounts(ReferenceProcessorScenarioAdditionalHarvestCostColumns.TempDb);

			this.m_intNumberOfGridViews = 0;
            frmMain.g_oFrmMain.Enabled = true;
            if (ParentControl != null) ParentControl.Enabled = true;

           
			
		}
		public void LoadDataSet(System.Data.SQLite.SQLiteConnection p_conn,
			string strSQL, string strTableName)
		{
			this.AddDataSet(p_conn, strSQL, strTableName);
		}
		public void LoadDataSet(System.Data.SQLite.SQLiteConnection p_conn, string strSQL)
        {
			this.AddDataSet(p_conn, strSQL, "DataSet" + Convert.ToString(this.m_intNumberOfGridViews + 1).Trim());
        }
		public void LoadDataSet(string strConn, string strSQL)
		{
			this.AddDataSet(strConn,strSQL,"DataSet" + Convert.ToString(this.m_intNumberOfGridViews + 1).Trim());
		}
		public void LoadDataSet(string strConn, string strSQL,string strDataSetName)
		{
			this.AddDataSet(strConn,strSQL,strDataSetName);
		}
		public void LoadDataSetToEdit(string strConn, string strSQL,string strDataSetName,string[] strColumnsToEdit,int intColumnsToEditCount,string[] strRecordKeyColumns)
		{
			this.AddDataSetToEdit(strConn,strSQL,strDataSetName,strColumnsToEdit,intColumnsToEditCount,strRecordKeyColumns);
		}

		private void AddDataSet(string strConn,string strSQL,string strDataSetName)
		{
			this.m_intArrayCount++;
            if (!_bUsingSQLite)
            {
                this.uc_gridview1 = new uc_gridview(strConn, strSQL, strDataSetName);
            }
            else
            {
                this.uc_gridview1 = new uc_gridview();
                this.uc_gridview1.LoadGridViewSqlite(strConn, strSQL, strDataSetName);
            }

			this.uc_gridview1.ReferenceGridViewForm=this;
			this.uc_gridview_collection1.Add(this.uc_gridview1);
			if (this.uc_gridview1.m_intError==0)
			{
				this.m_intNumberOfGridViews++;
				this.Controls.Add(this.uc_gridview1);
				MenuItem p_menuitem = new MenuItem();
				MenuItem p_menuitem2 = new MenuItem();
				p_menuitem.Text = strDataSetName;
				p_menuitem2.Text = strDataSetName;
				m_ContextMenu.MenuItems.Add(p_menuitem);
				m_ContextMenu2.MenuItems.Add(p_menuitem2);
				p_menuitem.Click += new System.EventHandler(this.m_ContextMenu_Click);
				p_menuitem2.Click += new System.EventHandler(this.m_ContextMenu2_Click);
				for (int x=0; x<=this.btnMaxSize.DropDownMenu.MenuItems.Count-1;x++)
				{
					if (x < this.btnMaxSize.DropDownMenu.MenuItems.Count-1)
					{
						this.btnMaxSize.DropDownMenu.MenuItems[x].Checked=false;
					}
					else this.btnMaxSize.DropDownMenu.MenuItems[x].Checked=true;
					this.btnMultPane.DropDownMenu.MenuItems[x].Checked=true;
				}
				this.uc_gridview1.m_intID = this.m_intNumberOfGridViews;
				this.uc_gridview1.Width = this.Width - this.m_vScrollBar.Width;
				this.uc_gridview1.Visible=true;
				this.ResizeGridViewItem();
				this.resize_frmGridView();
			}

		}
		

		private void AddDataSet(System.Data.SQLite.SQLiteConnection p_conn,
						string strSQL,
						string strDataSetName)
		{
			this.m_intArrayCount++;
			this.uc_gridview1 = new uc_gridview(p_conn, strSQL, strDataSetName);
			this.uc_gridview1.ReferenceGridViewForm = this;
			this.uc_gridview_collection1.Add(this.uc_gridview1);
			if (this.uc_gridview1.m_intError == 0)
			{
				this.m_intNumberOfGridViews++;
				this.Controls.Add(this.uc_gridview1);
				MenuItem p_menuitem = new MenuItem();
				MenuItem p_menuitem2 = new MenuItem();
				p_menuitem.Text = strDataSetName;
				p_menuitem2.Text = strDataSetName;
				m_ContextMenu.MenuItems.Add(p_menuitem);
				m_ContextMenu2.MenuItems.Add(p_menuitem2);
				p_menuitem.Click += new System.EventHandler(this.m_ContextMenu_Click);
				p_menuitem2.Click += new System.EventHandler(this.m_ContextMenu2_Click);
				for (int x = 0; x <= this.btnMaxSize.DropDownMenu.MenuItems.Count - 1; x++)
				{
					if (x < this.btnMaxSize.DropDownMenu.MenuItems.Count - 1)
					{
						this.btnMaxSize.DropDownMenu.MenuItems[x].Checked = false;
					}
					else this.btnMaxSize.DropDownMenu.MenuItems[x].Checked = true;
					this.btnMultPane.DropDownMenu.MenuItems[x].Checked = true;
				}
				this.uc_gridview1.m_intID = this.m_intNumberOfGridViews;
				this.uc_gridview1.Width = this.Width - this.m_vScrollBar.Width;
				this.uc_gridview1.Visible = true;
				this.ResizeGridViewItem();
				this.resize_frmGridView();
			}
		}


		private void AddDataSetToEdit(string strConn,string strSQL,string strDataSetName,string[] strColumnsToEdit, int intColumnsToEditCount,string[] strRecordKeyColumns)
		{
			this.m_intArrayCount++;
			this.uc_gridview1 = new uc_gridview(strConn,strSQL,strDataSetName,strColumnsToEdit,intColumnsToEditCount,strRecordKeyColumns);
			this.uc_gridview1.ReferenceGridViewForm=this;
			this.uc_gridview_collection1.Add(this.uc_gridview1);
			if (this.uc_gridview1.m_intError==0)
			{
				this.m_intNumberOfGridViews++;
				this.Controls.Add(this.uc_gridview1);
				MenuItem p_menuitem = new MenuItem();
				MenuItem p_menuitem2 = new MenuItem();
				p_menuitem.Text = strDataSetName;
				p_menuitem2.Text = strDataSetName;
				m_ContextMenu.MenuItems.Add(p_menuitem);
				m_ContextMenu2.MenuItems.Add(p_menuitem2);
				p_menuitem.Click += new System.EventHandler(this.m_ContextMenu_Click);
				p_menuitem2.Click += new System.EventHandler(this.m_ContextMenu2_Click);
				for (int x=0; x<=this.btnMaxSize.DropDownMenu.MenuItems.Count-1;x++)
				{
					if (x < this.btnMaxSize.DropDownMenu.MenuItems.Count-1)
					{
						this.btnMaxSize.DropDownMenu.MenuItems[x].Checked=false;
					}
					else this.btnMaxSize.DropDownMenu.MenuItems[x].Checked=true;
					this.btnMultPane.DropDownMenu.MenuItems[x].Checked=true;
				}
				this.uc_gridview1.m_intID = this.m_intNumberOfGridViews;
				this.uc_gridview1.Width = this.Width - this.m_vScrollBar.Width;
				this.uc_gridview1.Visible=true;
				this.ResizeGridViewItem();
				this.resize_frmGridView();
			}

		}

		private void ResizeGridViewItem()
		{
			int y=0;
			int x=0;
			
			this.m_intRowCount=0;

			int intStartTop= this.toolBar1.Top + this.toolBar1.Height;
			if (this.btnMaxSize.Pushed==true)
			{
				this.m_intRowCount=1;
				this.m_vScrollBar.Value=0;

                //loop through each menu item
				for (int a=0; a<=this.m_ContextMenu.MenuItems.Count - 1; a++)
				{
					if (this.m_ContextMenu.MenuItems[a].Checked == true)
					{
						//loop through the connection and find the corresponding menu item
						//to position it and make it visible
						for (x=0; x<=this.uc_gridview_collection1.Count-1; x++)
						{
							if (this.uc_gridview_collection1.Item(x).DataSetName.Trim().ToUpper()
								== this.m_ContextMenu.MenuItems[a].Text.Trim().ToUpper())
							{
								this.uc_gridview_collection1.Item(x).Top = intStartTop - this.m_vScrollBar.Value;
								this.uc_gridview_collection1.Item(x).Height = this.ClientSize.Height - intStartTop;
								this.uc_gridview_collection1.Item(x).Width = this.Width - 20 - this.m_vScrollBar.Width;
								this.uc_gridview_collection1.Item(x).Left = 10;
								this.uc_gridview_collection1.Item(x).Visible=true;
								break;
							}
						}
					}
					else
					{
						//loop through the collection to find the corresponding
						//menu item and make it not visible
						for (x=0; x<=this.uc_gridview_collection1.Count-1; x++)
						{
							if (this.uc_gridview_collection1.Item(x).DataSetName.Trim().ToUpper()
								== this.m_ContextMenu.MenuItems[a].Text.Trim().ToUpper())
							{
								this.uc_gridview_collection1.Item(x).Visible = false;
								break;
							}
						}
					}
					
				}
				    
			}    
			else 
			{
				int p_Count = 0;
				int p_Count2 = 0;
				y=0;
				bool bDone=false;
				//get the number of checked menu items
				for (x=this.m_ContextMenu2.MenuItems.Count-1; x >= 0; x--)
				{
					if (this.m_ContextMenu2.MenuItems[x].Checked == true)
					{
						p_Count++;
						if (p_Count == 2)
							break;
					}
				}

				//loop through each menu item
				p_Count2 = 1;
				int p_Top = 0;
				int p_Ht = 0;
				int p_Wd = 0;
				int p_Left = 10;
				for (int a=this.m_ContextMenu2.MenuItems.Count - 1; a>=0; a--)
				{   
					//if menu is checked then make it visible and resize grid class
					if (this.m_ContextMenu2.MenuItems[a].Checked == true)
					{
						//loop through the gridview collection find the corresponding menu item
						//to position it and make it visible
						
						for (x=0; x<=this.uc_gridview_collection1.Count-1; x++)
						{
							if (this.uc_gridview_collection1.Item(x).DataSetName.Trim().ToUpper()
								== this.m_ContextMenu2.MenuItems[a].Text.Trim().ToUpper())
							{
								if (p_Count < 2)
								{
									//only one grid menu checked so enlarge grid class to size of form
									this.uc_gridview_collection1.Item(x).Top = intStartTop  - this.m_vScrollBar.Value;
									this.uc_gridview_collection1.Item(x).Height = this.ClientSize.Height - intStartTop;
									this.uc_gridview_collection1.Item(x).Width = this.Width - 20 - this.m_vScrollBar.Width;
									this.uc_gridview_collection1.Item(x).Left = 10;
									this.uc_gridview_collection1.Item(x).Visible=true;
									bDone=true;  //only one grid class visible so were done
									break;
								}
								else
								{
									//more than one menu item checked
									Math.DivRem(p_Count2,2,out y);  //get the remainder
									if (y != 0)
									{
										this.m_intRowCount++;
										//this is the left side grid classes
										if (p_Count2 == 1)
										{
										   
										   //this is the top left so dimension it and all
										   //subsequent grids will match the height and width
										   this.uc_gridview_collection1.Item(x).Top = this.m_intTopPos - this.m_vScrollBar.Value;
										   this.uc_gridview_collection1.Item(x).Height = (int)(this.ClientSize.Height * .50) - this.m_intTopPos;
										   this.uc_gridview_collection1.Item(x).Width = (int)(this.Width * .50) - 10 - this.m_vScrollBar.Width;
										   this.uc_gridview_collection1.Item(x).Left = (int)(this.Width * .50) - uc_gridview_collection1.Item(x).Width;
										   p_Top = this.uc_gridview_collection1.Item(x).Top ;
										   p_Wd = this.uc_gridview_collection1.Item(x).Width ;
										   p_Left = this.uc_gridview_collection1.Item(x).Left;
										   p_Ht   = this.uc_gridview_collection1.Item(x).Height;
										   this.uc_gridview_collection1.Item(x).Visible=true;
										}
										else 
										{
											this.uc_gridview_collection1.Item(x).Top = p_Top;
											this.uc_gridview_collection1.Item(x).Left = p_Left;
											this.uc_gridview_collection1.Item(x).Width = p_Wd;
											this.uc_gridview_collection1.Item(x).Height = p_Ht;
											
											p_Top = this.uc_gridview_collection1.Item(x).Top; // + this.uc_gridview_collection1.Item(x).Height;
											this.uc_gridview_collection1.Item(x).Visible=true;
										}
										

									}
									else
									{
									    //this is the right side grid classes
										this.uc_gridview_collection1.Item(x).Top = p_Top;
										this.uc_gridview_collection1.Item(x).Left = p_Left + p_Wd;
										this.uc_gridview_collection1.Item(x).Width = p_Wd;
										this.uc_gridview_collection1.Item(x).Height = p_Ht;
										this.uc_gridview_collection1.Item(x).Visible=true;
										p_Top += p_Ht;
										
									}
									p_Count2++;
								}
								if (bDone == true) break;
							}
						}
					}
					else
					{
						//loop through the collection to find the corresponding
						//menu item and make it not visible
						for (x=0; x<=this.uc_gridview_collection1.Count-1; x++)
						{
							if (this.uc_gridview_collection1.Item(x).DataSetName.Trim().ToUpper()
								== this.m_ContextMenu2.MenuItems[a].Text.Trim().ToUpper())
							{
								this.uc_gridview_collection1.Item(x).Visible = false;
								break;
							}
						}
					}
					
				}
				this.m_vScrollBar.Maximum = this.ClientSize.Height * this.m_intRowCount;


			}


			
		}

		private void frmGridView_Resize(object sender, System.EventArgs e)
		{
			this.resize_frmGridView();

		}
		public void resize_frmGridView()
		{
            if (this.MinimizeMainForm)
            {
                if (this.WindowState != System.Windows.Forms.FormWindowState.Minimized)
                {
                    frmMain.g_oFrmMain.WindowState = System.Windows.Forms.FormWindowState.Maximized;
                }
                else
                {
                    frmMain.g_oFrmMain.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                }

            }
			this.m_intRowCount=0;
			int intStartTop= this.toolBar1.Top + this.toolBar1.Height;
			bool bDone=false;
			int x=0;
			if (this.btnMaxSize.Pushed==true)
			{
				this.m_intRowCount=1;
				for (int b = 0; b <= this.btnMaxSize.DropDownMenu.MenuItems.Count-1;b++)
				{
					if (this.btnMaxSize.DropDownMenu.MenuItems[b].Checked == true)
					{
						for (int a=0; a<=this.uc_gridview_collection1.Count-1; a++)
						{
							if (this.btnMaxSize.DropDownMenu.MenuItems[b].Text.Trim().ToUpper() ==
								this.uc_gridview_collection1.Item(a).DataSetName.Trim().ToUpper())
							{
								this.uc_gridview_collection1.Item(a).Top = intStartTop;
								this.uc_gridview_collection1.Item(a).Height = this.ClientSize.Height - intStartTop;
								this.uc_gridview_collection1.Item(a).Width = this.Width - 20 - this.m_vScrollBar.Width;
								this.uc_gridview_collection1.Item(a).Left = 10;
								bDone=true;
								break;
							}
						}
					}
					if (bDone == true) break;
				}				    

			}    
			else 
			{
				int p_Count = 0;
				int p_Count2 = 0;
				

				//get the number of checked menu items
				for (x=this.m_ContextMenu2.MenuItems.Count-1; x >= 0; x--)
				{
					if (this.m_ContextMenu2.MenuItems[x].Checked == true)
					{
						p_Count++;
						if (p_Count == 2)
							break;
					}
				}

				//loop through each menu item
				p_Count2 = 1;
				int y=0;
				int p_Top = 0;
				int p_Ht = 0;
				int p_Wd = 0;
				int p_Left = 10;
				for (int a=this.m_ContextMenu2.MenuItems.Count - 1; a>=0; a--)
				{   
					//if menu is checked then make it visible and resize grid class
					if (this.m_ContextMenu2.MenuItems[a].Checked == true)
					{
						//loop through the gridview collection find the corresponding menu item
						//to position it and make it visible
						
						for (x=0; x<=this.uc_gridview_collection1.Count-1; x++)
						{
							if (this.uc_gridview_collection1.Item(x).DataSetName.Trim().ToUpper()
								== this.m_ContextMenu2.MenuItems[a].Text.Trim().ToUpper())
							{
								if (p_Count < 2)
								{
									//only one grid menu checked so enlarge grid class to size of form
									this.uc_gridview_collection1.Item(x).Top = intStartTop  - this.m_vScrollBar.Value;
									this.uc_gridview_collection1.Item(x).Height = this.ClientSize.Height - intStartTop;
									this.uc_gridview_collection1.Item(x).Width = this.Width - 20 - this.m_vScrollBar.Width;
									this.uc_gridview_collection1.Item(x).Left = 10;
									bDone=true;  //only one grid class visible so were done
									break;
								}
								else
								{
									//more than one menu item checked
									Math.DivRem(p_Count2,2,out y);  //get the remainder
									if (y != 0)
									{
										this.m_intRowCount++;
										//this is the left side grid classes
										if (p_Count2 == 1)
										{
											//this is the top left so dimension it and all
											//subsequent grids will match the height and width
											this.uc_gridview_collection1.Item(x).Top = this.m_intTopPos - this.m_vScrollBar.Value;
											this.uc_gridview_collection1.Item(x).Height = (int)(this.ClientSize.Height * .50) - this.m_intTopPos;
											this.uc_gridview_collection1.Item(x).Width = (int)(this.Width * .50) - 10 - this.m_vScrollBar.Width;
											this.uc_gridview_collection1.Item(x).Left = (int)(this.Width * .50) - uc_gridview_collection1.Item(x).Width;
											p_Top = this.uc_gridview_collection1.Item(x).Top ;
											p_Wd = this.uc_gridview_collection1.Item(x).Width ;
											p_Left = this.uc_gridview_collection1.Item(x).Left;
											p_Ht   = this.uc_gridview_collection1.Item(x).Height;
										}
										else 
										{
											this.uc_gridview_collection1.Item(x).Top = p_Top;
											this.uc_gridview_collection1.Item(x).Left = p_Left;
											this.uc_gridview_collection1.Item(x).Width = p_Wd;
											this.uc_gridview_collection1.Item(x).Height = p_Ht;
											
											p_Top = this.uc_gridview_collection1.Item(x).Top; // + this.uc_gridview_collection1.Item(x).Height;
										}
										

									}
									else
									{
										//this is the right side grid classes
										this.uc_gridview_collection1.Item(x).Top = p_Top;
										this.uc_gridview_collection1.Item(x).Left = p_Left + p_Wd;
										this.uc_gridview_collection1.Item(x).Width = p_Wd;
										this.uc_gridview_collection1.Item(x).Height = p_Ht;
										p_Top += p_Ht;
									}
									p_Count2++;
								}
								if (bDone == true) break;
							}
						}
					}
					else
					{
						//loop through the collection to find the corresponding
						//menu item and make it not visible
						for (x=0; x<=this.uc_gridview_collection1.Count-1; x++)
						{
							if (this.uc_gridview_collection1.Item(x).DataSetName.Trim().ToUpper()
								== this.m_ContextMenu2.MenuItems[a].Text.Trim().ToUpper())
							{
								this.uc_gridview_collection1.Item(x).Visible = false;
								break;
							}
						}
					}
				}
			}
			this.m_vScrollBar.Maximum = this.ClientSize.Height * this.m_intRowCount;

		}
		//remove a gridview menu item
		public void RemoveGridViewMenuItem(string strDataSetName)
		{

			for (int x=0; x<=this.m_ContextMenu.MenuItems.Count-1;x++)
			{
				if (this.m_ContextMenu.MenuItems[x].Text.Trim().ToUpper() == 
					strDataSetName.Trim().ToUpper())
				{
					this.m_ContextMenu.MenuItems.Remove(this.m_ContextMenu.MenuItems[x]);
					this.m_ContextMenu2.MenuItems.Remove(this.m_ContextMenu2.MenuItems[x]);
					break;
				}
			}
			
		}

		//remove all the gridview collection items
		public void RemoveGridViewCollectionItems()
		{
			int intCount;
			intCount = this.uc_gridview_collection1.Count;
			uc_gridview p_gridview;
			for (int a=this.uc_gridview_collection1.Count-1; a >= 0; a--)
			{
				p_gridview = new uc_gridview();
				p_gridview.ReferenceGridViewForm=this;
				p_gridview = this.uc_gridview_collection1.Item(a).getGridViewObject;
				p_gridview.CloseGridView();
     			this.uc_gridview_collection1.Remove(a);
			}

			this.m_intNumberOfGridViews = 0;
		}
        //remove a single gridview collection item
		public void RemoveGridViewCollectionItem(string strDataSetName)
		{
			for (int a=0; a<=this.uc_gridview_collection1.Count-1; a++)
			{
				if (strDataSetName.Trim().ToUpper() ==
					this.uc_gridview_collection1.Item(a).DataSetName.Trim().ToUpper())
				{
					this.uc_gridview_collection1.Remove(a);
					break;
				}
			}

			this.m_intNumberOfGridViews = this.m_intNumberOfGridViews-1;
			if (this.toolBar1.Buttons[1].Pushed==true) this.ResizeGridViewItem();
		}

		private void toolBar1_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			
			switch (this.toolBar1.Buttons.IndexOf(e.Button))
			{
				case 0:   //max
					if (this.m_intCurrBtn == this.toolBar1.Buttons.IndexOf(e.Button)) return;
					this.m_intCurrBtn = this.toolBar1.Buttons.IndexOf(e.Button);
					this.toolBar1.Buttons[0].Pushed = true;
					this.toolBar1.Buttons[1].Pushed = false;
                    this.ResizeGridViewItem();
					
					break;
				case 1:  //multiple pane
					if (this.m_intCurrBtn == this.toolBar1.Buttons.IndexOf(e.Button)) return;
					this.m_intCurrBtn = this.toolBar1.Buttons.IndexOf(e.Button);
					this.toolBar1.Buttons[0].Pushed = false;
					this.toolBar1.Buttons[1].Pushed = true;
					this.ResizeGridViewItem();
					
					break;
				case 3:
					System.Windows.Forms.FontDialog frmTemp = new FontDialog();
					frmTemp.ShowColor=true;
					frmTemp.Font=frmMain.g_oGridViewFont;
					frmTemp.Color = frmMain.g_oGridViewRowForegroundColor;
					if(frmTemp.ShowDialog() != DialogResult.Cancel )
					{
						frmMain.g_oGridViewFont = frmTemp.Font;
						frmMain.g_oGridViewRowForegroundColor = frmTemp.Color;

						for (int a=0; a<=this.uc_gridview_collection1.Count-1; a++)
						{
							if (uc_gridview_collection1.Item(a).m_dg.TableStyles.Count > 0)
							{
								uc_gridview_collection1.Item(a).m_dg.TableStyles[0].DataGrid.Font = frmTemp.Font;
								uc_gridview_collection1.Item(a).m_dg.TableStyles[0].ForeColor = frmTemp.Color;
							}
							
						}
					}
					frmTemp.Dispose();
					frmTemp=null;

					break;
				case 4:
					System.Windows.Forms.ColorDialog frmTemp2 = new ColorDialog();
					frmTemp2.Color = frmMain.g_oGridViewBackgroundColor;
					if(frmTemp2.ShowDialog() != DialogResult.Cancel )
					{
						for (int a=0; a<=this.uc_gridview_collection1.Count-1; a++)
						{
							frmMain.g_oGridViewBackgroundColor=frmTemp2.Color;
							uc_gridview_collection1.Item(a).m_dg.BackgroundColor=frmTemp2.Color;
						}
					}
					frmTemp2.Dispose();
					frmTemp2=null;

					break;
				case 5:
					System.Windows.Forms.ColorDialog frmTemp3 = new ColorDialog();
					frmTemp3.Color = frmMain.g_oGridViewRowBackgroundColor;
					if(frmTemp3.ShowDialog() != DialogResult.Cancel )
					{
						for (int a=0; a<=this.uc_gridview_collection1.Count-1; a++)
						{
							frmMain.g_oGridViewRowBackgroundColor=frmTemp3.Color;
							if (uc_gridview_collection1.Item(a).m_dg.TableStyles.Count > 0)
							{
								uc_gridview_collection1.Item(a).m_dg.TableStyles[0].BackColor=frmTemp3.Color;                  
							}
							
						}
					}
					frmTemp3.Dispose();
					frmTemp3=null;

					break;
				case 6:
					System.Windows.Forms.ColorDialog frmTemp4 = new ColorDialog();
					frmTemp4.Color = frmMain.g_oGridViewAlternateRowBackgroundColor;
					if(frmTemp4.ShowDialog() != DialogResult.Cancel )
					{
						for (int a=0; a<=this.uc_gridview_collection1.Count-1; a++)
						{
							frmMain.g_oGridViewAlternateRowBackgroundColor=frmTemp4.Color;
							if (uc_gridview_collection1.Item(a).m_dg.TableStyles.Count > 0)
							{
								uc_gridview_collection1.Item(a).m_dg.TableStyles[0].AlternatingBackColor=frmTemp4.Color;                  
							}
							
						}
					}
					frmTemp4.Dispose();
					frmTemp4=null;
					break;
				case 7:
					System.Windows.Forms.ColorDialog frmTemp5 = new ColorDialog();
					frmTemp5.Color = frmMain.g_oGridViewSelectedRowBackgroundColor;
					if(frmTemp5.ShowDialog() != DialogResult.Cancel )
					{
						for (int a=0; a<=this.uc_gridview_collection1.Count-1; a++)
						{
							frmMain.g_oGridViewSelectedRowBackgroundColor=frmTemp5.Color;
							if (uc_gridview_collection1.Item(a).m_dg.TableStyles.Count > 0)
							{
								uc_gridview_collection1.Item(a).m_dg.TableStyles[0].SelectionBackColor=frmTemp5.Color;                  
							}
							
						}
					}
					frmTemp5.Dispose();
					frmTemp5=null;
					break;

			}
			
		}

		/// <summary>
		/// option to tile multiple gridviews
		/// </summary>
		public void TileGridViews()
		{
			if (this.m_intNumberOfGridViews > 1)
			{
				this.toolBar1.Buttons[0].Pushed = false;
				this.toolBar1.Buttons[1].Pushed = true;
				this.ResizeGridViewItem();
			}

		}
		private void m_ContextMenu_Click(object sender,System.EventArgs e )
		{

			int x=0;
			//do a direct cast of the sender object to the type of object it is
			System.Windows.Forms.MenuItem p_menu = sender as System.Windows.Forms.MenuItem;
			if (p_menu.Checked == true)
			{
				
				p_menu.Checked=false;
				
			}
			else 
			{
				for (x=0; x <= this.m_ContextMenu.MenuItems.Count - 1;x++)
					if (this.m_ContextMenu.MenuItems[x].Text.Trim().ToUpper()
						== p_menu.Text.Trim().ToUpper())
					{
						p_menu.Checked=true; 
						if (this.toolBar1.Buttons[0].Pushed==false) 
						{
							this.toolBar1.Buttons[0].Pushed=true;
							this.toolBar1.Buttons[1].Pushed=false;
							this.m_intCurrBtn=0;

						}
					}
					else 
					{
						this.m_ContextMenu.MenuItems[x].Checked=false;
					}
                    
			}
			this.ResizeGridViewItem();

		    
		}
		private void m_ContextMenu2_Click(object sender, System.EventArgs e)
		{
			if (this.toolBar1.Buttons[1].Pushed==false) return;
			System.Windows.Forms.MenuItem p_menu = sender as System.Windows.Forms.MenuItem;
			if (p_menu.Checked == true)
			{
				
				p_menu.Checked=false;
				
			}
			else 
			{
				p_menu.Checked=true; 
                    
			}
			this.ResizeGridViewItem();
			

		}

		private void frmGridView_MouseWheel(Object sender, MouseEventArgs e)
		{
			if (this.btnMaxSize.Pushed==false)
			{
				if (e.Delta == 120)
				{
					if (this.m_vScrollBar.Value > this.m_vScrollBar.Minimum) 
					{

						if (this.m_vScrollBar.Value - this.m_vScrollBar.LargeChange < this.m_vScrollBar.Minimum) 
						{
							this.m_vScrollBar.Value = this.m_vScrollBar.Minimum;
						}
						else 
						{
							this.m_vScrollBar.Value = this.m_vScrollBar.Value - 
								this.m_vScrollBar.LargeChange;
						}
						this.RepositionGridViews();
					}
				}
				else 
				{
					if (this.m_vScrollBar.Value <= this.m_vScrollBar.Maximum) 
					{
						if (this.m_vScrollBar.Value + this.m_vScrollBar.LargeChange > 
							this.m_vScrollBar.Maximum) 
						{
							this.m_vScrollBar.Value = this.m_vScrollBar.Maximum;
						}
						else 
						{
							this.m_vScrollBar.Value = this.m_vScrollBar.Value +
								this.m_vScrollBar.LargeChange;
						}
						this.RepositionGridViews();
					}

				}
			}
			
		}
		private void vScrollBar_Scroll(Object sender, ScrollEventArgs e)
		{
			
		}

		private void vScrollBar_ValueChanged(Object sender, EventArgs e)
		{
			this.RepositionGridViews();
		}

		//called only when a scroll event occurs
		public void RepositionGridViews()
		{
			int x=0;
			int y=0;
			if (this.toolBar1.Buttons[0].Pushed == true)
			{
				//only 1 grid available at a time so find the visible grid
				//and size the grid to fill the entire form
				for (x=0; x <= this.uc_gridview_collection1.Count-1;x++)
				{
					if (this.uc_gridview_collection1.Item(x).Visible==true)
					{
						this.uc_gridview_collection1.Item(x).Top = this.m_intTopPos  - this.m_vScrollBar.Value;
						break;
					}
				}
			}
			else 
			{
			
				int p_Count = 0;
				int p_Count2 = 0;
				bool bDone = false;

				//get the number of checked menu items
				for (x=this.m_ContextMenu2.MenuItems.Count-1; x >= 0; x--)
				{
					if (this.m_ContextMenu2.MenuItems[x].Checked == true)
					{
						p_Count++;
						if (p_Count == 2)
							break;
					}
				}

				//loop through each menu item
				p_Count2 = 1;
				y=0;
				int p_Top = 0;
				int p_Ht = 0;
				for (int a=this.m_ContextMenu2.MenuItems.Count - 1; a>=0; a--)
				{   
					//if menu is checked then make it visible and resize grid class
					if (this.m_ContextMenu2.MenuItems[a].Checked == true)
					{
						//loop through the gridview collection find the corresponding menu item
						//to position it and make it visible
						
						for (x=0; x<=this.uc_gridview_collection1.Count-1; x++)
						{
							if (this.uc_gridview_collection1.Item(x).DataSetName.Trim().ToUpper()
								== this.m_ContextMenu2.MenuItems[a].Text.Trim().ToUpper())
							{
								if (p_Count < 2)
								{
									//only one grid menu checked so enlarge grid class to size of form
									this.uc_gridview_collection1.Item(x).Top = this.m_intTopPos - this.m_vScrollBar.Value;
									bDone=true;  //only one grid class visible so were done
									break;
								}
								else
								{
									//more than one menu item checked
									Math.DivRem(p_Count2,2,out y);  //get the remainder
									if (y != 0)
									{
										//this is the left side grid classes
										if (p_Count2 == 1)
										{
											//this is the top left so dimension it and all
											//subsequent grids will match the height and width
											this.uc_gridview_collection1.Item(x).Top = this.m_intTopPos - this.m_vScrollBar.Value;
											p_Top = this.uc_gridview_collection1.Item(x).Top ;
											p_Ht   = this.uc_gridview_collection1.Item(x).Height;
										}
										else 
										{
											this.uc_gridview_collection1.Item(x).Top = p_Top;
											
											p_Top = this.uc_gridview_collection1.Item(x).Top; // + this.uc_gridview_collection1.Item(x).Height;
											
										}
										

									}
									else
									{
										//this is the right side grid classes
										this.uc_gridview_collection1.Item(x).Top = p_Top;
										p_Top += p_Ht;
									}
									p_Count2++;
								}
								if (bDone == true) break;
							}
						}
					}
				}
			
			}
				
		}
		public void GridViewMaxSize(string strDataSetName)
		{
			int x=0;

			//find the menu item and check it and uncheck all the others
			for (x = 0; x <= this.m_ContextMenu.MenuItems.Count -1; x++)
			{
				if (this.m_ContextMenu.MenuItems[x].Text.Trim().ToUpper() == 
					strDataSetName.Trim().ToUpper())
				{
                     this.m_ContextMenu.MenuItems[x].Checked=true;
				}
				else
				{
					this.m_ContextMenu.MenuItems[x].Checked=false;
				}
			}
		    this.toolBar1.Buttons[1].Pushed=false;
			this.toolBar1.Buttons[0].Pushed=true;
			this.m_intCurrBtn = 0;
			this.ResizeGridViewItem();
			


		}
		private void Initialize()
		{
			this.uc_gridview1 = new uc_gridview();
			this.uc_gridview1.ReferenceGridViewForm=this;
			this.uc_gridview_collection1 = new uc_gridview_collection();
			this.uc_gridview_collection1.Add(this.uc_gridview1);
			this.m_ContextMenu = new ContextMenu();
			this.m_ContextMenu2 = new ContextMenu();
			this.btnMaxSize.DropDownMenu = this.m_ContextMenu;
			this.btnMultPane.DropDownMenu = this.m_ContextMenu2;
			
			this.Controls.Add(this.uc_gridview1);
			this.uc_gridview1.Visible=false;


			this.m_vScrollBar = new VScrollBar();
			this.Controls.Add(this.m_vScrollBar);
			this.m_vScrollBar.Dock = System.Windows.Forms.DockStyle.Right;
			this.m_vScrollBar.Visible=false;
			this.m_vScrollBar.Minimum=0;
			this.m_vScrollBar.Maximum=this.Height;
			this.m_vScrollBar.Value=0;
			this.m_vScrollBar.LargeChange = 20;
			this.m_vScrollBar.SmallChange = 10;
			this.m_vScrollBar.Visible=true;
			this.m_vScrollBar.Scroll += new ScrollEventHandler(
				vScrollBar_Scroll);
			this.m_vScrollBar.ValueChanged += new EventHandler(
				vScrollBar_ValueChanged);
			this.MouseWheel += new MouseEventHandler(frmGridView_MouseWheel);	
           

			this.m_intTopPos = this.toolBar1.Top + this.toolBar1.Height;

		}

		private void frmGridView_Load(object sender, System.EventArgs e)
		{
		
		}
		public string strProjectDirectory
		{
			get {return _strProjectDirectory;}
			set {_strProjectDirectory = value;}
		}
		public bool HarvestCostColumns
		{
			get {return _bHarvestCostColumns;}
			set {_bHarvestCostColumns = value;}
		}
		public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
		{
			get {return _frmCoreScenario;}
			set {_frmCoreScenario=value;}
		}
        public FIA_Biosum_Manager.frmProcessorScenario ReferenceProcessorScenarioForm
        {
            get { return _frmProcessorScenario; }
            set { _frmProcessorScenario = value; }
        }
        public Control ParentControl
        {
            get { return _oParentControl; }
            set { _oParentControl = value; }
        }
        public bool MinimizeMainForm
        {
            set { _bMinimizeMainForm = value; }
            get { return _bMinimizeMainForm; }
        }
        public string CallingClient
        {
            get { return _strCallingClient; }
            set { _strCallingClient = value; }
        }
        public uc_processor_scenario_additional_harvest_cost_columns ReferenceProcessorScenarioAdditionalHarvestCostColumns
        {
            get { return _uc_processor_scenario_additional_harvest_cost_columns; }
            set { _uc_processor_scenario_additional_harvest_cost_columns = value; }
        }
        public bool UsingSQLite
        {
            get { return _bUsingSQLite; }
            set { _bUsingSQLite = value; }
        }


    }
}

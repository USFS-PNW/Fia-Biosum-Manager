using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for frmEstEngRowCol.
	/// </summary>
	public class frmCollectionClone : System.Windows.Forms.Form
	{
		public System.Windows.Forms.Design.IWindowsFormsEditorService _wfes;
		private System.Windows.Forms.PropertyGrid propertyGrid1;
		private System.Windows.Forms.PropertyGrid propertyGrid2;
		private PropertyGrid_Collection PropertyGrid_Collection1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnHelp;
		private System.Windows.Forms.Button btnAdd;
		private System.Windows.Forms.Button btnRemove;
		private System.Windows.Forms.ListView listView1;
		private int m_intCurrSelect;
		private int m_intCurrGrid;
		private System.Windows.Forms.ImageList imageList1;
		private System.ComponentModel.IContainer components;
		public string m_strType;   //ERG=Excel Row Group Item, ER=Excel Row Item
		private int m_intPropertyGridCount=0;
		//SQL Variable Substitution
		private FIA_Biosum_Manager.SQLMacroSubstitutionVariable_Collection SQLMacroSubstitutionVariable_NewCollection1;
		private FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem m_SQLVarSubItem1;




		private System.Windows.Forms.Button btnUp;
		private System.Windows.Forms.Button btnDown;
		
        private bool _bEstEngRatio;
		private bool _bEstEngVarDirColumn;
		private int  _intIndexToDisplay=-1;
		private string _strReportColumns="";
		
		
		
		public frmCollectionClone()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			this.TopLevel = true;
			int intStringLength = (int)(this.CreateGraphics().MeasureString("0",this.listView1.Font).Width * 1.5);
			this.listView1.Columns.Add("", intStringLength, HorizontalAlignment.Right);
			this.listView1.Columns.Add("",180, HorizontalAlignment.Left);
			this.m_intCurrSelect=0;
			m_intCurrGrid=0;
			
			this.listView1.View = System.Windows.Forms.View.Details;
			
            this.propertyGrid1.Show();
			this.PropertyGrid_Collection1 = new PropertyGrid_Collection();


			//SQL Macro Variable Substitution
			this.SQLMacroSubstitutionVariable_NewCollection1 = new FIA_Biosum_Manager.SQLMacroSubstitutionVariable_Collection();

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCollectionClone));
			this.propertyGrid1 = new System.Windows.Forms.PropertyGrid();
			this.label1 = new System.Windows.Forms.Label();
			this.btnOK = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnHelp = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.btnAdd = new System.Windows.Forms.Button();
			this.btnRemove = new System.Windows.Forms.Button();
			this.btnUp = new System.Windows.Forms.Button();
			this.btnDown = new System.Windows.Forms.Button();
			this.listView1 = new System.Windows.Forms.ListView();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.SuspendLayout();
			// 
			// propertyGrid1
			// 
			this.propertyGrid1.CommandsVisibleIfAvailable = true;
			this.propertyGrid1.LargeButtons = false;
			this.propertyGrid1.LineColor = System.Drawing.SystemColors.ScrollBar;
			this.propertyGrid1.Location = new System.Drawing.Point(286, 24);
			this.propertyGrid1.Name = "propertyGrid1";
			this.propertyGrid1.Size = new System.Drawing.Size(266, 248);
			this.propertyGrid1.TabIndex = 0;
			this.propertyGrid1.Text = "propertyGrid1";
			this.propertyGrid1.ToolbarVisible = false;
			this.propertyGrid1.ViewBackColor = System.Drawing.SystemColors.Window;
			this.propertyGrid1.ViewForeColor = System.Drawing.SystemColors.WindowText;
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(289, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(96, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Properties:";
			// 
			// btnOK
			// 
			this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnOK.Location = new System.Drawing.Point(302, 296);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(80, 24);
			this.btnOK.TabIndex = 2;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCancel.Location = new System.Drawing.Point(387, 297);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(80, 24);
			this.btnCancel.TabIndex = 3;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// btnHelp
			// 
			this.btnHelp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnHelp.Location = new System.Drawing.Point(472, 296);
			this.btnHelp.Name = "btnHelp";
			this.btnHelp.Size = new System.Drawing.Size(80, 24);
			this.btnHelp.TabIndex = 4;
			this.btnHelp.Text = "Help";
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(16, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(88, 16);
			this.label2.TabIndex = 6;
			this.label2.Text = "Members:";
			// 
			// btnAdd
			// 
			this.btnAdd.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAdd.Location = new System.Drawing.Point(27, 248);
			this.btnAdd.Name = "btnAdd";
			this.btnAdd.Size = new System.Drawing.Size(80, 24);
			this.btnAdd.TabIndex = 7;
			this.btnAdd.Text = "Add";
			this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
			// 
			// btnRemove
			// 
			this.btnRemove.Enabled = false;
			this.btnRemove.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnRemove.Location = new System.Drawing.Point(126, 248);
			this.btnRemove.Name = "btnRemove";
			this.btnRemove.Size = new System.Drawing.Size(80, 24);
			this.btnRemove.TabIndex = 8;
			this.btnRemove.Text = "Remove";
			this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);
			// 
			// btnUp
			// 
			this.btnUp.Enabled = false;
			this.btnUp.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnUp.Image = ((System.Drawing.Image)(resources.GetObject("btnUp.Image")));
			this.btnUp.Location = new System.Drawing.Point(225, 32);
			this.btnUp.Name = "btnUp";
			this.btnUp.Size = new System.Drawing.Size(32, 32);
			this.btnUp.TabIndex = 9;
			this.btnUp.Click += new System.EventHandler(this.btnUp_Click);
			// 
			// btnDown
			// 
			this.btnDown.Enabled = false;
			this.btnDown.Image = ((System.Drawing.Image)(resources.GetObject("btnDown.Image")));
			this.btnDown.Location = new System.Drawing.Point(225, 72);
			this.btnDown.Name = "btnDown";
			this.btnDown.Size = new System.Drawing.Size(32, 32);
			this.btnDown.TabIndex = 10;
			this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
			// 
			// listView1
			// 
			this.listView1.FullRowSelect = true;
			this.listView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
			this.listView1.Location = new System.Drawing.Point(16, 24);
			this.listView1.MultiSelect = false;
			this.listView1.Name = "listView1";
			this.listView1.Size = new System.Drawing.Size(200, 216);
			this.listView1.TabIndex = 11;
			this.listView1.View = System.Windows.Forms.View.Details;
			this.listView1.Click += new System.EventHandler(this.listView1_Click);
			this.listView1.ItemActivate += new System.EventHandler(this.listView1_ItemActivate);
			this.listView1.BackColorChanged += new System.EventHandler(this.listView1_BackColorChanged);
			this.listView1.CursorChanged += new System.EventHandler(this.listView1_CursorChanged);
			this.listView1.SelectedIndexChanged += new System.EventHandler(this.listView1_SelectedIndexChanged);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// frmCollectionClone
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(570, 342);
			this.ControlBox = false;
			this.Controls.Add(this.listView1);
			this.Controls.Add(this.btnDown);
			this.Controls.Add(this.btnUp);
			this.Controls.Add(this.btnRemove);
			this.Controls.Add(this.btnAdd);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btnHelp);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.propertyGrid1);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCollectionClone";
			this.Text = "frmCollectionClone";
			this.Resize += new System.EventHandler(this.frmCollectionClone_Resize);
			this.Activated += new System.EventHandler(this.frmCollectionClone_Activated);
			this.ResumeLayout(false);

		}
		#endregion

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnAdd_Click(object sender, System.EventArgs e)
		{
			int x;
			string str="";
			switch (m_strType)
			{
                case "VS":
					str="VariableSubtitute";
					break;
				case "DATAVARSUB":
					str = "VariableSubstitute";
					break;
			}
			this.propertyGrid1.Hide();
			System.Drawing.SizeF StringLength = this.CreateGraphics().MeasureString(this.listView1.Items.Count.ToString().Trim(),this.listView1.Font);
			int intWidth= (int)((int)StringLength.Width * 1.5);  

            this.listView1.Columns[0].Width = intWidth;
			this.listView1.Columns[0].TextAlign = HorizontalAlignment.Right;
			
			System.Windows.Forms.ListViewItem entryListItem =
				this.listView1.Items.Add(this.listView1.Items.Count.ToString());
			entryListItem.BackColor = System.Drawing.Color.LightGray;
			entryListItem.ForeColor = System.Drawing.Color.Black;
			entryListItem.UseItemStyleForSubItems=false;


			ListViewItem.ListViewSubItem RowSubItem = 
				entryListItem.SubItems.Add(str + listView1.Items.Count.ToString());
			RowSubItem.ForeColor = System.Drawing.Color.Black;
			RowSubItem.BackColor = System.Drawing.Color.White;
			
			if (m_strType=="VS")
			{
				this.m_SQLVarSubItem1 = new FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem();
					
				this.propertyGrid2 = new PropertyGrid();
				propertyGrid2.PropertyValueChanged += new System.Windows.Forms.PropertyValueChangedEventHandler(this.PropertyGrid2_PropertyValueChanged);
				this.propertyGrid2.Name=this.m_intPropertyGridCount.ToString().Trim();
				this.propertyGrid2.ToolbarVisible=false;
				this.Controls.Add(this.propertyGrid2);
				this.propertyGrid2.Location=this.propertyGrid1.Location;
				this.propertyGrid2.Size = this.propertyGrid1.Size;
				this.PropertyGrid_Collection1.Add(propertyGrid2);

				
				this.m_SQLVarSubItem1.Index=this.listView1.Items.Count-1;
				this.m_SQLVarSubItem1.VariableName = str + listView1.Items.Count.ToString();
				this.m_SQLVarSubItem1.PropertyGridName=this.m_intPropertyGridCount.ToString().Trim();
				PropertyGrid_Collection1.Item(PropertyGrid_Collection1.Count-1).SelectedObject = this.m_SQLVarSubItem1;
				this.SQLMacroSubstitutionVariable_NewCollection1.Add(this.m_SQLVarSubItem1);
			}
			for (x=0;x<=this.PropertyGrid_Collection1.Count-1;x++)
			{
				this.PropertyGrid_Collection1.Item(x).Hide();
			}
			PropertyGrid_Collection1.Item(PropertyGrid_Collection1.Count-1).ExpandAllGridItems();
			PropertyGrid_Collection1.Item(PropertyGrid_Collection1.Count-1).Show();
			m_intCurrGrid=PropertyGrid_Collection1.Count-1;
			this.btnRemove.Enabled=true;
			this.btnDown.Enabled=false;
			if (listView1.Items.Count > 1)
			{
				this.btnUp.Enabled=true;
			}
			else
			{
				this.btnUp.Enabled=false;
			}
		    this.listView1.Items[this.listView1.Items.Count-1].Selected=true;
            EnableUpDownButtons();
			this.m_intPropertyGridCount++;
			


			
		}


		public void LoadProperties()
		{
			int x;
			int y;
			string str="";
			bool bNull=true;
			if (m_strType=="VS")   //variable substitution
			{
				if (frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count > 0)
				{
					this.listView1.Items.Clear();
					this.propertyGrid1.Hide();
					bNull=false;
					for (x=0;x<=frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count-1;x++)
					{
						//find x
						for (y=0;y<=frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count-1;y++)
						{
							if (x==frmMain.g_oSQLMacroSubstitutionVariable_Collection.Item(y).Index)
							{
								//load the current row group collection in the temporary new row group collection
								this.SQLMacroSubstitutionVariable_NewCollection1.Add(frmMain.g_oSQLMacroSubstitutionVariable_Collection.Item(y));
								System.Drawing.SizeF StringLength = this.CreateGraphics().MeasureString(this.listView1.Items.Count.ToString().Trim(),this.listView1.Font);
								int intWidth= (int)((int)StringLength.Width * 1.5);  

								this.listView1.Columns[0].Width = intWidth;
								this.listView1.Columns[0].TextAlign = HorizontalAlignment.Right;
			
								System.Windows.Forms.ListViewItem entryListItem =
									this.listView1.Items.Add(this.listView1.Items.Count.ToString());
								entryListItem.BackColor = System.Drawing.Color.LightGray;
								entryListItem.ForeColor = System.Drawing.Color.Black;
								entryListItem.UseItemStyleForSubItems=false;


								if (this.SQLMacroSubstitutionVariable_NewCollection1.Item(this.SQLMacroSubstitutionVariable_NewCollection1.Count-1).VariableName.Trim().Length > 0)
								{
									str=this.SQLMacroSubstitutionVariable_NewCollection1.Item(this.SQLMacroSubstitutionVariable_NewCollection1.Count-1).VariableName.Trim();
								}
								else
								{
									str="VariableSubstitution" + Convert.ToString(this.SQLMacroSubstitutionVariable_NewCollection1.Item(this.SQLMacroSubstitutionVariable_NewCollection1.Count-1).Index + 1).Trim();
								}
								ListViewItem.ListViewSubItem RowSubItem = 
									entryListItem.SubItems.Add(str);
								RowSubItem.ForeColor = System.Drawing.Color.Black;
								RowSubItem.BackColor = System.Drawing.Color.White;
						
								this.propertyGrid2 = new PropertyGrid();
								propertyGrid2.PropertyValueChanged += new System.Windows.Forms.PropertyValueChangedEventHandler(this.PropertyGrid2_PropertyValueChanged);
								
								this.propertyGrid2.Name=this.m_intPropertyGridCount.ToString().Trim();
								this.propertyGrid2.ToolbarVisible=false;
								this.Controls.Add(this.propertyGrid2);
								this.propertyGrid2.Location=this.propertyGrid1.Location;
								this.propertyGrid2.Size = this.propertyGrid1.Size;
								this.PropertyGrid_Collection1.Add(propertyGrid2);
								this.SQLMacroSubstitutionVariable_NewCollection1.Item(this.SQLMacroSubstitutionVariable_NewCollection1.Count-1).Index=this.listView1.Items.Count-1;
								this.SQLMacroSubstitutionVariable_NewCollection1.Item(this.SQLMacroSubstitutionVariable_NewCollection1.Count-1).PropertyGridName=this.m_intPropertyGridCount.ToString().Trim();
								PropertyGrid_Collection1.Item(PropertyGrid_Collection1.Count-1).SelectedObject = 
									this.SQLMacroSubstitutionVariable_NewCollection1.Item(this.SQLMacroSubstitutionVariable_NewCollection1.Count-1);
								this.m_intPropertyGridCount++;
								break;
							}
						}
					}
				}
			}

			if (bNull==false)
			{
				

				for (x=0;x<=this.PropertyGrid_Collection1.Count-1;x++)
				{
					this.PropertyGrid_Collection1.Item(x).Hide();
					this.PropertyGrid_Collection1.Item(x).ExpandAllGridItems();
				}
			
				if (this._intIndexToDisplay==-1 || this.listView1.Items.Count-1 < this._intIndexToDisplay)
				{
					PropertyGrid_Collection1.Item(PropertyGrid_Collection1.Count-1).Show();
					m_intCurrGrid=PropertyGrid_Collection1.Count-1;
					
					this.btnDown.Enabled=false;
					if (listView1.Items.Count > 1)
					{
						this.btnUp.Enabled=true;
					}
					else
					{
						this.btnUp.Enabled=false;
					}
					this.listView1.Items[this.listView1.Items.Count-1].Selected=true;
				}
				else
				{
					
					PropertyGrid_Collection1.Item(this._intIndexToDisplay).Show();
					listView1.Items[this._intIndexToDisplay].Selected=true;

				}

				EnableUpDownButtons();
			}
			if (listView1.Items.Count > 0) this.btnRemove.Enabled=true;
			
		}

		private void listView1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			
			
			
			if (listView1.SelectedItems.Count > 0)
			{
				ListViewItem.ListViewSubItem RowSubItem = listView1.Items[this.m_intCurrSelect].SubItems[1];
				RowSubItem.BackColor = System.Drawing.Color.White;
				RowSubItem.ForeColor = System.Drawing.Color.Black;
				
				ListViewItem entryListItem = listView1.SelectedItems[0];
				entryListItem.BackColor = System.Drawing.Color.LightGray;
				entryListItem.ForeColor = System.Drawing.Color.Black;
				entryListItem.UseItemStyleForSubItems=false;

				RowSubItem = entryListItem.SubItems[1];
				RowSubItem.BackColor = System.Drawing.Color.Blue;
				RowSubItem.ForeColor = System.Drawing.Color.White;
				m_intCurrSelect = listView1.SelectedItems[0].Index;
				if (this.m_strType=="VS")   //excel Row collection data
				{
					for (int y=0;y<=this.SQLMacroSubstitutionVariable_NewCollection1.Count-1;y++)
					{
						if (this.SQLMacroSubstitutionVariable_NewCollection1.Item(y).Index==m_intCurrSelect)
						{
							this.ShowPropertyGrid(SQLMacroSubstitutionVariable_NewCollection1.Item(y).PropertyGridName);
						}
						else
						{
							this.PropertyGrid_Collection1.Item(y).Hide();
						}
						
					}
				}
				
			listView1.SelectedItems.Clear();
			}
			EnableUpDownButtons();
		}
		private void listView1_ItemActivate(object sender, System.EventArgs e)
		{
			MessageBox.Show(listView1.FocusedItem.ToString());
		}

		private void listView1_Click(object sender, System.EventArgs e)
		{
			
		}

		private void listView1_CursorChanged(object sender, System.EventArgs e)
		{
			MessageBox.Show("CursorChanged");
		}

		private void listView1_BackColorChanged(object sender, System.EventArgs e)
		{
			MessageBox.Show("backcolorchanged");
		}

		

		public class PropertyGrid_Collection : System.Collections.CollectionBase
		{
			public PropertyGrid_Collection()
			{
				//
				// TODO: Add constructor logic here
				//
			}

			public void Add(System.Windows.Forms.PropertyGrid PropertyGrid1)
			{
				// vérify if object is not already in
				if (this.List.Contains(PropertyGrid1))
					throw new InvalidOperationException();
 
				// adding it
				this.List.Add(PropertyGrid1);
 
				// return collection
				//return this;
			}
			public void Remove(int index)
			{
				// Check to see if there is a widget at the supplied index.
				if (index > Count - 1 || index < 0)
					// If no widget exists, a messagebox is shown and the operation 
					// is cancelled.
				{
					System.Windows.Forms.MessageBox.Show("Index not valid!");
				}
				else
				{
					List.RemoveAt(index); 
				}
			}
			public PropertyGrid Item(int Index)
			{
				// The appropriate item is retrieved from the List object and
				// explicitly cast to the Widget type, then returned to the 
				// caller.
				return (PropertyGrid) List[Index];
			}




		}
		private void PropertyGrid2_PropertyValueChanged(object s, System.Windows.Forms.PropertyValueChangedEventArgs e)
		{
			bool bReset=false;
			int x;
			int y;
			//
			//SQL Variable Substitution Properties
			//
			if (this.m_strType=="VS")
			{
			    //do not change the variablename if it is a reserved word
				if (e.ChangedItem.Label.ToString().ToUpper().Trim()=="VARIABLENAME")
				{
					//see if the user attempted to change the name of a reserved macro variable name
					switch (e.OldValue.ToString().Trim().ToUpper())
					{
						case "COUNTYTABLE":
							bReset=true;
							break;
						case "SITETREETABLE":
							bReset=true;
							break;
						case "TREETABLE":
							bReset=true;
							break;
						case "CONDTABLE":
							bReset=true;
							break;
					    case "PLOTTABLE":
							bReset=true;
							break;
						case "DWMTABLE":
							bReset=true;
							break;
						case "CWDTABLE":
							bReset=true;
							break;
						case "FWDTABLE":
							bReset=true;
							break;
						case "PPSATABLE":
							bReset=true;
							break;
						case "POPSTRATUMTABLE":
							bReset=true;
							break;
						case "POPEVALTABLE":
							bReset=true;
							break;
						case "POPEVALGRPTABLE":
							bReset=true;
							break;
						case "POPESTNUNITTABLE":
							bReset=true;
							break;
						case "FORESTTYPETABLE":
							bReset=true;
							break;
						case "SPECIESTABLE":
							bReset=true;
							break;
						case "TREESPECIESGROUPTABLE":
							bReset=true;
							break;
						case "EUSAREAADJFACTOR":
							bReset=true;
							break;
						case "PPSAEUSCONNECTION":
							bReset=true;
							break;
						case "PPSACONDCONNECTION":
							bReset=true;
							break;
						case "CONDTREECONNECTION":
							bReset=true;
							break;
						case "STATECD":
							bReset=true;
							break;
						case "STATECD_LIST":
							bReset=true;
							break;
						case "EVALID":
							bReset=true;
							break;
						case "EVALID_LIST":
							bReset=true;
							break;
                        case "HARDWOOD":
							bReset=true;
							break;
                        case "SOFTWOOD":
							bReset=true;
							break;
						case "EUSTREETPA":
							bReset=true;
							break;
						case "EUSAREAADJFACTORBYTREEUNADJVALUE":
							bReset=true;
							break;
					}
					if (bReset)
						//reset the reserved macro variable name
						e.ChangedItem.PropertyDescriptor.SetValue(this.PropertyGrid_Collection1.Item(this.m_intCurrGrid).SelectedObject,e.OldValue);
					else
					{
						//check for a duplicate name
						y=0;
						for (x=0;x<=this.SQLMacroSubstitutionVariable_NewCollection1.Count-1;x++)
						{
							if (e.ChangedItem.Value.ToString().Trim().ToUpper() == 
								this.SQLMacroSubstitutionVariable_NewCollection1.Item(x).VariableName.Trim().ToUpper())
							{
								y++;
								if (y==2)
								{

									MessageBox.Show("Duplicate macro variable names are not allowed","QATools",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Information);	
									e.ChangedItem.PropertyDescriptor.SetValue(this.PropertyGrid_Collection1.Item(this.m_intCurrGrid).SelectedObject,e.OldValue);
									break;
								}
							}
						}
						
					}
				}
                
				if (this.SQLMacroSubstitutionVariable_NewCollection1.Item(m_intCurrGrid).VariableName.Trim().Length > 0)
				{
					bReset=false;
					switch (this.SQLMacroSubstitutionVariable_NewCollection1.Item(m_intCurrGrid).VariableName.Trim().ToUpper())
					{
						case "COUNTYTABLE":
							bReset=true;
							break;
						case "SITETREETABLE":
							bReset=true;
							break;
						case "TREETABLE":
							bReset=true;
							break;
						case "CONDTABLE":
							bReset=true;
							break;
						case "PLOTTABLE":
							bReset=true;
							break;
						case "DWMTABLE":
							bReset=true;
							break;
						case "CWDTABLE":
							bReset=true;
							break;
						case "FWDTABLE":
							bReset=true;
							break;
						case "PPSATABLE":
							bReset=true;
							break;
						case "POPSTRATUMTABLE":
							bReset=true;
							break;
						case "POPEVALTABLE":
							bReset=true;
							break;
						case "POPEVALGRPTABLE":
							bReset=true;
							break;
						case "POPESTNUNITTABLE":
							bReset=true;
							break;
						case "FORESTTYPETABLE":
							bReset=true;
							break;
						case "SPECIESTABLE":
							bReset=true;
							break;
						case "TREESPECIESGROUPTABLE":
							bReset=true;
							break;
						case "TREEREGIONALBIOMASSTABLE":
							bReset=true;
							break;
                        case "STATECD":
							bReset=true;
							break;
						case "EVALID":
							bReset=true;
							break;
						case "HARDWOOD":
							bReset=true;
							break;
						case "SOFTWOOD":
							bReset=true;
							break;


					}
					if (bReset)
					{
						e.ChangedItem.PropertyDescriptor.SetValue(this.PropertyGrid_Collection1.Item(this.m_intCurrGrid).SelectedObject,e.OldValue);
					}
					this.listView1.Items[m_intCurrSelect].SubItems[1].Text = this.SQLMacroSubstitutionVariable_NewCollection1.Item(m_intCurrGrid).VariableName;
				}
				else
				{
					this.listView1.Items[m_intCurrSelect].SubItems[1].Text= "VariableSubstitute" + Convert.ToString(this.SQLMacroSubstitutionVariable_NewCollection1.Item(m_intCurrGrid).Index+1).Trim();
				}

			}
			//
			//Label Property
			//
            if (e.ChangedItem.Label.ToString().ToUpper() == "NAME")
            {
            }

			this.PropertyGrid_Collection1.Item(m_intCurrGrid).Refresh();
		}

		private void btnRemove_Click(object sender, System.EventArgs e)
		{
			RemoveOne();
		}
		private void RemoveOne()
		{
			 bool bRemove;
			 int x;
			 int y;
			 int intSelect;
				/**********************************************
				  **lets see if we have one to remove
				  **********************************************/
					int index=this.m_intCurrSelect;
			        intSelect=index;
			if (index >= 0) 
			{

				if (this.m_strType=="VS")   //variable substution
				{
					bRemove=true;

					//check if this is a reserved variable
					switch (this.SQLMacroSubstitutionVariable_NewCollection1.Item(m_intCurrGrid).VariableName.Trim().ToUpper())
					{
						case "COUNTYTABLE":
							bRemove=false;
							break;
						case "SITETREETABLE":
							bRemove=false;
							break;
						case "TREETABLE":
							bRemove=false;
							break;
						case "CONDTABLE":
							bRemove=false;
							break;
						case "PLOTTABLE":
							bRemove=false;
							break;
						case "DWMTABLE":
							bRemove=false;
							break;
						case "CWDTABLE":
							bRemove=false;
							break;
						case "FWDTABLE":
							bRemove=false;
							break;
						case "PPSATABLE":
							bRemove=false;
							break;
						case "POPSTRATUMTABLE":
							bRemove=false;
							break;
						case "POPEVALTABLE":
							bRemove=false;
							break;
						case "POPEVALGRPTABLE":
							bRemove=false;
							break;
						case "POPESTNUNITTABLE":
							bRemove=false;
							break;
						case "FORESTTYPETABLE":
							bRemove=false;
							break;
						case "SPECIESTABLE":
							bRemove=false;
							break;
						case "TREESPECIESGROUPTABLE":
							bRemove=false;
							break;
                        case "TREEREGIONALBIOMASSTABLE":
							bRemove=false;
							break;
						case "EUSAREAADJFACTOR":
							bRemove=false;
							break;
						case "PPSAEUSCONNECTION":
							bRemove=false;
							break;
						case "PPSACONDCONNECTION":
							bRemove=false;
							break;
						case "CONDTREECONNECTION":
							bRemove=false;
							break;
					}
					if (!bRemove) MessageBox.Show("Reserved macro variables cannot be removed","QATools",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Information);
					else
					{
						//locate the current property associated with the listview
						for (x=0;x<=this.SQLMacroSubstitutionVariable_NewCollection1.Count-1;x++)
						{
							if (this.SQLMacroSubstitutionVariable_NewCollection1.Item(x).Index==index)
							{
								//remove the property grid
								this.RemovePropertyGrid(this.SQLMacroSubstitutionVariable_NewCollection1.Item(x).PropertyGridName);
								this.SQLMacroSubstitutionVariable_NewCollection1.Remove(x);
								//subtract 1 from the index of each item below the one we just removed
								for (y=0;y<=this.SQLMacroSubstitutionVariable_NewCollection1.Count-1;y++)
								{
									if (this.SQLMacroSubstitutionVariable_NewCollection1.Item(y).Index > index)
									{
										this.SQLMacroSubstitutionVariable_NewCollection1.Item(y).Index = this.SQLMacroSubstitutionVariable_NewCollection1.Item(y).Index - 1;
									}
								}									
								break;
							}
						}
						/**********************************************
								  **remove the ONE that is selected
								  **********************************************/
						if (index == 0 && listView1.Items.Count==1) 
						{
                            
							listView1.Items.Clear();
							btnRemove.Enabled=false;
						}
						else 
						{
							
							//*see if were at the top of the list
							if (index == 0 && listView1.Items.Count > 2) 
							{
								intSelect=0;
								//look for index 1
								for (x=0;x<=this.SQLMacroSubstitutionVariable_NewCollection1.Count-1;x++)
								{
									if (this.SQLMacroSubstitutionVariable_NewCollection1.Item(x).Index==index)
									{
										this.ShowPropertyGrid(SQLMacroSubstitutionVariable_NewCollection1.Item(x).PropertyGridName);
										break;
									}
								}
							}
							else 
							{
								
								//*see if were at the bottom
								if (index+1==listView1.Items.Count) 
								{
									//look for 1 index above the last item
									for (x=0;x<=this.SQLMacroSubstitutionVariable_NewCollection1.Count-1;x++)
									{
										if (this.SQLMacroSubstitutionVariable_NewCollection1.Item(x).Index==index-1)
										{
											this.ShowPropertyGrid(SQLMacroSubstitutionVariable_NewCollection1.Item(x).PropertyGridName);												
											break;
										}
									}
									this.m_intCurrSelect=index-1;
									intSelect=index-1;
								}
								else
								{
									//look for 1 index above the last item
									for (x=0;x<=this.SQLMacroSubstitutionVariable_NewCollection1.Count-1;x++)
									{
										if (this.SQLMacroSubstitutionVariable_NewCollection1.Item(x).Index==index)
										{
											this.ShowPropertyGrid(SQLMacroSubstitutionVariable_NewCollection1.Item(x).PropertyGridName);
											break;
										}
									}
									intSelect=index;
								}
							}
							this.listView1.Items.Remove(listView1.Items[index]);
						}
						for (x=index; x<=listView1.Items.Count-1;x++) 
						{
							listView1.Items[x].SubItems[0].Text = x.ToString();
						}
					}
					
				}


			}
			if (listView1.Items.Count > 0) listView1.Items[intSelect].Selected=true;
			System.Drawing.SizeF StringLength = this.CreateGraphics().MeasureString(Convert.ToString(this.listView1.Items.Count-1).Trim(),this.listView1.Font);
			int intWidth= (int)((int)StringLength.Width * 1.5);  
			if (this.PropertyGrid_Collection1.Count==0) this.propertyGrid1.Show();
			this.listView1.Columns[0].Width = intWidth;
			this.listView1.Columns[0].TextAlign = HorizontalAlignment.Right;
			EnableUpDownButtons();
			
		}
		private void EnableUpDownButtons()
		{
			if (listView1.Items.Count==0 || listView1.Items.Count==1)
			{
				this.btnDown.Enabled=false;
				this.btnUp.Enabled=false;
			}
			else if (this.m_intCurrSelect==0 && listView1.Items.Count>0)
			{
				this.btnDown.Enabled=true;
				this.btnUp.Enabled=false;
			}
			else if (this.m_intCurrSelect==listView1.Items.Count-1)
			{
				this.btnDown.Enabled=false;
				this.btnUp.Enabled=true;
			}
			else
			{
				this.btnDown.Enabled=true;
				this.btnUp.Enabled=true;

			}
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
		    int x;
			//frmMain.g_bSave=true;

			if (this.m_strType=="VS")  //sql macro variable substituion
			{
				//clear all the Row objects
				for (x=frmMain.g_oSQLMacroSubstitutionVariable_Collection.Count-1;x>=0;x--)
				{
					frmMain.g_oSQLMacroSubstitutionVariable_Collection.Remove(x);
				}
				//save the new objects
				for (x=0;x<=this.SQLMacroSubstitutionVariable_NewCollection1.Count-1;x++)
				{
					frmMain.g_oSQLMacroSubstitutionVariable_Collection.Add(this.SQLMacroSubstitutionVariable_NewCollection1.Item(x));
				}
				//clear the collection
				for (x=0;x<=this.SQLMacroSubstitutionVariable_NewCollection1.Count-1;x++)
				{
					this.SQLMacroSubstitutionVariable_NewCollection1.Remove(x);
				}
			}







			for (x=this.PropertyGrid_Collection1.Count-1;x>=0;x--)
			{
				this.PropertyGrid_Collection1.Remove(x);
			}
			this.Close();


		}

		private void btnUp_Click(object sender, System.EventArgs e)
		{
			int intSaveCurrSelect = this.m_intCurrSelect;
			string str1;
			string str2;
			int int1=0;
			int int2=0;
			int x;
			//move the current selection up one
			str1 = this.listView1.Items[intSaveCurrSelect-1].SubItems[1].Text;
			str2 = this.listView1.Items[intSaveCurrSelect].SubItems[1].Text;
			this.listView1.Items[intSaveCurrSelect].SubItems[1].Text=str1;
			this.listView1.Items[intSaveCurrSelect-1].SubItems[1].Text=str2;
			
			if (this.m_strType=="VS")  //variable substitution
			{
				for (x=0;x<=this.SQLMacroSubstitutionVariable_NewCollection1.Count-1;x++)
				{
					if (this.SQLMacroSubstitutionVariable_NewCollection1.Item(x).Index==intSaveCurrSelect)
					{
						int2=x;
					}
					else
					{
						if (this.SQLMacroSubstitutionVariable_NewCollection1.Item(x).Index==intSaveCurrSelect-1)
						{
							int1=x;
						}
					}
				}
				//swap locations with the one directly above the currently selected item
				this.SQLMacroSubstitutionVariable_NewCollection1.Item(int1).Index = intSaveCurrSelect;
				this.SQLMacroSubstitutionVariable_NewCollection1.Item(int2).Index = intSaveCurrSelect-1;
			}
            



			this.PropertyGrid_Collection1.Item(this.m_intCurrGrid).Refresh();
			this.listView1.Items[intSaveCurrSelect-1].Selected=true;

		}

		private void btnDown_Click(object sender, System.EventArgs e)
		{
			int intSaveCurrSelect = this.m_intCurrSelect;
			string str1;
			string str2;
			int int1=0;
			int int2=0;
			int x;
			//move the current selection down one
			str1 = this.listView1.Items[intSaveCurrSelect+1].SubItems[1].Text;
			str2 = this.listView1.Items[intSaveCurrSelect].SubItems[1].Text;
			this.listView1.Items[intSaveCurrSelect].SubItems[1].Text=str1;
			this.listView1.Items[intSaveCurrSelect+1].SubItems[1].Text=str2;
			if (this.m_strType=="VS")  //variable substitution
			{
				for (x=0;x<=this.SQLMacroSubstitutionVariable_NewCollection1.Count-1;x++)
				{
					if (this.SQLMacroSubstitutionVariable_NewCollection1.Item(x).Index==intSaveCurrSelect)
					{
						int2=x;
					}
					else
					{
						if (this.SQLMacroSubstitutionVariable_NewCollection1.Item(x).Index==intSaveCurrSelect+1)
						{
							int1=x;
						}
					}
				}
				//swap locations with the one directly above the currently selected item
				this.SQLMacroSubstitutionVariable_NewCollection1.Item(int1).Index = intSaveCurrSelect;
				this.SQLMacroSubstitutionVariable_NewCollection1.Item(int2).Index = intSaveCurrSelect+1;
			}
            



			this.PropertyGrid_Collection1.Item(this.m_intCurrGrid).Refresh();
			this.listView1.Items[intSaveCurrSelect+1].Selected=true;
		}

		private void frmCollectionClone_Resize(object sender, System.EventArgs e)
		{
			int x;

			//left
			this.btnHelp.Left = this.ClientSize.Width - listView1.Left - this.btnHelp.Width;
			this.btnCancel.Left = this.btnHelp.Left - this.btnCancel.Width - 5;
			this.btnOK.Left = this.btnCancel.Left - this.btnOK.Width - 5;

			//top
			this.btnOK.Top = this.ClientSize.Height - 24 - this.btnOK.Height;
			this.btnCancel.Top = this.btnOK.Top;
			this.btnHelp.Top = this.btnOK.Top;

			this.propertyGrid1.Width = this.ClientSize.Width - this.propertyGrid1.Left - listView1.Left;
			this.propertyGrid1.Height = this.btnOK.Top - listView1.Top -  24;
			for (x=0;x<=this.PropertyGrid_Collection1.Count-1;x++)
			{
				this.PropertyGrid_Collection1.Item(x).Size=this.propertyGrid1.Size;
			}

			this.listView1.Height = this.ClientSize.Height - 130;
			//listview buttons top
			this.btnAdd.Top = this.listView1.Height + this.listView1.Top + 8;
			this.btnRemove.Top = this.btnAdd.Top;
			

			
		}

		private void frmCollectionClone_Activated(object sender, System.EventArgs e)
		{
			this.frmCollectionClone_Resize(sender,e);
		}
		private void ShowPropertyGrid(string p_strName)
		{
			for (int x=0;x<=PropertyGrid_Collection1.Count-1;x++)
			{
				if (PropertyGrid_Collection1.Item(x).Name.Trim() == p_strName.Trim())
				{
					PropertyGrid_Collection1.Item(x).Show();
					this.m_intCurrGrid=x;
				}
				else
				{
					PropertyGrid_Collection1.Item(x).Hide();
				}
			}
		}
		private void RemovePropertyGrid(string p_strName)
		{
			for (int x=0;x<=PropertyGrid_Collection1.Count-1;x++)
			{
				if (PropertyGrid_Collection1.Item(x).Name.Trim() == p_strName.Trim())
				{
					PropertyGrid_Collection1.Item(x).Hide();
					PropertyGrid_Collection1.Remove(x);
					break;
				}
			}
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
		}
	
		public int IndexToDisplay
		{
			set {_intIndexToDisplay=value;}
			get {return _intIndexToDisplay;}
		}

		
		

	}
}

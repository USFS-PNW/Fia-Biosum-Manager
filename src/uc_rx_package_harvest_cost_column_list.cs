using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_rx_package_harvest_cost_column_list.
	/// </summary>
	public class uc_rx_package_harvest_cost_column_list : System.Windows.Forms.UserControl
	{
		public int m_intError=0;
		public string m_strError="";
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.ListView lvRxHarvestCostColumns;
		private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvAlternateRowColors=new ListViewAlternateBackgroundColors();
		const int COLUMN_NULL=0;
		const int COLUMN_PACKAGE=1;
		const int COLUMN_RX=2;
		const int COLUMN_CYCLE=3;
		const int COLUMN_FIELD=4;
		const int COLUMN_DESC=5;
		private System.Windows.Forms.Label lblDesc;
		private Queries m_oQueries = new Queries();
		private RxTools m_oRxTools = new RxTools();
		private string m_strColumnNameList="";
        private string m_strHarvestTableColumnNameList = "";
        private FIA_Biosum_Manager.frmRxPackageItem _frmRxPackageItem=null;
    

			
			
    
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FIA_Biosum_Manager.frmRxPackageItem ReferenceFormRxPackageItem
		{
			get {return this._frmRxPackageItem;}
			set {this._frmRxPackageItem=value;}

		}
		
		
		public uc_rx_package_harvest_cost_column_list()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.lvRxHarvestCostColumns.View = System.Windows.Forms.View.Details;
			this.m_oLvAlternateRowColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvAlternateRowColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvAlternateRowColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
			this.m_oLvAlternateRowColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
			this.m_oLvAlternateRowColors.ReferenceSelectedRowBackgroundColor=frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvAlternateRowColors.CustomFullRowSelect=true;
			this.m_oLvAlternateRowColors.ReferenceListView = lvRxHarvestCostColumns;
			if (frmMain.g_oGridViewFont != null) this.lvRxHarvestCostColumns.Font = frmMain.g_oGridViewFont;

			// TODO: Add any initialization after the InitializeComponent call

		}
		public void loadvalues()
		{
			int x,y,z;

			m_oQueries.m_oFvs.LoadDatasource=true;
			m_oQueries.m_oReference.LoadDatasource=true;
			m_oQueries.LoadDatasources(true);
			
			this.lvRxHarvestCostColumns.Clear();
			
			this.m_oLvAlternateRowColors.InitializeRowCollection();
			this.lvRxHarvestCostColumns.Columns.Add("",2,HorizontalAlignment.Left);
			this.lvRxHarvestCostColumns.Columns.Add("RxPackage",80,HorizontalAlignment.Left);
			this.lvRxHarvestCostColumns.Columns.Add("Rx", 80, HorizontalAlignment.Left);
			this.lvRxHarvestCostColumns.Columns.Add("SimCycle", 80, HorizontalAlignment.Left);
			this.lvRxHarvestCostColumns.Columns.Add("Harvest Cost Column", 200, HorizontalAlignment.Left);
			this.lvRxHarvestCostColumns.Columns.Add("Description", 300, HorizontalAlignment.Left);

			this.m_intError=0;
			this.m_strError="";

            for (x = 0; x <= this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count - 1; x++)
            {
                //
                //1st simulation
                //
                if (ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear1Rx.Trim() ==
                      ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim())
                {
                    if (ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceHarvestCostColumnCollection != null)
                    {
                        for (y = 0;
                             y <= ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceHarvestCostColumnCollection.Count - 1;
                             y++)
                        {
                            AddItemToList(1,
                                ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceHarvestCostColumnCollection.Item(y));
                        }
                    }
                }
                //
                //2nd simulation
                //
                if (ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear2Rx.Trim() ==
                      ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim())
                {
                    if (ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceHarvestCostColumnCollection != null)
                    {
                        for (y = 0;
                             y <= ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceHarvestCostColumnCollection.Count - 1;
                             y++)
                        {
                            AddItemToList(2,
                                ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceHarvestCostColumnCollection.Item(y));
                        }
                    }
                }
                //
                //3rd simulation
                //
                if (ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear3Rx.Trim() ==
                      ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim())
                {
                    if (ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceHarvestCostColumnCollection != null)
                    {
                        for (y = 0;
                             y <= ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceHarvestCostColumnCollection.Count - 1;
                             y++)
                        {
                            AddItemToList(3,
                                ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceHarvestCostColumnCollection.Item(y));
                        }
                    }
                }
                //
                //4th simulation
                //
                if (ReferenceFormRxPackageItem.ReferenceRxPackageItem.SimulationYear4Rx.Trim() ==
                      ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim())
                {
                    if (ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceHarvestCostColumnCollection != null)
                    {
                        for (y = 0;
                             y <= ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceHarvestCostColumnCollection.Count - 1;
                             y++)
                        {
                            AddItemToList(4,
                                ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).ReferenceHarvestCostColumnCollection.Item(y));
                        }
                    }
                }
            }
		
			
			this.m_oLvAlternateRowColors.ListView();

            this.ReOrderList();

          
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_rx_package_harvest_cost_column_list));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblDesc = new System.Windows.Forms.Label();
            this.lvRxHarvestCostColumns = new System.Windows.Forms.ListView();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(648, 392);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.lblDesc);
            this.panel1.Controls.Add(this.lvRxHarvestCostColumns);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 18);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(642, 371);
            this.panel1.TabIndex = 0;
            this.panel1.Resize += new System.EventHandler(this.panel1_Resize);
            // 
            // lblDesc
            // 
            this.lblDesc.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblDesc.Location = new System.Drawing.Point(0, 0);
            this.lblDesc.Name = "lblDesc";
            this.lblDesc.Size = new System.Drawing.Size(642, 51);
            this.lblDesc.TabIndex = 1;
            this.lblDesc.Text = resources.GetString("lblDesc.Text");
            this.lblDesc.Click += new System.EventHandler(this.lblDesc_Click);
            // 
            // lvRxHarvestCostColumns
            // 
            this.lvRxHarvestCostColumns.GridLines = true;
            this.lvRxHarvestCostColumns.HideSelection = false;
            this.lvRxHarvestCostColumns.Location = new System.Drawing.Point(8, 54);
            this.lvRxHarvestCostColumns.MultiSelect = false;
            this.lvRxHarvestCostColumns.Name = "lvRxHarvestCostColumns";
            this.lvRxHarvestCostColumns.Size = new System.Drawing.Size(616, 306);
            this.lvRxHarvestCostColumns.TabIndex = 0;
            this.lvRxHarvestCostColumns.UseCompatibleStateImageBehavior = false;
            this.lvRxHarvestCostColumns.View = System.Windows.Forms.View.Details;
            this.lvRxHarvestCostColumns.SelectedIndexChanged += new System.EventHandler(this.lvRxHarvestCostColumns_SelectedIndexChanged);
            this.lvRxHarvestCostColumns.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lvRxHarvestCostColumns_MouseUp);
            // 
            // uc_rx_package_harvest_cost_column_list
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_rx_package_harvest_cost_column_list";
            this.Size = new System.Drawing.Size(648, 392);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void lblDesc_Click(object sender, System.EventArgs e)
		{
		
		}

		private void panel1_Resize(object sender, System.EventArgs e)
		{
			this.lvRxHarvestCostColumns.Width = this.panel1.ClientSize.Width - this.lvRxHarvestCostColumns.Left - this.panel1.AutoScrollMargin.Width;
			this.lvRxHarvestCostColumns.Height = this.panel1.ClientSize.Height - this.lvRxHarvestCostColumns.Top - this.panel1.AutoScrollMargin.Height;
		}
       
        private void AddItemToList(int p_intCycle,FIA_Biosum_Manager.RxItemHarvestCostColumnItem oItem)
        {

            this.lvRxHarvestCostColumns.Items.Add("");
            this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].UseItemStyleForSubItems = false;
            for (int z = 1; z <= this.lvRxHarvestCostColumns.Columns.Count - 1; z++)
            {
                this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems.Add(" ");
            }

            this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems[COLUMN_PACKAGE].Text = this.ReferenceFormRxPackageItem.m_oRxPackageItem.RxPackageId;
            this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems[COLUMN_RX].Text = oItem.RxId;
            this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems[COLUMN_CYCLE].Text = p_intCycle.ToString().Trim();
            this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems[COLUMN_FIELD].Text = oItem.HarvestCostColumn;
            this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].SubItems[COLUMN_DESC].Text = oItem.Description;
           

            this.m_oLvAlternateRowColors.AddRow();
            this.m_oLvAlternateRowColors.AddColumns(lvRxHarvestCostColumns.Items.Count - 1, this.lvRxHarvestCostColumns.Columns.Count);

            
            this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items.Count - 1].Selected = true;
        }
        public void RemoveRxItemsFromList(string p_strRx, string p_strFvsCycle)
        {
            int x;
            int y;
            for (x = this.lvRxHarvestCostColumns.Items.Count - 1; x >= 0; x--)
            {

                if (this.lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_RX].Text.Trim() == p_strRx.Trim() &&
                    this.lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text.Trim() == p_strFvsCycle)
                    this.lvRxHarvestCostColumns.Items.Remove(lvRxHarvestCostColumns.Items[x]);

            }
        }
        public void AddRxItemsToList(string p_strRx, string p_strFvsCycle)
        {
            bool bFound = false;
            int x, y;
            for (x = this.lvRxHarvestCostColumns.Items.Count - 1; x >= 0; x--)
            {

                if (this.lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_RX].Text.Trim() == p_strRx.Trim() &&
                    this.lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text.Trim() == p_strFvsCycle)
                {
                    bFound = true;
                    break;
                }

            }
            if (bFound == false)
            {
                //remove any treatments that have the same cycle as our current treatment
                for (x = this.lvRxHarvestCostColumns.Items.Count - 1; x >= 0; x--)
                {
                    if (this.lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_RX].Text.Trim() != p_strRx.Trim() &&
                        this.lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text.Trim() == p_strFvsCycle)
                    {
                        this.lvRxHarvestCostColumns.Items.Remove(lvRxHarvestCostColumns.Items[x]);
                    }
                }
                for (x = 0; x <= this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Count - 1; x++)
                {
                    if (this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).RxId.Trim() ==
                        p_strRx.Trim())
                    {
                        for (y = 0; y <= ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).m_oHarvestCostColumnItem_Collection1.Count - 1; y++)
                        {
                            this.AddItemToList(Convert.ToInt32(p_strFvsCycle), this.ReferenceFormRxPackageItem.ReferenceRxItemCollection.Item(x).m_oHarvestCostColumnItem_Collection1.Item(y));
                        }

                    }
                }
            }
            //check to see if they need to be reordered
            //just concerned with rx items where later fvs cycles are before earlier fvs cycles
            int intFvsCycle = -1;
            for (x = 0; x <= this.lvRxHarvestCostColumns.Items.Count - 1; x++)
            {
                if (intFvsCycle == -1 && this.lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text.Trim().Length > 0)
                {
                    intFvsCycle = Convert.ToInt32(lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text);
                }
                if (lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text.Trim().Length > 0)
                {
                    if (Convert.ToInt32(lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text) > intFvsCycle)
                    {
                        ReOrderList();
                        break;
                    }
                }

            }
        }
        private void ReOrderList()
        {
            System.Windows.Forms.ListView oLv;
            int x, y, z;
            string strCurFvsCycle = "";
            int intSet = 0;
            int[] intSetArray = new int[this.lvRxHarvestCostColumns.Items.Count];
            //put the current list into sets
            for (x = 0; x <= this.lvRxHarvestCostColumns.Items.Count - 1; x++)
            {
                if (x == 0)
                {
                    intSet = 0;
                    intSetArray[x] = 0;
                    strCurFvsCycle = lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text;
                }
                else if (strCurFvsCycle != lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text)
                {
                    intSet++;
                    intSetArray[x] = intSet;
                    strCurFvsCycle = lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text;
                }
                else
                {
                    intSetArray[x] = intSet;
                }
            }
            //get the min and max cycle year values
            int intMin = 1000;
            int intMax = -1;
            for (x = 0; x <= this.lvRxHarvestCostColumns.Items.Count - 1; x++)
            {
                if (this.lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text.Trim().Length != 0)
                {
                    if (Convert.ToInt32(lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text) <
                        intMin)
                    {
                        intMin = Convert.ToInt32(lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text);
                    }
                    else if (Convert.ToInt32(lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text) >
                        intMax)
                    {
                        intMax = Convert.ToInt32(lvRxHarvestCostColumns.Items[x].SubItems[COLUMN_CYCLE].Text);
                    }
                }
            }
            //reorder the set 
            int intOrder = -1;
            string strCurrentFvsCycle = "";
            for (x = intMin; x <= intMax; x++)
            {
                for (y = 0; y <= this.lvRxHarvestCostColumns.Items.Count - 1; y++)
                {
                    //bypass package command
                    if (this.lvRxHarvestCostColumns.Items[y].SubItems[COLUMN_CYCLE].Text.Trim().Length != 0)
                    {
                        //check if the cycle exists
                        if (Convert.ToInt32(lvRxHarvestCostColumns.Items[y].SubItems[COLUMN_CYCLE].Text) == x)
                        {
                            //increment the order index
                            intOrder++;
                            //check to see if package command already equals the order index
                            for (z = 0; z <= this.lvRxHarvestCostColumns.Items.Count - 1; z++)
                            {
                                if (this.lvRxHarvestCostColumns.Items[z].SubItems[COLUMN_CYCLE].Text.Trim().Length == 0 &&
                                    intSetArray[z] == intOrder)
                                {
                                    //found package command that already equals order index so increment order index by 1
                                    intOrder++;
                                    break;
                                }
                            }

                            strCurrentFvsCycle = lvRxHarvestCostColumns.Items[y].SubItems[COLUMN_CYCLE].Text;
                            for (z = y; z <= this.lvRxHarvestCostColumns.Items.Count - 1; z++)
                            {
                                //everything with the current fvs cycle assign the new order index
                                if (strCurrentFvsCycle == lvRxHarvestCostColumns.Items[z].SubItems[COLUMN_CYCLE].Text)
                                {
                                    intSetArray[z] = intOrder;
                                }
                                else break;
                            }
                            break;

                        }

                    }
                }
            }

            //copy the new order to the work listview object
            oLv = new ListView();

            for (x = 0; x <= intSet; x++)
            {
                for (y = 0; y <= intSetArray.Length - 1; y++)
                {
                    if (intSetArray[y] == x)
                    {
                        oLv.Items.Add("");
                        for (z = 1; z <= this.lvRxHarvestCostColumns.Columns.Count - 1; z++)
                        {
                            oLv.Items[oLv.Items.Count - 1].SubItems.Add(this.lvRxHarvestCostColumns.Items[y].SubItems[z].Text);
                        }


                    }
                }
            }
            //copy the work listview values to the destination listview
            for (x = 0; x <= this.lvRxHarvestCostColumns.Items.Count - 1; x++)
            {
                for (z = 1; z <= this.lvRxHarvestCostColumns.Columns.Count - 1; z++)
                {
                    lvRxHarvestCostColumns.Items[x].SubItems[z].Text = oLv.Items[x].SubItems[z].Text;
                }
            }

            oLv.Dispose();


            
            //reinitialize the alternate row colors
            for (x = 0; x <= this.lvRxHarvestCostColumns.Items.Count - 1; x++)
            {
                for (y = 0; y <= this.lvRxHarvestCostColumns.Columns.Count - 1; y++)
                {
                    this.m_oLvAlternateRowColors.ListViewSubItem(lvRxHarvestCostColumns.Items[x].Index, y, lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.Items[x].Index].SubItems[y], false);
                }
            }







        }
        

        
		
        private void lvRxHarvestCostColumns_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvRxHarvestCostColumns.SelectedItems.Count > 0)
                m_oLvAlternateRowColors.DelegateListViewItem(lvRxHarvestCostColumns.SelectedItems[0]);
        }

        private void lvRxHarvestCostColumns_MouseUp(object sender, MouseEventArgs e)
        {
           
            try
            {
                if (e.Button == MouseButtons.Left)
                {
                    int intRowHt = lvRxHarvestCostColumns.Items[0].Bounds.Height;
                    double dblRow = (double)(e.Y / intRowHt);
                    this.lvRxHarvestCostColumns.Items[lvRxHarvestCostColumns.TopItem.Index + (int)dblRow - 1].Selected = true;

                }
            }
            catch
            {
            }
        }
        

		
	}
}

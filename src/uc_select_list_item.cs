using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Collections.Generic;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_select_table.
	/// </summary>
	public class uc_select_list_item : System.Windows.Forms.UserControl
	{
		public System.Windows.Forms.ListBox listBox1;
		private System.Windows.Forms.Button btnCancel;
		public System.Windows.Forms.Button btnOK;
		public System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
		public string strError;
		public int intError;
		public System.Windows.Forms.GroupBox groupBox2;
		public System.Windows.Forms.Label lblMsg;
		//private System.Windows.Forms.Form p_ParentForm;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_select_list_item()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
            //p_ParentForm = this.ParentForm;
			
			//MessageBox.Show(p_ParentForm.Name);
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
			this.lblMsg = new System.Windows.Forms.Label();
			this.listBox1 = new System.Windows.Forms.ListBox();
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnOK = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.lblTitle = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.SuspendLayout();
			// 
			// lblMsg
			// 
			this.lblMsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblMsg.Location = new System.Drawing.Point(8, 264);
			this.lblMsg.Name = "lblMsg";
			this.lblMsg.Size = new System.Drawing.Size(368, 24);
			this.lblMsg.TabIndex = 3;
			this.lblMsg.Text = "lblMsg";
			this.lblMsg.Visible = false;
			// 
			// listBox1
			// 
			this.listBox1.Location = new System.Drawing.Point(16, 24);
			this.listBox1.Name = "listBox1";
			this.listBox1.Size = new System.Drawing.Size(360, 199);
			this.listBox1.TabIndex = 0;
			// 
			// btnCancel
			// 
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(184, 229);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(64, 24);
			this.btnCancel.TabIndex = 2;
			this.btnCancel.Text = "Cancel";
			// 
			// btnOK
			// 
			this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnOK.Location = new System.Drawing.Point(120, 229);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(64, 24);
			this.btnOK.TabIndex = 1;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.groupBox2);
			this.groupBox1.Controls.Add(this.lblTitle);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(432, 368);
			this.groupBox1.TabIndex = 3;
			this.groupBox1.TabStop = false;
			this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.listBox1);
			this.groupBox2.Controls.Add(this.lblMsg);
			this.groupBox2.Controls.Add(this.btnCancel);
			this.groupBox2.Controls.Add(this.btnOK);
			this.groupBox2.Location = new System.Drawing.Point(16, 48);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(384, 296);
			this.groupBox2.TabIndex = 26;
			this.groupBox2.TabStop = false;
			// 
			// lblTitle
			// 
			this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
			this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTitle.ForeColor = System.Drawing.Color.Green;
			this.lblTitle.Location = new System.Drawing.Point(3, 16);
			this.lblTitle.Name = "lblTitle";
			this.lblTitle.Size = new System.Drawing.Size(426, 32);
			this.lblTitle.TabIndex = 25;
			this.lblTitle.Text = "Table Select";
			// 
			// uc_select_list_item
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_select_list_item";
			this.Size = new System.Drawing.Size(432, 368);
			this.Resize += new System.EventHandler(this.uc_select_list_item_Resize);
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		

		private void uc_select_list_item_Resize(object sender, System.EventArgs e)
		{

		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.OK;
		}
		public void Initialize_Width(string strLargestStringInList)
		{
			((frmDialog)this.ParentForm).Width = (int)this.CreateGraphics().MeasureString(strLargestStringInList,this.listBox1.Font).Width + 150;

		}

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.groupBox2.Left = 10 ;
				this.groupBox2.Width =  this.groupBox1.Width - 20;
				this.groupBox2.Height = this.groupBox1.Height - this.lblTitle.Top - this.lblTitle.Height -  5;
				this.listBox1.Height = this.groupBox2.Height - this.listBox1.Top - this.lblMsg.Height - this.btnOK.Height - 5;
				this.listBox1.Left = 10;
				this.listBox1.Width  = this.groupBox2.Width - 20;
				this.btnOK.Top = this.listBox1.Top +  this.listBox1.Height + 2;
				this.btnOK.Left = (int)(this.groupBox2.Width  * .50) - this.btnOK.Width;
				this.btnCancel.Top = this.btnOK.Top;
				this.btnCancel.Left = this.btnOK.Left + this.btnOK.Width;
				this.lblMsg.Top = this.btnOK.Top + this.btnOK.Height;
				this.lblMsg.Left = 5;
			}
			catch
			{
			}
		}
        public void loadvalues(string strConn, string strSelectSQL, string strColumn)
        {
        }
		/// <summary>
		/// Populate listbox with the strColumn values
		/// </summary>
		/// <param name="p_oDataMgr"></param>
		/// <param name="p_oConn"></param>
		/// <param name="strSQL"></param>
		/// <param name="strColumn"></param>
		public void loadvalues(SQLite.ADO.DataMgr p_oDataMgr,
							   System.Data.SQLite.SQLiteConnection p_oConn,
                               string strSQL, string strColumn)
        {
            listBox1.Items.Clear();
			p_oDataMgr.SqlQueryReader(p_oConn, strSQL);
            if (p_oDataMgr.m_DataReader.HasRows)
            {
                while (p_oDataMgr.m_DataReader.Read())
                {
                    if (p_oDataMgr.m_DataReader[strColumn] != System.DBNull.Value &&
						p_oDataMgr.m_DataReader[strColumn].ToString().Trim().Length > 0)
                    {
                        listBox1.Items.Add(p_oDataMgr.m_DataReader[strColumn].ToString().Trim());
                    }
                }
            }
            else
            {
                MessageBox.Show("No items to add to menu list", "FIA Bisoum");
            }
        }
        public void loadvalues(List<string> p_strList)
        {
            if (p_strList != null && p_strList[0] != null)
            {
                foreach (string strItem in p_strList)
                    listBox1.Items.Add(strItem);
            }
            else
            {
                MessageBox.Show("No items to add to menu list", "FIA Bisoum");
            }
        }
	}
}

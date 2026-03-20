using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_table_exists_dialog.
	/// </summary>
	public class uc_table_exists_dialog : System.Windows.Forms.UserControl
	{
		public System.Windows.Forms.RadioButton rdoOverwrite;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.RadioButton radioButton1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnCopy;
		private System.Windows.Forms.Button Cancel;
		private System.Windows.Forms.TextBox txtTable;
		public System.Windows.Forms.ListBox listBox1;
		private System.Windows.Forms.Label label2;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_table_exists_dialog()
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
			this.rdoOverwrite = new System.Windows.Forms.RadioButton();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.listBox1 = new System.Windows.Forms.ListBox();
			this.txtTable = new System.Windows.Forms.TextBox();
			this.Cancel = new System.Windows.Forms.Button();
			this.btnCopy = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.radioButton1 = new System.Windows.Forms.RadioButton();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// rdoOverwrite
			// 
			this.rdoOverwrite.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.rdoOverwrite.Location = new System.Drawing.Point(41, 72);
			this.rdoOverwrite.Name = "rdoOverwrite";
			this.rdoOverwrite.Size = new System.Drawing.Size(408, 32);
			this.rdoOverwrite.TabIndex = 4;
			this.rdoOverwrite.Text = "Overwrite the existing table";
			this.rdoOverwrite.Click += new System.EventHandler(this.rdoOverwrite_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.listBox1);
			this.groupBox1.Controls.Add(this.txtTable);
			this.groupBox1.Controls.Add(this.Cancel);
			this.groupBox1.Controls.Add(this.btnCopy);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.radioButton1);
			this.groupBox1.Controls.Add(this.rdoOverwrite);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(472, 392);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(16, 362);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(440, 16);
			this.label2.TabIndex = 6;
			this.label2.Text = "DB Tables";
			// 
			// listBox1
			// 
			this.listBox1.Location = new System.Drawing.Point(16, 272);
			this.listBox1.Name = "listBox1";
			this.listBox1.Size = new System.Drawing.Size(440, 82);
			this.listBox1.TabIndex = 5;
			// 
			// txtTable
			// 
			this.txtTable.Location = new System.Drawing.Point(48, 168);
			this.txtTable.Name = "txtTable";
			this.txtTable.Size = new System.Drawing.Size(368, 20);
			this.txtTable.TabIndex = 0;
			this.txtTable.Text = "";
			// 
			// Cancel
			// 
			this.Cancel.Location = new System.Drawing.Point(232, 206);
			this.Cancel.Name = "Cancel";
			this.Cancel.Size = new System.Drawing.Size(80, 40);
			this.Cancel.TabIndex = 2;
			this.Cancel.Text = "Cancel";
			this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
			// 
			// btnCopy
			// 
			this.btnCopy.Location = new System.Drawing.Point(152, 206);
			this.btnCopy.Name = "btnCopy";
			this.btnCopy.Size = new System.Drawing.Size(80, 40);
			this.btnCopy.TabIndex = 1;
			this.btnCopy.Text = "Copy";
			this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(38, 32);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(376, 24);
			this.label1.TabIndex = 3;
			this.label1.Text = "!!A table with the same name already exists !!";
			// 
			// radioButton1
			// 
			this.radioButton1.Checked = true;
			this.radioButton1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.radioButton1.Location = new System.Drawing.Point(40, 120);
			this.radioButton1.Name = "radioButton1";
			this.radioButton1.Size = new System.Drawing.Size(392, 32);
			this.radioButton1.TabIndex = 1;
			this.radioButton1.TabStop = true;
			this.radioButton1.Text = "Keep both tables and assign a new table name";
			this.radioButton1.Click += new System.EventHandler(this.radioButton1_Click);
			// 
			// uc_table_exists_dialog
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_table_exists_dialog";
			this.Size = new System.Drawing.Size(472, 392);
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void Cancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.Cancel;
		}

		private void btnCopy_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.OK;
		}

		private void radioButton1_Click(object sender, System.EventArgs e)
		{
			this.txtTable.Enabled=true;
		
		}

		private void rdoOverwrite_Click(object sender, System.EventArgs e)
		{
			this.txtTable.Enabled=false;
		
		}
	
		public string strTable
		{
			get
			{
                return this.txtTable.Text;			 
			}
			set
			{
				this.txtTable.Text = value;
			}
		}
		public string strOkButtonText
		{
			get
			{
				return this.btnCopy.Text;
			}
			set
			{
				this.btnCopy.Text = value;
			}
		}
		public string strDBFileLabel
		{
			set
			{
				this.label2.Text=value;
			}
		}

	
	
	}
}

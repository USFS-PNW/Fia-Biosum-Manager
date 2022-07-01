using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for frmAbout.
	/// </summary>
	public class frmAbout : System.Windows.Forms.Form
	{
		public System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblOwner;
		private System.Windows.Forms.Label lblVersion;
		public System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Label lblDesc;
		private System.Windows.Forms.GroupBox grpboxOwner;
		public System.Windows.Forms.GroupBox grpboxDesc;
        private Label LblReleaseDate;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAbout()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

            this.lblVersion.Text = "Beta Version " + frmMain.g_strAppVer;

            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(assembly.Location);
            DateTime lastModified = fileInfo.LastWriteTime;
            this.LblReleaseDate.Text = "Release date: " + lastModified.ToString("MMMM dd, yyyy");
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmAbout));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lblOwner = new System.Windows.Forms.Label();
            this.lblVersion = new System.Windows.Forms.Label();
            this.lblDesc = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.grpboxOwner = new System.Windows.Forms.GroupBox();
            this.grpboxDesc = new System.Windows.Forms.GroupBox();
            this.LblReleaseDate = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.grpboxOwner.SuspendLayout();
            this.grpboxDesc.SuspendLayout();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(582, 127);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.White;
            this.label1.Font = new System.Drawing.Font("Tahoma", 22F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(4, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(150, 104);
            this.label1.TabIndex = 1;
            this.label1.Text = "FIA Biosum";
            // 
            // lblOwner
            // 
            this.lblOwner.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblOwner.Location = new System.Drawing.Point(29, 18);
            this.lblOwner.Name = "lblOwner";
            this.lblOwner.Size = new System.Drawing.Size(499, 19);
            this.lblOwner.TabIndex = 3;
            this.lblOwner.Text = "USDA Forest Service, PNW FIA Portland Forestry Sciences Lab";
            // 
            // lblVersion
            // 
            this.lblVersion.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVersion.Location = new System.Drawing.Point(186, 156);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(211, 25);
            this.lblVersion.TabIndex = 4;
            this.lblVersion.Text = "Beta Version 5.8.7";
            // 
            // lblDesc
            // 
            this.lblDesc.Location = new System.Drawing.Point(29, 22);
            this.lblDesc.Name = "lblDesc";
            this.lblDesc.Size = new System.Drawing.Size(499, 102);
            this.lblDesc.TabIndex = 6;
            this.lblDesc.Text = resources.GetString("lblDesc.Text");
            this.lblDesc.Click += new System.EventHandler(this.lblDesc_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(471, 429);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(96, 37);
            this.btnOK.TabIndex = 7;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // grpboxOwner
            // 
            this.grpboxOwner.Controls.Add(this.lblOwner);
            this.grpboxOwner.Location = new System.Drawing.Point(10, 203);
            this.grpboxOwner.Name = "grpboxOwner";
            this.grpboxOwner.Size = new System.Drawing.Size(556, 46);
            this.grpboxOwner.TabIndex = 8;
            this.grpboxOwner.TabStop = false;
            // 
            // grpboxDesc
            // 
            this.grpboxDesc.Controls.Add(this.lblDesc);
            this.grpboxDesc.Location = new System.Drawing.Point(10, 277);
            this.grpboxDesc.Name = "grpboxDesc";
            this.grpboxDesc.Size = new System.Drawing.Size(556, 131);
            this.grpboxDesc.TabIndex = 9;
            this.grpboxDesc.TabStop = false;
            this.grpboxDesc.Text = "Description";
            // 
            // LblReleaseDate
            // 
            this.LblReleaseDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblReleaseDate.Location = new System.Drawing.Point(173, 178);
            this.LblReleaseDate.Name = "LblReleaseDate";
            this.LblReleaseDate.Size = new System.Drawing.Size(295, 19);
            this.LblReleaseDate.TabIndex = 10;
            this.LblReleaseDate.Text = "Release date";
            // 
            // frmAbout
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
            this.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.ClientSize = new System.Drawing.Size(582, 478);
            this.Controls.Add(this.LblReleaseDate);
            this.Controls.Add(this.grpboxDesc);
            this.Controls.Add(this.grpboxOwner);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "frmAbout";
            this.Text = "About FIA Biosum";
            this.TopMost = true;
            this.Resize += new System.EventHandler(this.frmAbout_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.grpboxOwner.ResumeLayout(false);
            this.grpboxDesc.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void frmAbout_Resize(object sender, System.EventArgs e)
		{
        
		}
		public void resize_frmAbout()
		{
			
			this.lblVersion.Left = (int)(this.Width * .50) - (int)(this.lblVersion.Width * .50);
			this.grpboxOwner.Left = (int)(this.Width * .50) - (int)(this.grpboxOwner.Width * .50);
			this.grpboxDesc.Left = (int)(this.Width * .50) - (int)(this.grpboxDesc.Width * .50);

		}

        private void lblDesc_Click(object sender, EventArgs e)
        {

        }
	}
}

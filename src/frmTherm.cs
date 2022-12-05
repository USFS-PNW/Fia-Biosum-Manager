using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for frmTherm.
	/// </summary>
	public class frmTherm : System.Windows.Forms.Form
	{
		public System.Windows.Forms.ProgressBar progressBar1;
		public System.Windows.Forms.Button btnCancel;
		private bool m_bAbortProcess=false;
		
		public System.Windows.Forms.Label lblMsg;
		public System.Windows.Forms.ProgressBar progressBar2;
		private FIA_Biosum_Manager.frmDialog m_frmDialog;
		private string m_strClient;
		public System.Windows.Forms.Label lblMsg2;
        public Button btnMinimize;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTherm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmTherm(FIA_Biosum_Manager.frmDialog p_frmDialog,string p_strClient)
		{
			InitializeComponent();
			m_frmDialog = p_frmDialog;
			this.m_strClient = p_strClient;
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);


		}
		public frmTherm(FIA_Biosum_Manager.frmDialog p_frmDialog,string p_strClient,string p_strTitle,string p_strThermCount)
		{
			InitializeComponent();
			m_frmDialog = p_frmDialog;
			this.m_strClient = p_strClient;
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			this.StartTherm(p_strTitle,p_strThermCount);
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

		public bool AbortProcess
		{
			get
			{
				return this.m_bAbortProcess;
			}
			set
			{
				this.m_bAbortProcess = value;
				
			}
		}
		public void Increment(int intValue)
		{
			try
			{
				this.progressBar1.Value = intValue;
			}
			catch 
			{
			}

		}
		private void StartTherm(string strTitle, string p_strNumberOfTherms)
		{
			
			this.Text = strTitle;
			this.Visible=false;
			this.btnCancel.Visible=true;
			this.lblMsg.Visible=true;
			this.progressBar1.Minimum=0;
			this.progressBar1.Visible=true;
			this.progressBar1.Maximum = 10;
			this.lblMsg.Text="";

			if (p_strNumberOfTherms=="2")
			{
				this.progressBar2.Size = this.progressBar1.Size;
				this.progressBar2.Left = this.progressBar1.Left;
				this.progressBar2.Top = Convert.ToInt32(this.progressBar1.Top + (this.progressBar1.Height * 3));
				this.lblMsg2.Top = this.progressBar2.Top + this.progressBar2.Height + 5;
				this.lblMsg2.Show();
				this.Height = this.lblMsg2.Top + this.lblMsg2.Height + this.btnCancel.Height + 50;
				this.btnCancel.Top = this.ClientSize.Height - this.btnCancel.Height - 5;
				this.progressBar2.Visible=true;
				this.lblMsg2.Text="";
			}
			this.AbortProcess = false;
			this.Refresh();
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Visible=true;
			
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblMsg = new System.Windows.Forms.Label();
            this.progressBar2 = new System.Windows.Forms.ProgressBar();
            this.lblMsg2 = new System.Windows.Forms.Label();
            this.btnMinimize = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // progressBar1
            // 
            this.progressBar1.ForeColor = System.Drawing.Color.ForestGreen;
            this.progressBar1.Location = new System.Drawing.Point(8, 8);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(472, 24);
            this.progressBar1.TabIndex = 0;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(214, 59);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(72, 24);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            // 
            // lblMsg
            // 
            this.lblMsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMsg.Location = new System.Drawing.Point(8, 40);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(472, 16);
            this.lblMsg.TabIndex = 2;
            this.lblMsg.Text = "lblMsg";
            this.lblMsg.Visible = false;
            // 
            // progressBar2
            // 
            this.progressBar2.ForeColor = System.Drawing.Color.ForestGreen;
            this.progressBar2.Location = new System.Drawing.Point(367, 56);
            this.progressBar2.Name = "progressBar2";
            this.progressBar2.Size = new System.Drawing.Size(47, 24);
            this.progressBar2.TabIndex = 3;
            this.progressBar2.Visible = false;
            // 
            // lblMsg2
            // 
            this.lblMsg2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMsg2.Location = new System.Drawing.Point(9, 64);
            this.lblMsg2.Name = "lblMsg2";
            this.lblMsg2.Size = new System.Drawing.Size(183, 16);
            this.lblMsg2.TabIndex = 4;
            this.lblMsg2.Text = "lblMsg2";
            this.lblMsg2.Visible = false;
            // 
            // btnMinimize
            // 
            this.btnMinimize.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMinimize.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMinimize.Location = new System.Drawing.Point(420, 61);
            this.btnMinimize.Name = "btnMinimize";
            this.btnMinimize.Size = new System.Drawing.Size(62, 21);
            this.btnMinimize.TabIndex = 5;
            this.btnMinimize.Text = "Minimize";
            this.btnMinimize.Click += new System.EventHandler(this.btnMinimize_Click);
            // 
            // frmTherm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(490, 86);
            this.ControlBox = false;
            this.Controls.Add(this.btnMinimize);
            this.Controls.Add(this.lblMsg2);
            this.Controls.Add(this.progressBar2);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.progressBar1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmTherm";
            this.ResumeLayout(false);

		}
		#endregion

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			//MessageBox.Show("cancel button pressed");
			switch (this.m_strClient.Trim().ToUpper())
			{
			    case "ADD PLOT DATA":
				   this.m_frmDialog.uc_plot_input1.StopThread();
                   break;
				case "APPEND FVS INPUT DATA":
				   this.m_frmDialog.uc_fvs_input1.StopThread();
				   break;
				case "FVS OUT DATA":
				   this.m_frmDialog.uc_fvs_output1.StopThread();
					break;
                case "ADD TEXT FILE POP TABLE DATA":
					this.m_frmDialog.uc_plot_input1.StopThread();
					break;
				case "ADD TEXT FILE PLOT,COND,SITE TREE, & TREE TABLE DATA":
					this.m_frmDialog.uc_plot_input1.StopThread();
					break;
				case "ADD MS ACCESS POP TABLE DATA":
					this.m_frmDialog.uc_plot_input1.StopThread();
					break;
				case "ADD MS ACCESS PLOT,COND,SITE TREE, & TREE TABLE DATA":
					this.m_frmDialog.uc_plot_input1.StopThread();
                    break;
                case "FVS IN FOREST TABLE":
                    this.m_frmDialog.uc_fvs_output1.StopThread();
                    break;

                default:
					break;
			}
			
		}

        private void btnMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
	}
}

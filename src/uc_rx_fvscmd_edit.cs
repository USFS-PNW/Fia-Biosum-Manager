using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_rx_fvscmd_edit.
	/// </summary>
	public class uc_rx_fvscmd_edit : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.GroupBox grpboxFVSCmd;
		private System.Windows.Forms.ComboBox cmbFVSCmd;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtFVSCmdVariantList;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.GroupBox grpboxP1;
		private System.Windows.Forms.TextBox txtP1Desc;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtP1;
		private System.Windows.Forms.TextBox txtP1Valid;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.GroupBox grpboxP2;
		private System.Windows.Forms.TextBox txtP2Valid;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.TextBox txtP2;
		private System.Windows.Forms.TextBox txtP2Desc;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.GroupBox grpboxP3;
		private System.Windows.Forms.TextBox txtP3Valid;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.TextBox txtP3;
		private System.Windows.Forms.TextBox txtP3Desc;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label lblP3Defa;
		private System.Windows.Forms.Label lblP2Defa;
		private System.Windows.Forms.Label lblP1Defa;
		private System.Windows.Forms.GroupBox grpboxP4;
		private System.Windows.Forms.Label lblP4Defa;
		private System.Windows.Forms.TextBox txtP4Valid;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.TextBox txtP4;
		private System.Windows.Forms.TextBox txtP4Desc;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.GroupBox grpboxP5;
		private System.Windows.Forms.Label lblP5Defa;
		private System.Windows.Forms.TextBox txtP5Valid;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.TextBox txtP5;
		private System.Windows.Forms.TextBox txtP5Desc;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.GroupBox grpboxP6;
		private System.Windows.Forms.Label lblP6Defa;
		private System.Windows.Forms.TextBox txtP6Valid;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.TextBox txtP6Desc;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.GroupBox grpboxP7;
		private System.Windows.Forms.Label lblP7Defa;
		private System.Windows.Forms.TextBox txtP7Valid;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.TextBox txtP7;
		private System.Windows.Forms.TextBox txtP7Desc;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.GroupBox grpboxOther;
		private System.Windows.Forms.TextBox txtOther;

		public FIA_Biosum_Manager.ResizeFormUsingVisibleScrollBars m_oResizeForm = new ResizeFormUsingVisibleScrollBars();

	    private FIA_Biosum_Manager.frmRxItemFvsCmdItem _frmFvsCmdItem;
		
		private System.Windows.Forms.TextBox txtFVSCmdDesc;
		private Queries m_oQueries = new Queries();
		private ado_data_access m_oAdo = new ado_data_access();

		private System.Windows.Forms.TextBox txtP6;
		private System.Windows.Forms.ComboBox cmbFilter;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.TextBox txtOtherDesc;
		private System.Windows.Forms.Label label19;
		private bool _bRxPackageEdit=false;

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_rx_fvscmd_edit()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			m_oResizeForm.ScrollBarParentControl=panel1;
			m_oResizeForm.ResizeHeight=false;
			m_oResizeForm.MaximumWidth = 1000;
			//m_oResizeForm.MaximumWidth=770;
			//m_oResizeForm.MaximumHeight=630;

			// TODO: Add any initialization after the InitializeComponent call

		}
	    ~uc_rx_fvscmd_edit()
		{
			m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);

		}
		
		private void LoadFvsCommandsComboBox()
		{
			m_oQueries.m_oFvs.LoadDatasource=true;
			m_oQueries.LoadDatasources(true);



			this.cmbFVSCmd.Items.Clear();

			
			m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile,"",""));
           

			m_oAdo.m_OleDbDataReader.Close();

		}
		protected void LoadFvsCommandProperties(string p_strFvsCmd)
		{
			string strCurrFvsCmd="";			
			if (p_strFvsCmd.Trim().ToUpper()==strCurrFvsCmd.Trim().ToUpper())
			{

			}
			else
			{
				this.txtP1.Text = "";
				this.txtP2.Text = "";
				this.txtP3.Text = "";
				this.txtP4.Text = "";
				this.txtP5.Text = "";
				this.txtP6.Text = "";
				this.txtP7.Text = "";
				this.txtOther.Text = "";

			}

			this.cmbFVSCmd.Text = p_strFvsCmd;
			this.txtFVSCmdDesc.Text="";
			this.txtP1Desc.Text="";
			this.txtP1Valid.Text ="";
			this.lblP1Defa.Text="";
			this.txtP2Desc.Text="";
			this.txtP2Valid.Text="";
			this.lblP2Defa.Text="";
			this.txtP3Desc.Text="";
			this.txtP3Valid.Text="";
			this.lblP3Defa.Text="";
			this.txtP4Desc.Text="";
			this.txtP4Valid.Text ="";
			this.lblP4Defa.Text="";
			this.txtP5Desc.Text="";
			this.txtP5Valid.Text ="";
			this.lblP5Defa.Text="";
			this.txtP6Desc.Text="";
			this.txtP6Valid.Text="";
			this.lblP6Defa.Text="";
			this.txtP7Desc.Text="";
			this.txtP7Valid.Text="";
			this.lblP7Defa.Text="";
			this.txtFVSCmdVariantList.Text = "";
			this.txtOtherDesc.Text = "";

			this.grpboxP1.Enabled=true;
			this.grpboxP2.Enabled=true;
			this.grpboxP3.Enabled=true;
			this.grpboxP4.Enabled=true;
			this.grpboxP5.Enabled=true;
			this.grpboxP6.Enabled=true;
			this.grpboxP7.Enabled=true;
			this.grpboxOther.Enabled=true;


			this.grpboxP1.Text = "Parameter 1";
			this.grpboxP2.Text = "Parameter 2";
			this.grpboxP3.Text = "Parameter 3";
			this.grpboxP4.Text = "Parameter 4";
			this.grpboxP5.Text = "Parameter 5";
			this.grpboxP6.Text = "Parameter 6";
			this.grpboxP7.Text = "Parameter 7";
			this.grpboxOther.Text = "Other";
			if (p_strFvsCmd.Trim().Length > 0)
			{
				m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL);
				if (m_oAdo.m_OleDbDataReader.HasRows)
				{
					while (m_oAdo.m_OleDbDataReader.Read())
					{
						if (m_oAdo.m_OleDbDataReader["desc"] != System.DBNull.Value)
						{
							this.txtFVSCmdDesc.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["desc"]);
						}
						if (m_oAdo.m_OleDbDataReader["fvs_variant_list"] != System.DBNull.Value)
						{
							this.txtFVSCmdVariantList.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["fvs_variant_list"]);
						}

						if (m_oAdo.m_OleDbDataReader["p1_label"] != System.DBNull.Value)
						{
							this.grpboxP1.Text = "Parameter 1: " + m_oAdo.m_OleDbDataReader["p1_label"].ToString().Trim();
						}

						if (this.grpboxP1.Text.Trim().ToUpper() == "PARAMETER 1: SIMULATION CYCLE")
							if (m_oAdo.m_OleDbDataReader["p1_desc"] != System.DBNull.Value)
							{
								this.txtP1Desc.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p1_desc"]);
							}
						if (m_oAdo.m_OleDbDataReader["p1_default"] != System.DBNull.Value)
						{
							this.lblP1Defa.Text = "Default = " + Convert.ToString(m_oAdo.m_OleDbDataReader["p1_default"]);
						}
						if (m_oAdo.m_OleDbDataReader["p1_validvalues"] != System.DBNull.Value)
						{
							this.txtP1Valid.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p1_validvalues"]);
						}
						if (this.grpboxP1.Text.Trim().ToUpper() == "PARAMETER 1: SIMULATION CYCLE")
						{
							this.grpboxP1.Text = this.grpboxP1.Text.Trim() + " (Defined in treatment package)";
							this.grpboxP1.Enabled=false;
						}
					
						if (m_oAdo.m_OleDbDataReader["p2_label"] != System.DBNull.Value)
						{
							this.grpboxP2.Text = "Parameter 2: " + m_oAdo.m_OleDbDataReader["p2_label"].ToString().Trim();
						}
						if (m_oAdo.m_OleDbDataReader["p2_desc"] != System.DBNull.Value)
						{
							this.txtP2Desc.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p2_desc"]);
						}
						if (m_oAdo.m_OleDbDataReader["p2_default"] != System.DBNull.Value)
						{
							this.lblP2Defa.Text = "Default = " + Convert.ToString(m_oAdo.m_OleDbDataReader["p2_default"]);
						}
						if (m_oAdo.m_OleDbDataReader["p2_validvalues"] != System.DBNull.Value)
						{
							this.txtP2Valid.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p2_validvalues"]);
						}

						if (m_oAdo.m_OleDbDataReader["p3_label"] != System.DBNull.Value)
						{
							this.grpboxP3.Text = "Parameter 3: " + m_oAdo.m_OleDbDataReader["p3_label"].ToString().Trim();
						}
						if (m_oAdo.m_OleDbDataReader["p3_desc"] != System.DBNull.Value)
						{
							this.txtP3Desc.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p3_desc"]);
						}
						if (m_oAdo.m_OleDbDataReader["p3_default"] != System.DBNull.Value)
						{
							this.lblP3Defa.Text = "Default = " + Convert.ToString(m_oAdo.m_OleDbDataReader["p3_default"]);
						}
						if (m_oAdo.m_OleDbDataReader["p3_validvalues"] != System.DBNull.Value)
						{
							this.txtP3Valid.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p3_validvalues"]);
						}

						if (m_oAdo.m_OleDbDataReader["p4_label"] != System.DBNull.Value)
						{
							this.grpboxP4.Text = "Parameter 4: " + m_oAdo.m_OleDbDataReader["p4_label"].ToString().Trim();
						}
						if (m_oAdo.m_OleDbDataReader["p4_desc"] != System.DBNull.Value)
						{
							this.txtP4Desc.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p4_desc"]);
						}
						if (m_oAdo.m_OleDbDataReader["p4_default"] != System.DBNull.Value)
						{
							this.lblP4Defa.Text = "Default = " + Convert.ToString(m_oAdo.m_OleDbDataReader["p4_default"]);
						}
						if (m_oAdo.m_OleDbDataReader["p4_validvalues"] != System.DBNull.Value)
						{
							this.txtP4Valid.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p4_validvalues"]);
						}

						if (m_oAdo.m_OleDbDataReader["p5_label"] != System.DBNull.Value)
						{
							this.grpboxP5.Text = "Parameter 5: " + m_oAdo.m_OleDbDataReader["p5_label"].ToString().Trim();
						}
						if (m_oAdo.m_OleDbDataReader["p5_desc"] != System.DBNull.Value)
						{
							this.txtP5Desc.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p5_desc"]);
						}
						if (m_oAdo.m_OleDbDataReader["p5_default"] != System.DBNull.Value)
						{
							this.lblP5Defa.Text = "Default = " + Convert.ToString(m_oAdo.m_OleDbDataReader["p5_default"]);
						}
						if (m_oAdo.m_OleDbDataReader["p5_validvalues"] != System.DBNull.Value)
						{
							this.txtP5Valid.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p5_validvalues"]);
						}

						if (m_oAdo.m_OleDbDataReader["p6_label"] != System.DBNull.Value)
						{
							this.grpboxP6.Text = "Parameter 6: " + m_oAdo.m_OleDbDataReader["p6_label"].ToString().Trim();
						}
						if (m_oAdo.m_OleDbDataReader["p6_desc"] != System.DBNull.Value)
						{
							this.txtP6Desc.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p6_desc"]);
						}
						if (m_oAdo.m_OleDbDataReader["p6_default"] != System.DBNull.Value)
						{
							this.lblP6Defa.Text = "Default = " + Convert.ToString(m_oAdo.m_OleDbDataReader["p6_default"]);
						}
						if (m_oAdo.m_OleDbDataReader["p6_validvalues"] != System.DBNull.Value)
						{
							this.txtP6Valid.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p6_validvalues"]);
						}

						if (m_oAdo.m_OleDbDataReader["p7_label"] != System.DBNull.Value)
						{
							this.grpboxP7.Text = "Parameter 7: " + m_oAdo.m_OleDbDataReader["p7_label"].ToString().Trim();
						}
						if (m_oAdo.m_OleDbDataReader["p7_desc"] != System.DBNull.Value)
						{
							this.txtP7Desc.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p7_desc"]);
						}
						if (m_oAdo.m_OleDbDataReader["p7_default"] != System.DBNull.Value)
						{
							this.lblP7Defa.Text = "Default = " + Convert.ToString(m_oAdo.m_OleDbDataReader["p7_default"]);
						}
						if (m_oAdo.m_OleDbDataReader["p7_validvalues"] != System.DBNull.Value)
						{
							this.txtP7Valid.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["p7_validvalues"]);
						}

						if (m_oAdo.m_OleDbDataReader["other_label"] != System.DBNull.Value)
						{
							this.grpboxOther.Text = "Other: " + m_oAdo.m_OleDbDataReader["other_label"].ToString().Trim();
						}
						if (m_oAdo.m_OleDbDataReader["other_desc"] != System.DBNull.Value)
						{
							txtOtherDesc.Text = Convert.ToString(m_oAdo.m_OleDbDataReader["other_desc"]);
						}
					

				


					}
				}
				this.m_oAdo.m_OleDbDataReader.Close();
			}

			if (this.grpboxP1.Text.Trim().ToUpper() == "PARAMETER 1: SIMULATION CYCLE (DEFINED IN TREATMENT PACKAGE)")
			{
				this.txtP1.Text = "NA";
			}

			

			if (this.grpboxP1.Text.IndexOf(":",0) < 0) this.grpboxP1.Enabled=false;
			if (this.grpboxP2.Text.IndexOf(":",0) < 0) this.grpboxP2.Enabled=false;
			if (this.grpboxP3.Text.IndexOf(":",0) < 0) this.grpboxP3.Enabled=false;
			if (this.grpboxP4.Text.IndexOf(":",0) < 0) this.grpboxP4.Enabled=false;
			if (this.grpboxP5.Text.IndexOf(":",0) < 0) this.grpboxP5.Enabled=false;
			if (this.grpboxP6.Text.IndexOf(":",0) < 0) this.grpboxP6.Enabled=false;
			if (this.grpboxP7.Text.IndexOf(":",0) < 0) this.grpboxP7.Enabled=false;
			if (this.grpboxOther.Text.IndexOf(":",0) < 0) this.grpboxOther.Enabled=false;

			if (this.txtP1.Text != "NA")
			{
				if (this.grpboxP1.Enabled && 
					this.txtP1.Text.Trim().Length == 0  &&
					this.lblP1Defa.Text.Trim().Length > 0)
				{
					this.txtP1.Text = this.lblP1Defa.Text.Substring(10,this.lblP1Defa.Text.Length-10);
				}
			}
			if (this.grpboxP2.Enabled && 
				this.txtP2.Text.Trim().Length == 0  &&
				this.lblP2Defa.Text.Trim().Length > 0)
			{
				this.txtP2.Text = this.lblP2Defa.Text.Substring(10,this.lblP2Defa.Text.Length-10);
			}
			if (this.grpboxP3.Enabled && 
				this.txtP3.Text.Trim().Length == 0  &&
				this.lblP3Defa.Text.Trim().Length > 0)
			{
				this.txtP3.Text = this.lblP3Defa.Text.Substring(10,this.lblP3Defa.Text.Length-10);
			}
			if (this.grpboxP4.Enabled && 
				this.txtP4.Text.Trim().Length == 0  &&
				this.lblP4Defa.Text.Trim().Length > 0)
			{
				this.txtP4.Text = this.lblP4Defa.Text.Substring(10,this.lblP4Defa.Text.Length-10);
			}
			if (this.grpboxP5.Enabled && 
				this.txtP5.Text.Trim().Length == 0  &&
				this.lblP5Defa.Text.Trim().Length > 0)
			{
				this.txtP5.Text = this.lblP5Defa.Text.Substring(10,this.lblP5Defa.Text.Length-10);
			}
			if (this.grpboxP6.Enabled && 
				this.txtP6.Text.Trim().Length == 0  &&
				this.lblP6Defa.Text.Trim().Length > 0)
			{
				this.txtP6.Text = this.lblP6Defa.Text.Substring(10,this.lblP6Defa.Text.Length-10);
			}
			
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
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.panel1 = new System.Windows.Forms.Panel();
			this.grpboxOther = new System.Windows.Forms.GroupBox();
			this.txtOtherDesc = new System.Windows.Forms.TextBox();
			this.label19 = new System.Windows.Forms.Label();
			this.txtOther = new System.Windows.Forms.TextBox();
			this.grpboxP6 = new System.Windows.Forms.GroupBox();
			this.lblP6Defa = new System.Windows.Forms.Label();
			this.txtP6Valid = new System.Windows.Forms.TextBox();
			this.label14 = new System.Windows.Forms.Label();
			this.txtP6 = new System.Windows.Forms.TextBox();
			this.txtP6Desc = new System.Windows.Forms.TextBox();
			this.label15 = new System.Windows.Forms.Label();
			this.grpboxP5 = new System.Windows.Forms.GroupBox();
			this.lblP5Defa = new System.Windows.Forms.Label();
			this.txtP5Valid = new System.Windows.Forms.TextBox();
			this.label12 = new System.Windows.Forms.Label();
			this.txtP5 = new System.Windows.Forms.TextBox();
			this.txtP5Desc = new System.Windows.Forms.TextBox();
			this.label13 = new System.Windows.Forms.Label();
			this.grpboxP4 = new System.Windows.Forms.GroupBox();
			this.lblP4Defa = new System.Windows.Forms.Label();
			this.txtP4Valid = new System.Windows.Forms.TextBox();
			this.label10 = new System.Windows.Forms.Label();
			this.txtP4 = new System.Windows.Forms.TextBox();
			this.txtP4Desc = new System.Windows.Forms.TextBox();
			this.label11 = new System.Windows.Forms.Label();
			this.grpboxP3 = new System.Windows.Forms.GroupBox();
			this.lblP3Defa = new System.Windows.Forms.Label();
			this.txtP3Valid = new System.Windows.Forms.TextBox();
			this.label7 = new System.Windows.Forms.Label();
			this.txtP3 = new System.Windows.Forms.TextBox();
			this.txtP3Desc = new System.Windows.Forms.TextBox();
			this.label8 = new System.Windows.Forms.Label();
			this.grpboxP2 = new System.Windows.Forms.GroupBox();
			this.lblP2Defa = new System.Windows.Forms.Label();
			this.txtP2Valid = new System.Windows.Forms.TextBox();
			this.label5 = new System.Windows.Forms.Label();
			this.txtP2 = new System.Windows.Forms.TextBox();
			this.txtP2Desc = new System.Windows.Forms.TextBox();
			this.label6 = new System.Windows.Forms.Label();
			this.grpboxP1 = new System.Windows.Forms.GroupBox();
			this.lblP1Defa = new System.Windows.Forms.Label();
			this.txtP1Valid = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.txtP1 = new System.Windows.Forms.TextBox();
			this.txtP1Desc = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.grpboxFVSCmd = new System.Windows.Forms.GroupBox();
			this.label18 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbFilter = new System.Windows.Forms.ComboBox();
			this.label2 = new System.Windows.Forms.Label();
			this.txtFVSCmdVariantList = new System.Windows.Forms.TextBox();
			this.txtFVSCmdDesc = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbFVSCmd = new System.Windows.Forms.ComboBox();
			this.grpboxP7 = new System.Windows.Forms.GroupBox();
			this.lblP7Defa = new System.Windows.Forms.Label();
			this.txtP7Valid = new System.Windows.Forms.TextBox();
			this.label16 = new System.Windows.Forms.Label();
			this.txtP7 = new System.Windows.Forms.TextBox();
			this.txtP7Desc = new System.Windows.Forms.TextBox();
			this.label17 = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			this.panel1.SuspendLayout();
			this.grpboxOther.SuspendLayout();
			this.grpboxP6.SuspendLayout();
			this.grpboxP5.SuspendLayout();
			this.grpboxP4.SuspendLayout();
			this.grpboxP3.SuspendLayout();
			this.grpboxP2.SuspendLayout();
			this.grpboxP1.SuspendLayout();
			this.grpboxFVSCmd.SuspendLayout();
			this.grpboxP7.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.panel1);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(872, 2184);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// panel1
			// 
			this.panel1.AutoScroll = true;
			this.panel1.Controls.Add(this.grpboxOther);
			this.panel1.Controls.Add(this.grpboxP6);
			this.panel1.Controls.Add(this.grpboxP5);
			this.panel1.Controls.Add(this.grpboxP4);
			this.panel1.Controls.Add(this.grpboxP3);
			this.panel1.Controls.Add(this.grpboxP2);
			this.panel1.Controls.Add(this.grpboxP1);
			this.panel1.Controls.Add(this.grpboxFVSCmd);
			this.panel1.Controls.Add(this.grpboxP7);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.panel1.Location = new System.Drawing.Point(3, 18);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(866, 2163);
			this.panel1.TabIndex = 0;
			this.panel1.Resize += new System.EventHandler(this.panel1_Resize);
			// 
			// grpboxOther
			// 
			this.grpboxOther.Controls.Add(this.txtOtherDesc);
			this.grpboxOther.Controls.Add(this.label19);
			this.grpboxOther.Controls.Add(this.txtOther);
			this.grpboxOther.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpboxOther.Location = new System.Drawing.Point(8, 1880);
			this.grpboxOther.Name = "grpboxOther";
			this.grpboxOther.Size = new System.Drawing.Size(848, 272);
			this.grpboxOther.TabIndex = 8;
			this.grpboxOther.TabStop = false;
			this.grpboxOther.Text = "Other";
			// 
			// txtOtherDesc
			// 
			this.txtOtherDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtOtherDesc.Location = new System.Drawing.Point(16, 168);
			this.txtOtherDesc.Multiline = true;
			this.txtOtherDesc.Name = "txtOtherDesc";
			this.txtOtherDesc.Size = new System.Drawing.Size(808, 88);
			this.txtOtherDesc.TabIndex = 9;
			this.txtOtherDesc.Text = "";
			// 
			// label19
			// 
			this.label19.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label19.Location = new System.Drawing.Point(16, 144);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(136, 16);
			this.label19.TabIndex = 8;
			this.label19.Text = "Description";
			// 
			// txtOther
			// 
			this.txtOther.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtOther.Location = new System.Drawing.Point(16, 24);
			this.txtOther.Multiline = true;
			this.txtOther.Name = "txtOther";
			this.txtOther.Size = new System.Drawing.Size(816, 112);
			this.txtOther.TabIndex = 7;
			this.txtOther.Text = "";
			// 
			// grpboxP6
			// 
			this.grpboxP6.Controls.Add(this.lblP6Defa);
			this.grpboxP6.Controls.Add(this.txtP6Valid);
			this.grpboxP6.Controls.Add(this.label14);
			this.grpboxP6.Controls.Add(this.txtP6);
			this.grpboxP6.Controls.Add(this.txtP6Desc);
			this.grpboxP6.Controls.Add(this.label15);
			this.grpboxP6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpboxP6.Location = new System.Drawing.Point(8, 1416);
			this.grpboxP6.Name = "grpboxP6";
			this.grpboxP6.Size = new System.Drawing.Size(848, 224);
			this.grpboxP6.TabIndex = 6;
			this.grpboxP6.TabStop = false;
			this.grpboxP6.Text = "Parameter 6";
			// 
			// lblP6Defa
			// 
			this.lblP6Defa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblP6Defa.Location = new System.Drawing.Point(120, 16);
			this.lblP6Defa.Name = "lblP6Defa";
			this.lblP6Defa.Size = new System.Drawing.Size(680, 24);
			this.lblP6Defa.TabIndex = 10;
			this.lblP6Defa.Text = "Default Value";
			// 
			// txtP6Valid
			// 
			this.txtP6Valid.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP6Valid.Location = new System.Drawing.Point(16, 160);
			this.txtP6Valid.Multiline = true;
			this.txtP6Valid.Name = "txtP6Valid";
			this.txtP6Valid.Size = new System.Drawing.Size(808, 56);
			this.txtP6Valid.TabIndex = 9;
			this.txtP6Valid.Text = "";
			// 
			// label14
			// 
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(16, 136);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(136, 16);
			this.label14.TabIndex = 8;
			this.label14.Text = "Valid Values";
			// 
			// txtP6
			// 
			this.txtP6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP6.Location = new System.Drawing.Point(16, 16);
			this.txtP6.Name = "txtP6";
			this.txtP6.Size = new System.Drawing.Size(88, 22);
			this.txtP6.TabIndex = 7;
			this.txtP6.Text = "";
			// 
			// txtP6Desc
			// 
			this.txtP6Desc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP6Desc.Location = new System.Drawing.Point(16, 72);
			this.txtP6Desc.Multiline = true;
			this.txtP6Desc.Name = "txtP6Desc";
			this.txtP6Desc.Size = new System.Drawing.Size(808, 56);
			this.txtP6Desc.TabIndex = 6;
			this.txtP6Desc.Text = "";
			// 
			// label15
			// 
			this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.Location = new System.Drawing.Point(16, 48);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(136, 16);
			this.label15.TabIndex = 5;
			this.label15.Text = "Description";
			// 
			// grpboxP5
			// 
			this.grpboxP5.Controls.Add(this.lblP5Defa);
			this.grpboxP5.Controls.Add(this.txtP5Valid);
			this.grpboxP5.Controls.Add(this.label12);
			this.grpboxP5.Controls.Add(this.txtP5);
			this.grpboxP5.Controls.Add(this.txtP5Desc);
			this.grpboxP5.Controls.Add(this.label13);
			this.grpboxP5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpboxP5.Location = new System.Drawing.Point(8, 1184);
			this.grpboxP5.Name = "grpboxP5";
			this.grpboxP5.Size = new System.Drawing.Size(848, 224);
			this.grpboxP5.TabIndex = 5;
			this.grpboxP5.TabStop = false;
			this.grpboxP5.Text = "Parameter 5";
			// 
			// lblP5Defa
			// 
			this.lblP5Defa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblP5Defa.Location = new System.Drawing.Point(120, 16);
			this.lblP5Defa.Name = "lblP5Defa";
			this.lblP5Defa.Size = new System.Drawing.Size(680, 24);
			this.lblP5Defa.TabIndex = 10;
			this.lblP5Defa.Text = "Default Value";
			// 
			// txtP5Valid
			// 
			this.txtP5Valid.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP5Valid.Location = new System.Drawing.Point(16, 160);
			this.txtP5Valid.Multiline = true;
			this.txtP5Valid.Name = "txtP5Valid";
			this.txtP5Valid.Size = new System.Drawing.Size(808, 56);
			this.txtP5Valid.TabIndex = 9;
			this.txtP5Valid.Text = "";
			// 
			// label12
			// 
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(16, 136);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(136, 16);
			this.label12.TabIndex = 8;
			this.label12.Text = "Valid Values";
			// 
			// txtP5
			// 
			this.txtP5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP5.Location = new System.Drawing.Point(16, 16);
			this.txtP5.Name = "txtP5";
			this.txtP5.Size = new System.Drawing.Size(88, 22);
			this.txtP5.TabIndex = 7;
			this.txtP5.Text = "";
			// 
			// txtP5Desc
			// 
			this.txtP5Desc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP5Desc.Location = new System.Drawing.Point(16, 72);
			this.txtP5Desc.Multiline = true;
			this.txtP5Desc.Name = "txtP5Desc";
			this.txtP5Desc.Size = new System.Drawing.Size(808, 56);
			this.txtP5Desc.TabIndex = 6;
			this.txtP5Desc.Text = "";
			// 
			// label13
			// 
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.Location = new System.Drawing.Point(16, 48);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(136, 16);
			this.label13.TabIndex = 5;
			this.label13.Text = "Description";
			// 
			// grpboxP4
			// 
			this.grpboxP4.Controls.Add(this.lblP4Defa);
			this.grpboxP4.Controls.Add(this.txtP4Valid);
			this.grpboxP4.Controls.Add(this.label10);
			this.grpboxP4.Controls.Add(this.txtP4);
			this.grpboxP4.Controls.Add(this.txtP4Desc);
			this.grpboxP4.Controls.Add(this.label11);
			this.grpboxP4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpboxP4.Location = new System.Drawing.Point(8, 952);
			this.grpboxP4.Name = "grpboxP4";
			this.grpboxP4.Size = new System.Drawing.Size(848, 224);
			this.grpboxP4.TabIndex = 4;
			this.grpboxP4.TabStop = false;
			this.grpboxP4.Text = "Parameter 4";
			// 
			// lblP4Defa
			// 
			this.lblP4Defa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblP4Defa.Location = new System.Drawing.Point(120, 16);
			this.lblP4Defa.Name = "lblP4Defa";
			this.lblP4Defa.Size = new System.Drawing.Size(680, 24);
			this.lblP4Defa.TabIndex = 10;
			this.lblP4Defa.Text = "Default Value";
			// 
			// txtP4Valid
			// 
			this.txtP4Valid.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP4Valid.Location = new System.Drawing.Point(16, 160);
			this.txtP4Valid.Multiline = true;
			this.txtP4Valid.Name = "txtP4Valid";
			this.txtP4Valid.Size = new System.Drawing.Size(808, 56);
			this.txtP4Valid.TabIndex = 9;
			this.txtP4Valid.Text = "";
			// 
			// label10
			// 
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(16, 136);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(136, 16);
			this.label10.TabIndex = 8;
			this.label10.Text = "Valid Values";
			// 
			// txtP4
			// 
			this.txtP4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP4.Location = new System.Drawing.Point(16, 16);
			this.txtP4.Name = "txtP4";
			this.txtP4.Size = new System.Drawing.Size(88, 22);
			this.txtP4.TabIndex = 7;
			this.txtP4.Text = "";
			// 
			// txtP4Desc
			// 
			this.txtP4Desc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP4Desc.Location = new System.Drawing.Point(16, 72);
			this.txtP4Desc.Multiline = true;
			this.txtP4Desc.Name = "txtP4Desc";
			this.txtP4Desc.Size = new System.Drawing.Size(808, 56);
			this.txtP4Desc.TabIndex = 6;
			this.txtP4Desc.Text = "";
			// 
			// label11
			// 
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.Location = new System.Drawing.Point(16, 48);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(136, 16);
			this.label11.TabIndex = 5;
			this.label11.Text = "Description";
			// 
			// grpboxP3
			// 
			this.grpboxP3.Controls.Add(this.lblP3Defa);
			this.grpboxP3.Controls.Add(this.txtP3Valid);
			this.grpboxP3.Controls.Add(this.label7);
			this.grpboxP3.Controls.Add(this.txtP3);
			this.grpboxP3.Controls.Add(this.txtP3Desc);
			this.grpboxP3.Controls.Add(this.label8);
			this.grpboxP3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpboxP3.Location = new System.Drawing.Point(8, 720);
			this.grpboxP3.Name = "grpboxP3";
			this.grpboxP3.Size = new System.Drawing.Size(848, 224);
			this.grpboxP3.TabIndex = 3;
			this.grpboxP3.TabStop = false;
			this.grpboxP3.Text = "Parameter 3";
			// 
			// lblP3Defa
			// 
			this.lblP3Defa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblP3Defa.Location = new System.Drawing.Point(120, 16);
			this.lblP3Defa.Name = "lblP3Defa";
			this.lblP3Defa.Size = new System.Drawing.Size(680, 24);
			this.lblP3Defa.TabIndex = 10;
			this.lblP3Defa.Text = "Default Value";
			// 
			// txtP3Valid
			// 
			this.txtP3Valid.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP3Valid.Location = new System.Drawing.Point(16, 160);
			this.txtP3Valid.Multiline = true;
			this.txtP3Valid.Name = "txtP3Valid";
			this.txtP3Valid.Size = new System.Drawing.Size(808, 56);
			this.txtP3Valid.TabIndex = 9;
			this.txtP3Valid.Text = "";
			// 
			// label7
			// 
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(16, 136);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(136, 16);
			this.label7.TabIndex = 8;
			this.label7.Text = "Valid Values";
			// 
			// txtP3
			// 
			this.txtP3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP3.Location = new System.Drawing.Point(16, 16);
			this.txtP3.Name = "txtP3";
			this.txtP3.Size = new System.Drawing.Size(88, 22);
			this.txtP3.TabIndex = 7;
			this.txtP3.Text = "";
			// 
			// txtP3Desc
			// 
			this.txtP3Desc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP3Desc.Location = new System.Drawing.Point(16, 72);
			this.txtP3Desc.Multiline = true;
			this.txtP3Desc.Name = "txtP3Desc";
			this.txtP3Desc.Size = new System.Drawing.Size(808, 56);
			this.txtP3Desc.TabIndex = 6;
			this.txtP3Desc.Text = "";
			// 
			// label8
			// 
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(16, 48);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(136, 16);
			this.label8.TabIndex = 5;
			this.label8.Text = "Description";
			// 
			// grpboxP2
			// 
			this.grpboxP2.Controls.Add(this.lblP2Defa);
			this.grpboxP2.Controls.Add(this.txtP2Valid);
			this.grpboxP2.Controls.Add(this.label5);
			this.grpboxP2.Controls.Add(this.txtP2);
			this.grpboxP2.Controls.Add(this.txtP2Desc);
			this.grpboxP2.Controls.Add(this.label6);
			this.grpboxP2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpboxP2.Location = new System.Drawing.Point(8, 480);
			this.grpboxP2.Name = "grpboxP2";
			this.grpboxP2.Size = new System.Drawing.Size(848, 232);
			this.grpboxP2.TabIndex = 2;
			this.grpboxP2.TabStop = false;
			this.grpboxP2.Text = "Parameter 2";
			// 
			// lblP2Defa
			// 
			this.lblP2Defa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblP2Defa.Location = new System.Drawing.Point(120, 16);
			this.lblP2Defa.Name = "lblP2Defa";
			this.lblP2Defa.Size = new System.Drawing.Size(680, 24);
			this.lblP2Defa.TabIndex = 11;
			this.lblP2Defa.Text = "Default Value";
			// 
			// txtP2Valid
			// 
			this.txtP2Valid.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP2Valid.Location = new System.Drawing.Point(20, 166);
			this.txtP2Valid.Multiline = true;
			this.txtP2Valid.Name = "txtP2Valid";
			this.txtP2Valid.Size = new System.Drawing.Size(808, 56);
			this.txtP2Valid.TabIndex = 9;
			this.txtP2Valid.Text = "";
			// 
			// label5
			// 
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(20, 142);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(136, 16);
			this.label5.TabIndex = 8;
			this.label5.Text = "Valid Values";
			// 
			// txtP2
			// 
			this.txtP2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP2.Location = new System.Drawing.Point(16, 16);
			this.txtP2.Name = "txtP2";
			this.txtP2.Size = new System.Drawing.Size(88, 22);
			this.txtP2.TabIndex = 7;
			this.txtP2.Text = "";
			// 
			// txtP2Desc
			// 
			this.txtP2Desc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP2Desc.Location = new System.Drawing.Point(16, 72);
			this.txtP2Desc.Multiline = true;
			this.txtP2Desc.Name = "txtP2Desc";
			this.txtP2Desc.Size = new System.Drawing.Size(808, 56);
			this.txtP2Desc.TabIndex = 6;
			this.txtP2Desc.Text = "";
			// 
			// label6
			// 
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(16, 48);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(136, 16);
			this.label6.TabIndex = 5;
			this.label6.Text = "Description";
			// 
			// grpboxP1
			// 
			this.grpboxP1.Controls.Add(this.lblP1Defa);
			this.grpboxP1.Controls.Add(this.txtP1Valid);
			this.grpboxP1.Controls.Add(this.label4);
			this.grpboxP1.Controls.Add(this.txtP1);
			this.grpboxP1.Controls.Add(this.txtP1Desc);
			this.grpboxP1.Controls.Add(this.label3);
			this.grpboxP1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpboxP1.Location = new System.Drawing.Point(8, 240);
			this.grpboxP1.Name = "grpboxP1";
			this.grpboxP1.Size = new System.Drawing.Size(848, 232);
			this.grpboxP1.TabIndex = 1;
			this.grpboxP1.TabStop = false;
			this.grpboxP1.Text = "Parameter 1";
			// 
			// lblP1Defa
			// 
			this.lblP1Defa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblP1Defa.Location = new System.Drawing.Point(128, 16);
			this.lblP1Defa.Name = "lblP1Defa";
			this.lblP1Defa.Size = new System.Drawing.Size(680, 24);
			this.lblP1Defa.TabIndex = 11;
			this.lblP1Defa.Text = "Default Value";
			// 
			// txtP1Valid
			// 
			this.txtP1Valid.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP1Valid.Location = new System.Drawing.Point(20, 166);
			this.txtP1Valid.Multiline = true;
			this.txtP1Valid.Name = "txtP1Valid";
			this.txtP1Valid.Size = new System.Drawing.Size(808, 56);
			this.txtP1Valid.TabIndex = 9;
			this.txtP1Valid.Text = "";
			// 
			// label4
			// 
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(20, 142);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(136, 16);
			this.label4.TabIndex = 8;
			this.label4.Text = "Valid Values";
			// 
			// txtP1
			// 
			this.txtP1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP1.Location = new System.Drawing.Point(16, 16);
			this.txtP1.Name = "txtP1";
			this.txtP1.Size = new System.Drawing.Size(88, 22);
			this.txtP1.TabIndex = 7;
			this.txtP1.Text = "";
			// 
			// txtP1Desc
			// 
			this.txtP1Desc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP1Desc.Location = new System.Drawing.Point(16, 72);
			this.txtP1Desc.Multiline = true;
			this.txtP1Desc.Name = "txtP1Desc";
			this.txtP1Desc.Size = new System.Drawing.Size(808, 56);
			this.txtP1Desc.TabIndex = 6;
			this.txtP1Desc.Text = "";
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(16, 48);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(136, 16);
			this.label3.TabIndex = 5;
			this.label3.Text = "Description";
			// 
			// grpboxFVSCmd
			// 
			this.grpboxFVSCmd.Controls.Add(this.label18);
			this.grpboxFVSCmd.Controls.Add(this.label9);
			this.grpboxFVSCmd.Controls.Add(this.cmbFilter);
			this.grpboxFVSCmd.Controls.Add(this.label2);
			this.grpboxFVSCmd.Controls.Add(this.txtFVSCmdVariantList);
			this.grpboxFVSCmd.Controls.Add(this.txtFVSCmdDesc);
			this.grpboxFVSCmd.Controls.Add(this.label1);
			this.grpboxFVSCmd.Controls.Add(this.cmbFVSCmd);
			this.grpboxFVSCmd.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpboxFVSCmd.Location = new System.Drawing.Point(8, 8);
			this.grpboxFVSCmd.Name = "grpboxFVSCmd";
			this.grpboxFVSCmd.Size = new System.Drawing.Size(848, 224);
			this.grpboxFVSCmd.TabIndex = 0;
			this.grpboxFVSCmd.TabStop = false;
			this.grpboxFVSCmd.Text = "FVS Command";
			// 
			// label18
			// 
			this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.Location = new System.Drawing.Point(16, 48);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(64, 16);
			this.label18.TabIndex = 9;
			this.label18.Text = "Command";
			// 
			// label9
			// 
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(16, 24);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(48, 16);
			this.label9.TabIndex = 8;
			this.label9.Text = "Filter";
			// 
			// cmbFilter
			// 
			this.cmbFilter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFilter.Items.AddRange(new object[] {
														   "All",
														   "Fire and Fuels Extension"});
			this.cmbFilter.Location = new System.Drawing.Point(96, 16);
			this.cmbFilter.Name = "cmbFilter";
			this.cmbFilter.Size = new System.Drawing.Size(224, 24);
			this.cmbFilter.TabIndex = 7;
			this.cmbFilter.Text = "All";
			this.cmbFilter.SelectedIndexChanged += new System.EventHandler(this.cmbFilter_SelectedIndexChanged);
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(328, 24);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(80, 16);
			this.label2.TabIndex = 6;
			this.label2.Text = "Variant List";
			// 
			// txtFVSCmdVariantList
			// 
			this.txtFVSCmdVariantList.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtFVSCmdVariantList.Location = new System.Drawing.Point(328, 48);
			this.txtFVSCmdVariantList.Name = "txtFVSCmdVariantList";
			this.txtFVSCmdVariantList.Size = new System.Drawing.Size(488, 22);
			this.txtFVSCmdVariantList.TabIndex = 5;
			this.txtFVSCmdVariantList.Text = "";
			// 
			// txtFVSCmdDesc
			// 
			this.txtFVSCmdDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtFVSCmdDesc.Location = new System.Drawing.Point(16, 95);
			this.txtFVSCmdDesc.Multiline = true;
			this.txtFVSCmdDesc.Name = "txtFVSCmdDesc";
			this.txtFVSCmdDesc.Size = new System.Drawing.Size(808, 120);
			this.txtFVSCmdDesc.TabIndex = 4;
			this.txtFVSCmdDesc.Text = "";
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(16, 71);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(136, 16);
			this.label1.TabIndex = 3;
			this.label1.Text = "Description";
			// 
			// cmbFVSCmd
			// 
			this.cmbFVSCmd.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFVSCmd.Location = new System.Drawing.Point(96, 48);
			this.cmbFVSCmd.Name = "cmbFVSCmd";
			this.cmbFVSCmd.Size = new System.Drawing.Size(224, 24);
			this.cmbFVSCmd.TabIndex = 0;
			this.cmbFVSCmd.SelectedIndexChanged += new System.EventHandler(this.cmbFVSCmd_SelectedIndexChanged);
			// 
			// grpboxP7
			// 
			this.grpboxP7.Controls.Add(this.lblP7Defa);
			this.grpboxP7.Controls.Add(this.txtP7Valid);
			this.grpboxP7.Controls.Add(this.label16);
			this.grpboxP7.Controls.Add(this.txtP7);
			this.grpboxP7.Controls.Add(this.txtP7Desc);
			this.grpboxP7.Controls.Add(this.label17);
			this.grpboxP7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grpboxP7.Location = new System.Drawing.Point(8, 1648);
			this.grpboxP7.Name = "grpboxP7";
			this.grpboxP7.Size = new System.Drawing.Size(848, 224);
			this.grpboxP7.TabIndex = 7;
			this.grpboxP7.TabStop = false;
			this.grpboxP7.Text = "Parameter 7";
			// 
			// lblP7Defa
			// 
			this.lblP7Defa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblP7Defa.Location = new System.Drawing.Point(120, 16);
			this.lblP7Defa.Name = "lblP7Defa";
			this.lblP7Defa.Size = new System.Drawing.Size(680, 24);
			this.lblP7Defa.TabIndex = 10;
			this.lblP7Defa.Text = "Default Value";
			// 
			// txtP7Valid
			// 
			this.txtP7Valid.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP7Valid.Location = new System.Drawing.Point(16, 160);
			this.txtP7Valid.Multiline = true;
			this.txtP7Valid.Name = "txtP7Valid";
			this.txtP7Valid.Size = new System.Drawing.Size(808, 56);
			this.txtP7Valid.TabIndex = 9;
			this.txtP7Valid.Text = "";
			// 
			// label16
			// 
			this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.Location = new System.Drawing.Point(16, 136);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(136, 16);
			this.label16.TabIndex = 8;
			this.label16.Text = "Valid Values";
			// 
			// txtP7
			// 
			this.txtP7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP7.Location = new System.Drawing.Point(16, 16);
			this.txtP7.Name = "txtP7";
			this.txtP7.Size = new System.Drawing.Size(88, 22);
			this.txtP7.TabIndex = 7;
			this.txtP7.Text = "";
			// 
			// txtP7Desc
			// 
			this.txtP7Desc.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtP7Desc.Location = new System.Drawing.Point(16, 72);
			this.txtP7Desc.Multiline = true;
			this.txtP7Desc.Name = "txtP7Desc";
			this.txtP7Desc.Size = new System.Drawing.Size(808, 56);
			this.txtP7Desc.TabIndex = 6;
			this.txtP7Desc.Text = "";
			// 
			// label17
			// 
			this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.Location = new System.Drawing.Point(16, 48);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(136, 16);
			this.label17.TabIndex = 5;
			this.label17.Text = "Description";
			// 
			// uc_rx_fvscmd_edit
			// 
			this.Controls.Add(this.groupBox1);
			this.Name = "uc_rx_fvscmd_edit";
			this.Size = new System.Drawing.Size(872, 2184);
			this.groupBox1.ResumeLayout(false);
			this.panel1.ResumeLayout(false);
			this.grpboxOther.ResumeLayout(false);
			this.grpboxP6.ResumeLayout(false);
			this.grpboxP5.ResumeLayout(false);
			this.grpboxP4.ResumeLayout(false);
			this.grpboxP3.ResumeLayout(false);
			this.grpboxP2.ResumeLayout(false);
			this.grpboxP1.ResumeLayout(false);
			this.grpboxFVSCmd.ResumeLayout(false);
			this.grpboxP7.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void panel1_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.grpboxP1.Width = this.panel1.Width - this.grpboxP1.Left-25; 
				this.txtP1Desc.Width = this.grpboxP1.Width - this.txtP1Desc.Left-25;
				this.txtP1Valid.Width = this.txtP1Desc.Width;

				this.grpboxFVSCmd.Width = this.grpboxP1.Width;
				this.txtFVSCmdDesc.Width = this.txtP1Desc.Width;

				this.grpboxP2.Width = this.grpboxP1.Width; 
				this.txtP2Desc.Width = this.txtP1Desc.Width;
				this.txtP2Valid.Width = this.txtP1Desc.Width;

				this.grpboxP3.Width = this.grpboxP1.Width; 
				this.txtP3Desc.Width = this.txtP1Desc.Width;
				this.txtP3Valid.Width = this.txtP1Desc.Width;

				this.grpboxP4.Width = this.grpboxP1.Width; 
				this.txtP4Desc.Width = this.txtP1Desc.Width;
				this.txtP4Valid.Width = this.txtP1Desc.Width;

				this.grpboxP5.Width = this.grpboxP1.Width; 
				this.txtP5Desc.Width = this.txtP1Desc.Width;
				this.txtP5Valid.Width = this.txtP1Desc.Width;

				this.grpboxP6.Width = this.grpboxP1.Width; 
				this.txtP6Desc.Width = this.txtP1Desc.Width;
				this.txtP6Valid.Width = this.txtP1Desc.Width;

				this.grpboxP7.Width = this.grpboxP1.Width; 
				this.txtP7Desc.Width = this.txtP1Desc.Width;
				this.txtP7Valid.Width = this.txtP1Desc.Width;

				this.grpboxOther.Width = this.grpboxP1.Width; 
				this.txtOther.Width = this.txtP1Desc.Width;
				this.txtOther.Width = this.txtP1Desc.Width;


			}
			catch (Exception err)
			{
			}
		}

		private void cmbFVSCmd_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.LoadFvsCommandProperties(cmbFVSCmd.Text);
		}
		protected void ReloadFVSCommandList()
		{	
			this.cmbFVSCmd.Items.Clear();
			this.cmbFVSCmd.Text="";
			m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL);

			if (m_oAdo.m_OleDbDataReader.HasRows)
			{
				while (m_oAdo.m_OleDbDataReader.Read())
				{
					if (m_oAdo.m_OleDbDataReader["fvscmd"] != System.DBNull.Value)
					{
						cmbFVSCmd.Items.Add(Convert.ToString(m_oAdo.m_OleDbDataReader["fvscmd"]));
					}
				}
			}
			m_oAdo.m_OleDbDataReader.Close();
		}
		private void cmbFilter_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		   this.ReloadFVSCommandList();	
		   this.LoadFvsCommandProperties("");
		}
		public bool RxPackageEdit
		{
			get {return _bRxPackageEdit;}
			set {_bRxPackageEdit=value;}
		}
	
		public FIA_Biosum_Manager.frmRxItemFvsCmdItem ReferenceFormFvsCmdItem
		{
			get {return this._frmFvsCmdItem;}
			set {this._frmFvsCmdItem=value;}

		}
	}
}

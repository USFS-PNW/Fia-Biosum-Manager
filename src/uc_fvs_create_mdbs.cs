using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Threading;
using System.Collections.Generic;

namespace FIA_Biosum_Manager
{


    public class uc_fvs_create_mdbs : System.Windows.Forms.UserControl
    {
        private GroupBox groupBox1;
        private TextBox textBox1;
        private ToolTip toolTip1;
        private IContainer components;
        private Button button1;
        private string m_strProjDir;
        private Queries m_oQueries;
        private Tables m_oTables;
        public int m_intError;
        private ado_data_access m_ado;
        private dao_data_access m_dao;
        private string m_strTempMDBFileConnectionString;
        private env m_oEnv;
        private bool m_bDebug;

        public uc_fvs_create_mdbs(string p_strProjDir)
        {
            InitializeComponent();
            this.m_strProjDir = p_strProjDir;

            this.m_oQueries = new Queries();
            m_oQueries.m_oFvs.LoadDatasource = true;
            m_oQueries.m_oFIAPlot.LoadDatasource = true;
            m_oQueries.LoadDatasources(true);

            if (m_oQueries.m_oFvs.m_strFvsTreeTable.Trim().Length == 0)
            {
                m_oQueries.m_oFvs.m_strFvsTreeTable = Tables.FVS.DefaultFVSTreeTableName;
            }

            this.m_ado = new ado_data_access();
            this.m_dao = new dao_data_access();
            this.m_strTempMDBFileConnectionString = this.m_ado.getMDBConnString(this.m_oQueries.m_strTempDbFile, "", "");
            this.m_ado.OpenConnection(this.m_strTempMDBFileConnectionString);
            if (this.m_ado.m_intError != 0)
            {
                this.m_intError = this.m_ado.m_intError;
                this.m_ado = null;
                return;

            }
            this.m_oEnv = new env();
            this.m_bDebug = frmMain.g_bDebug;
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Location = new System.Drawing.Point(32, 17);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(684, 377);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "groupBox1";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(20, 336);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(126, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Create MDBs";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(20, 21);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(638, 285);
            this.textBox1.TabIndex = 1;
            // 
            // uc_fvs_create_mdbs
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_fvs_create_mdbs";
            this.Size = new System.Drawing.Size(746, 421);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        private void CreateMDBs_Main()
        {
            // When you need to see how to update progress bars:
            // RunAppend_Main in uc_fvs_output.cs! (also good for seeing how to interact with m_intError;
            var textBoxValue = frmMain.g_oDelegate.GetControlPropertyValue(this.textBox1, "Text", false);
            frmMain.g_oDelegate.SetControlPropertyValue(this.textBox1, "Text", textBoxValue+ " Magic");
            return;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //if (this.lstFvsOutput.CheckedItems.Count == 0)
            //{
            //    MessageBox.Show("No Boxes Are Checked", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
            //    return;
            //}

            //this.m_frmTherm = new frmTherm(((frmDialog)ParentForm), "FVS OUT DATA",
            //    "FVS Output", "2");
            //m_frmTherm.Visible = false;
            //this.m_frmTherm.lblMsg.Text = "";
            //this.m_frmTherm.TopMost = true;

            //this.cmbStep.Enabled = false;
            //this.btnExecute.Enabled = false;
            //this.btnChkAll.Enabled = false;
            //this.btnClearAll.Enabled = false;
            //this.btnRefresh.Enabled = false;
            //this.btnClose.Enabled = false;
            //this.btnHelp.Enabled = false;
            //this.btnCancel.Visible = false;
            //this.btnViewLogFile.Enabled = false;
            //this.btnViewPostLogFile.Enabled = false;
            //this.btnAuditDb.Enabled = false;
            //this.btnPostAppendAuditDb.Enabled = false;

            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Maximum", 100);
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Minimum", 0);
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 0);
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg2, "Text", "Overall Progress");
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "");
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Visible", true);
            //m_frmTherm.Show((frmDialog)ParentForm);

            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            frmMain.g_oDelegate.m_oThread = new System.Threading.Thread(new System.Threading.ThreadStart(CreateMDBs_Main));
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
        }

        

        internal void uc_fvs_create_mdbs_Resize()
        {
            return;
        }
    }
}

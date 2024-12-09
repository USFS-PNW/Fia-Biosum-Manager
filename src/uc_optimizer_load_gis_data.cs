using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    public partial class uc_optimizer_load_gis_data : UserControl
    {
        private GisTools m_oGisTools = new GisTools();
        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultTreatmentOptimizerFile;
        string m_strMasterDb;
        string m_strSourceField;

        public uc_optimizer_load_gis_data(frmMain p_frmMain)
        {
            InitializeComponent();
            this.m_oEnv = new env();
            load_values();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            ((frmDialog)ParentForm).ParentControl.Enabled = true;
            this.ParentForm.Close();
        }

        public void load_values()
        {
            m_strMasterDb = frmMain.g_oEnv.strApplicationDataDirectory.Trim() + frmMain.g_strBiosumDataDir +
                "\\" + Tables.TravelTime.DefaultMasterTravelTimeDbFile;
            if (System.IO.File.Exists(m_strMasterDb))
            {
                txtDataFile.Text = m_strMasterDb;
                System.IO.FileInfo fileInfo = new System.IO.FileInfo(m_strMasterDb);
                lblLastUpdated.Text = $@"Last Updated {fileInfo.LastWriteTime.ToString("MMMM dd, yyyy")}";
                double fileSizeKb = fileInfo.Length / 1000 ;
                lblFileSize.Text = $@"File Size {fileSizeKb} KB";
            }
            else
            {
                MessageBox.Show("The source gis_travel_times_master.db is required but does not exist in the " +
                    System.IO.Path.GetDirectoryName(m_strMasterDb) + " folder. \r\n\r\n" +
                    "Please download a copy of this database into the FIABiosum folder!!", "FIA Biosum");
                BtnLoad.Enabled = false;
                return;
            }
            bool bTablesHaveData;
            bool bTablesExist = m_oGisTools.CheckForExistingDataSqlite(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(), out bTablesHaveData);
            ckBackupData.Visible = bTablesHaveData;
            if (ckBackupData.Visible)
            {
                ckBackupData.Checked = true;
                string strMessage = "BioSum has found existing data in your gis data tables. Do you wish to overwrite existing data? " +
                            "This process cannot be reversed!!";
                DialogResult res = MessageBox.Show(strMessage, "FIA BioSum", MessageBoxButtons.YesNo);
                if (res != DialogResult.Yes)
                {
                    return;
                }
            }
            // Check for existence of MoveDist_ft_REPLACEMENT field
            m_strSourceField = "MoveDist_ft_REPLACEMENT";
            if (m_oGisTools.CheckPlotGisTable(m_strMasterDb, m_strSourceField))
            {
                ckUpdateYardingDistance.Visible = true;
                ckUpdateYardingDistance.Checked = true;
            }
            else
            {
                ckUpdateYardingDistance.Visible = false;
                ckUpdateYardingDistance.Checked = false;
            }
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            bool bSuccess = true;
            // Validate max one-way hours
            double dblMaxHours = 999;
            bSuccess = Double.TryParse(txtMaxOneWayHours.Text, out dblMaxHours);
            if (!bSuccess || dblMaxHours < 0)
            {
                MessageBox.Show("A valid numeric value is required!!", "FIA BioSum");
                txtMaxOneWayHours.Focus();
                return;
            }

            frmMain.g_oFrmMain.ActivateStandByAnimation(
                frmMain.g_oFrmMain.WindowState,
                frmMain.g_oFrmMain.Left,
                frmMain.g_oFrmMain.Height,
                frmMain.g_oFrmMain.Width,
                frmMain.g_oFrmMain.Top);
            frmMain.g_sbpInfo.Text = "Loading GIS Data...Stand by";

            if (ckBackupData.Checked == true)
            {
                frmMain.g_sbpInfo.Text = "Backing up old GIS Data...Stand by";
                bSuccess = m_oGisTools.BackupGisData();
            }
            if (bSuccess == true)
            {
                if (!ckUpdateYardingDistance.Checked)
                {
                    m_strSourceField = "";
                }
                int intRowCount = m_oGisTools.LoadSqliteGisData(m_strSourceField, dblMaxHours, (frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile));
                if (intRowCount < 1)
                {
                    MessageBox.Show("An error occurred while loading the GIS data!!", "FIA BioSum");
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();
                    return;
                }
            }
            else
            {
                MessageBox.Show("An error occurred while backing up the tables. GIS data load terminated!!", "FIA BioSum");
                frmMain.g_sbpInfo.Text = "Ready";
                frmMain.g_oFrmMain.DeactivateStandByAnimation();
                return;
            }
            frmMain.g_sbpInfo.Text = "Ready";
            frmMain.g_oFrmMain.DeactivateStandByAnimation();
            MessageBox.Show("If you updated existing GIS data, verify the selected sites in Treatment Optimizer. GIS data successfully loaded!");
        }
    }


}

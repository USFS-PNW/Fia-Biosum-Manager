using SQLite.ADO;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for version_control.
	/// </summary>
	public class version_control
	{
		private ado_data_access m_oAdo = new ado_data_access();
		private dao_data_access m_oDao = new dao_data_access();
		const int APP_VERSION_MAJOR=0;
		const int APP_VERSION_MINOR1=1;
		const int APP_VERSION_MINOR2=2;		
		private string[] m_strAppVerArray=null;
		private string[] m_strDbVerArray=null;
		private string m_strTempDbFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");
		private string m_strProjectVersion="1.0.0";
		private string[] m_strProjectVersionArray=null;
        private bool bProjectVersionArrayUsed = false;
		private string _strProjDir="";
		private FIA_Biosum_Manager.frmMain _oFrmMain=null;
		string strUpdateSql="";
		string strInsertSql="";
        string m_strFVSTreeTable = "";
        string m_strFVSTreeDbFile = "";
        object oMissing = System.Reflection.Missing.Value;


		string[] m_strScenarioRuleDefinitionsTableArray = {"SCENARIO",
															"SCENARIO_COSTS",
															"SCENARIO_DATASOURCE",
															"SCENARIO_HARVEST_COST_COLUMNS",
															"SCENARIO_LAND_OWNER_GROUPS",
															"SCENARIO_PLOT_FILTER",
															"SCENARIO_PLOT_FILTER_MISC",
															"SCENARIO_COND_FILTER",
															"SCENARIO_COND_FILTER_MISC",
															"SCENARIO_PSITES",
															"SCENARIO_RX_INTENSITY",
		                                                    "SCENARIO_FVS_VARIABLES_TIEBREAKER",
		                                                    "SCENARIO_FVS_VARIABLES_OPTIMIZATION",
		                                                    "SCENARIO_FVS_VARIABLES_OVERALL_EFFECTIVE",
		                                                    "SCENARIO_FVS_VARIABLES",
                                                            "SCENARIO_PROCESSOR_SCENARIO_SELECT"};

        string[] m_strProcessorScenarioRuleDefinitionsTableArray = {"SCENARIO",
															"SCENARIO_COSTS",
															"SCENARIO_DATASOURCE",
															"SCENARIO_HARVEST_COST_COLUMNS",
															"SCENARIO_LAND_OWNER_GROUPS",
															"SCENARIO_PLOT_FILTER",
															"SCENARIO_PLOT_FILTER_MISC",
															"SCENARIO_COND_FILTER",
															"SCENARIO_COND_FILTER_MISC",
															"SCENARIO_PSITES",
															"SCENARIO_RX_INTENSITY",
		                                                    "SCENARIO_FVS_VARIABLES_TIEBREAKER",
		                                                    "SCENARIO_FVS_VARIABLES_OPTIMIZATION",
		                                                    "SCENARIO_FVS_VARIABLES_OVERALL_EFFECTIVE",
		                                                    "SCENARIO_FVS_VARIABLES",
                                                            "SCENARIO_PROCESSOR_SCENARIO_SELECT"};


		

		public version_control()
		{
			//
			// TODO: Add constructor logic here
			//
			m_oDao.CreateMDB(m_strTempDbFile);
			m_oDao.m_DaoWorkspace.Close();
			
			


		}
		/// <summary>
		/// Check the project's application version and update to the current version
		/// if different.
		/// </summary>
		public void PerformVersionCheck()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//version_control.PerformVersionCheck \r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }
            frmMain.g_oFrmMain.ActivateStandByAnimation(
                frmMain.g_oFrmMain.WindowState,
                frmMain.g_oFrmMain.Left,
                frmMain.g_oFrmMain.Height,
                frmMain.g_oFrmMain.Width,
                frmMain.g_oFrmMain.Top);
            bool bPerformCheck = true;
            string strProjVersionFile = this.ReferenceProjectDirectory.Trim() + "\\application.version";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: strProjVersionFile=" + strProjVersionFile + "\r\n");

            m_strAppVerArray = frmMain.g_oUtils.ConvertListToArray(frmMain.g_strAppVer, ".");
            string strProjVersion = "";
            if (System.IO.File.Exists(strProjVersionFile))
            {

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: open application version file\r\n");
                try
                {
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: instantiate streamreader and open file\r\n");
                    //Open the file in a stream reader.
                    System.IO.StreamReader s = new System.IO.StreamReader(strProjVersionFile);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  application version file opened with no errors\r\n");

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  streamreader.ReadLine\r\n");
                    //Split the first line into the columns       
                    strProjVersion = s.ReadLine();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  streamreader.ReadLine successful\r\n");
                    s.Close();
                    s.Dispose();
                    s = null;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  streamreader close and dispose successful\r\n");

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  strProjVersion=" + strProjVersion + "\r\n");
                    if (strProjVersion.Trim() == frmMain.g_strAppVer.Trim())
                    {
                        bPerformCheck = false;
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  bPerformCheck=false\r\n");
                    }
                    else
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  bPerformCheck=true\r\n");

                        if (strProjVersion.Trim().Length > 0)
                        {
                            this.m_strProjectVersion = strProjVersion.Trim();
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Convert " + m_strProjectVersion + " to an array\r\n");
                            m_strProjectVersionArray = frmMain.g_oUtils.ConvertListToArray(m_strProjectVersion, ".");
                            bProjectVersionArrayUsed = true;
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            {
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Conversion to array completed\r\n");
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: m_strProjectVersionArray[APP_VERSION_MAJOR]=" + m_strProjectVersionArray[APP_VERSION_MAJOR] + " m_strProjectVersionArray[APP_VERSION_MINOR1]=" + m_strProjectVersionArray[APP_VERSION_MINOR1] + " m_strProjectVersionArray[APP_VERSION_MINOR2]=" + m_strProjectVersionArray[APP_VERSION_MINOR2] + "\r\n");
                            }

                        }
                    }
                }
                catch (Exception err)
                {
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: !!Error opening Application.Version File!! ERROR=" + err.Message + "r\n");
                }
            }
            else
            {
                m_strProjectVersionArray = frmMain.g_oUtils.ConvertListToArray(m_strProjectVersion, ".");
                bProjectVersionArrayUsed = true;
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Check m_strProjectVersionArray\r\n");
            
            try
            {
                if (bProjectVersionArrayUsed)
                {
                    if (m_strProjectVersionArray != null)
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Project Version=" + m_strProjectVersion + " Application Version=" + frmMain.g_strAppVer + "\r\n");
                        //no database updates need to be made with these versions
                        if (frmMain.g_strAppVer == "5.1.2" && m_strProjectVersion == "5.1.1")
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression01\r\n");
                            UpdateProjectVersionFile(strProjVersionFile);
                            bPerformCheck = false;
                        }
                        else if (frmMain.g_strAppVer == "5.0.5" && m_strProjectVersion == "5.0.4")
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression02\r\n");

                            UpdateProjectVersionFile(strProjVersionFile);
                            bPerformCheck = false;

                        }
                        else if (frmMain.g_strAppVer == "5.1.0" && (m_strProjectVersion == "5.0.4" ||
                                                                    m_strProjectVersion == "5.0.5"))
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression03\r\n");

                            UpdateProjectVersionFile(strProjVersionFile);
                            bPerformCheck = false;
                        }

                    }
                }
            }
            catch (Exception error)
            {
                bPerformCheck = false;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: !!Error opening Application.Version File!! ERROR=" + error.Message + "r\n");
            }

            //check for partial update
            if (bPerformCheck)
            {
                if (m_strProjectVersion.Trim().Length > 0)
                {
                    if ((frmMain.g_strAppVer == "5.1.4" || frmMain.g_strAppVer=="5.1.5" ||
                        frmMain.g_strAppVer == "5.1.6" ||
                        frmMain.g_strAppVer == "5.1.7" || 
                        frmMain.g_strAppVer == "5.2.0" ||
                        frmMain.g_strAppVer == "5.2.1" ||
                        frmMain.g_strAppVer == "5.2.2") && m_strProjectVersionArray[APP_VERSION_MAJOR] == "5")
                    {
                        //
                        //plot fvs variant assignments table had a major upgrade release with 5.1.4
                        //
                        if (m_strProjectVersionArray[APP_VERSION_MINOR1] == "0" ||
                            (m_strProjectVersionArray[APP_VERSION_MINOR1] == "1" &&
                             (m_strProjectVersionArray[APP_VERSION_MINOR2] == "0" ||
                               m_strProjectVersionArray[APP_VERSION_MINOR2] == "1" ||
                               m_strProjectVersionArray[APP_VERSION_MINOR2] == "2" ||
                               m_strProjectVersionArray[APP_VERSION_MINOR2] == "3")))
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression04\r\n");

                            //project version is 5.0.? or 5.1.0 to 5.1.3
                            UpdateFVSPlotVariantAssignmentsTable();
                            UpdateProjectVersionFile(strProjVersionFile);
                            bPerformCheck = false;
                        }
                        else
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression05\r\n");

                            UpdateProjectVersionFile(strProjVersionFile);
                            bPerformCheck = false;
                        }
                            
                    }
                    if ((Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) < 5) ||
                       (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                       Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 2 &&
                       Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) == 0))
                        UpgradeFVSOutTreeListFiles();

                    if (frmMain.g_strAppVer == "5.3.0" || (frmMain.g_strAppVer=="5.3.1" && m_strProjectVersion!="5.3.0"))
                    {
                        if ((Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) <= 4) ||
                            ((Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) <= 5 &&
                             Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 2)))
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression06\r\n");

                            UpgradeToPrePostSeqNumMatrix();
                            UpdateAuditDbFile_5_3_0();
                            UpdateDatasources_5_3_0();
                            System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show("Version control has detected that the FVS_POTFIRE tables need to be converted to the new Biosum version.\r\n\r\nBy selecting 'Yes', Biosum will convert the FVS POTFIRE table to the new version specificiations.\r\n\r\nBy selecting 'No', the FVS_POTFIRE tables will need to be recreated through FVS.\r\n\r\nDo you want Biosum version control to upgrade the FVS_POTFIRE tables? (Y/N)", "FIA Bisoum", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                            if (result == System.Windows.Forms.DialogResult.Yes)
                            {
                                UpgradeFVSOutPOTFireFiles();
                            }
                            result = System.Windows.Forms.MessageBox.Show("Version control has detected that the PREPOST tables need to be converted to the new Biosum version.\r\n\r\nBy selecting 'Yes', Biosum will convert the PREPOST tables to the new version specificiations.\r\n\r\nBy selecting 'No', the PREPOST tables will need to be repopulated by Biosum in the FVS OUTPUT process.\r\n\r\nDo you want Biosum version control to upgrade the PREPOST tables? (Y/N)", "FIA Bisoum", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                            if (result == System.Windows.Forms.DialogResult.Yes)
                            {
                                UpgradeFVSOutPREPOSTTables();
                            }
                            
                                                       
                            if ((Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5))
                            {
                                UpdateProjectVersionFile(strProjVersionFile);
                                bPerformCheck = false;
                            }


                            
                        }
                              
                    }
                    else if (frmMain.g_strAppVer == "5.3.1" && m_strProjectVersion == "5.3.0")
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression07\r\n");

                                UpdateProjectVersionFile(strProjVersionFile);
                                bPerformCheck = false;
                    }
                    else if (frmMain.g_strAppVer == "5.3.2" && (m_strProjectVersion == "5.3.0" || m_strProjectVersion=="5.3.1"))
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression08\r\n");

                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if ((frmMain.g_strAppVer == "5.4.0" || frmMain.g_strAppVer=="5.4.1") && (m_strProjectVersion=="5.3.2" || m_strProjectVersion == "5.3.0" || m_strProjectVersion == "5.3.1"))
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression09\r\n");

                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if (frmMain.g_strAppVer == "5.4.1" && m_strProjectVersion == "5.4.0")
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression10\r\n");

                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if (frmMain.g_strAppVer == "5.5.1" && m_strProjectVersion == "5.5.0")
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression11\r\n");

                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if (frmMain.g_strAppVer == "5.5.2" && m_strProjectVersion == "5.5.1")
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression12\r\n");

                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if (frmMain.g_strAppVer == "5.5.3" && (m_strProjectVersion == "5.5.2" || m_strProjectVersion=="5.5.1"))
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression13\r\n");

                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if (frmMain.g_strAppVer == "5.5.7" && (frmMain.g_strAppVer == "5.5.6" || frmMain.g_strAppVer == "5.5.5" || frmMain.g_strAppVer == "5.5.4" || m_strProjectVersion == "5.5.3" || m_strProjectVersion == "5.5.2" || m_strProjectVersion == "5.5.1"))
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression14\r\n");

                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if ((frmMain.g_strAppVer == "5.5.0" || frmMain.g_strAppVer == "5.5.1" || frmMain.g_strAppVer == "5.5.2" || frmMain.g_strAppVer == "5.5.3" || frmMain.g_strAppVer == "5.5.4" || frmMain.g_strAppVer == "5.5.5" || frmMain.g_strAppVer == "5.5.6" || frmMain.g_strAppVer == "5.5.7") && (m_strProjectVersion == "5.4.0" || m_strProjectVersion == "5.4.1" || m_strProjectVersion == "5.4.2"))
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression15\r\n");

                        UpdateFVSPlotVariantAssignmentsTable();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    else if ((frmMain.g_strAppVer == "5.5.7" || frmMain.g_strAppVer == "5.5.6" || frmMain.g_strAppVer == "5.5.5" || frmMain.g_strAppVer == "5.5.4" || frmMain.g_strAppVer == "5.5.3" || frmMain.g_strAppVer == "5.5.2" || frmMain.g_strAppVer == "5.5.1") && (m_strProjectVersion == "5.5.0" || m_strProjectVersion == "5.5.1" || m_strProjectVersion == "5.5.2" || m_strProjectVersion == "5.5.3" || m_strProjectVersion == "5.5.4" || m_strProjectVersion == "5.5.5" || m_strProjectVersion == "5.5.6" || m_strProjectVersion == "5.5.7"))
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Expression16\r\n");

                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                         Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) > 5) &&
                         (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                          Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 5))
                    {
                        UpdateDatasources_5_6_0();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;

                    }
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) > 6) &&
                           (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 6))
                    {
                        if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                             Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) > 5) &&
                            (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                             Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) == 6))
                        {
                            if (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) < 2)
                            {
                                Update_5_6_2();

                            }

                            if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                                Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) == 7))
                            {
                                UpdateDatasources_5_7_0();
                                Update_5_7_0();
                            }

                            UpdateProjectVersionFile(strProjVersionFile);
                            bPerformCheck = false;
                        }
                    }
                    //5.7.7 is Processor redesign
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) >= 7 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) >= 7) &&
                           (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 7 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) < 7))
                    {
                        UpdateDatasources_5_7_7();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    //5.7.8 updates harvest_costs and scenario_harvest_method tables
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) >= 7 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) >= 8) &&
                           (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 7 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) < 8))
                    {
                        UpdateDatasources_5_7_8();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    //5.7.9 moves tree diam and species groups into scenario-specific tables
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) >= 7 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) >= 9) &&
                           (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 7 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) < 9))
                    {
                        UpdateDatasources_5_7_9();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    //5.8.0 restructures tree_species table and moves reference tables into user's %appData% directory
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) >= 8 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) >= 0) &&
                            (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) == 7))
                    {
                        UpdateDatasources_5_8_0();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    //5.8.4 modifications to Core, tree species, harvest methods tables and new OPCOST reference database
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) >= 8 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) >= 4) &&
                           (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 8 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) < 1))
                    {
                        UpdateDatasources_5_8_4();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    //5.8.5 modifications for DWM; Updated configurations for tree_species and OpCost
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) >= 8 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) >= 5) &&
                           (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 8 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) < 5))
                    {
                        UpdateDatasources_5_8_5();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    //5.8.6 modifications to Core; Phase 1 of redesign
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) >= 8 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) >= 6) &&
                           (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 8 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) < 6))
                    {
                        UpdateDatasources_5_8_6();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    //5.8.7 modifications to Optimizer tables; Clean-up Plot and Cond tables
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) >= 8 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) >= 7) &&
                           (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 8 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) < 7))
                    {
                        UpdateDatasources_5_8_9();  // Allow v5.8.6 to upgrade directly to v5.8.9
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    //5.8.8 Adding Optimizer context database
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) >= 8 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) >= 8) &&
                           (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 8 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) < 9))
                    {
                        UpdateDatasources_5_8_9(); // Allow v5.8.7 to upgrade to v5.8.9
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    //5.8.10 New columns required for national FIADB calculations
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) >= 8 &&
                            Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) >= 9) &&
                           (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 8 &&
                            Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) < 10))
                    {
                        UpdateDatasources_5_8_10();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    //5.9.0 More new columns for FIADB; No more Oracle interface
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) >= 9) &&
                        (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) == 8 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) >= 10))
                    {
                        UpdateDatasources_5_9_0();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    //Upgrade to 5.10.1 from 5.9.0 or 5.10.0
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) >= 10 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) == 1) &&
                        (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) >= 9))
                    {
                        UpdateDatasources_5_10_1();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                }
            }

            if (bPerformCheck)
            {
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: bPerformCheck=true\r\n");


                string strInfo = frmMain.g_sbpInfo.Text;
                frmMain.g_sbpInfo.Text = "Version Update: Checking Project Table...Stand by";
                CheckProjectTable();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Core Analysis Scenario Rule Definitions...Stand by";
                CheckCoreScenarioRuleDefinitionTables();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Project Datasource Table Records...Stand by";
                CheckProjectDatasourceTableRecords();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Project Datasource Tables...Stand by";
                CheckProjectDatasourceTables();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Pre-Populated Reference Tables...Stand by";
                CheckProjectReferenceDatasourceTables();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Core Scenario Datasource Table Records...Stand by";
                CheckCoreScenarioDatasourceTableRecords();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Processor Scenario Datasource Table Records...Stand by";
                CheckProcessorScenarioDatasourceTableRecords();


                frmMain.g_sbpInfo.Text = "Version Update: Checking Rx Values...Stand By";
                CheckRxValues();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Fvs Out Pre-Post Rx Values...Stand By";
                CheckFVSOutPrePostValues();

                //frmMain.g_sbpInfo.Text = "Version Update: Checking Core Scenario Results Table Records...Stand by";
                //CheckCoreScenarioResultsTables();
                CleanUp();

                UpdateProjectVersionFile(strProjVersionFile);

                frmMain.g_sbpInfo.Text = strInfo;
            }
            frmMain.g_oFrmMain.DeactivateStandByAnimation();

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Leaving\r\n");


        }
        private void UpdateProjectVersionFile(string p_strProjectVersionFile)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//version_control.UpdateProjectVersionFile \r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }
            if (System.IO.File.Exists(p_strProjectVersionFile))
                System.IO.File.Delete(p_strProjectVersionFile);
            frmMain.g_oUtils.WriteText(p_strProjectVersionFile, frmMain.g_strAppVer);
            
        }
		/// <summary>
		/// Update project table with any current version changes.
		/// </summary>
		private void CheckProjectTable()
		{
			int x,y;
            //create a link to the project table
		   // m_oDao.CreateTableLink(m_strTempDbFile,"project_current",this.ProjectDirectory,"project");
            //m_oAdo.OpenConnection(m_oAdo.getMDBConnString(this.m_strTempDbFile,"",""));
			m_oAdo.OpenConnection(m_oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                                                          frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile,"",""));
		    //frmMain.g_oTables.m_oProject.CreateProjectTable(m_oAdo,m_oAdo.m_OleDbConnection);
			if (m_oAdo.ColumnExist(m_oAdo.m_OleDbConnection,"project","application_version")==false)
			{
				m_oAdo.m_strSQL = "ALTER TABLE project ADD COLUMN application_version TEXT(11)";
				m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection,m_oAdo.m_strSQL);
			}
			else
			{
				
			}

			m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
			

			//string[] strSourceColumnsArray = m_oAdo.getFieldNamesArray(m_oAdo.m_OleDbConnection,"SELECT * FROM project_current");
			//string[] strDestColumnsArray = m_oAdo.getFieldNamesArray(m_oAdo.m_OleDbConnection,"SELECT * FROM project");
			//for (x=0;x<=strDestColumnsArray.Length-1;x++)
			//{

			//}



			
           
		}
		/// <summary>
		/// Update the scenario rule definitions db file with current version.
		/// </summary>
		private void CheckCoreScenarioRuleDefinitionTables()
		{
			
			int x,y,z;
			string strColumns="";
			string strSql="";
			string strTempDbFile="";
			string[] strProjectTablesArray;
			string[] strAppTablesArray;
			System.Data.OleDb.OleDbConnection oConn;
			string[] strDestColArray=null;
			string[] strSourceColArray=null;
			dao_data_access oDao = new dao_data_access();
			ado_data_access oAdo = new ado_data_access();
			string strTableName="";
			string strFile1=this.ReferenceProjectDirectory.Trim() + "\\core\\db\\scenario_core_rule_definitions.mdb";
			string strFile2=this.ReferenceProjectDirectory.Trim() + "\\core\\db\\scenario.mdb";
			if (System.IO.File.Exists(strFile1)==false) 
			{
				this.ReferenceMainForm.frmProject.uc_project1.CreateOptimizerScenarioRuleDefinitionDbAndTables(strFile1);
				if (System.IO.File.Exists(strFile2)==true)
				{
					string[] strTablesToLink = null;
					oDao.getTableNames(strFile2,ref strTablesToLink);
					for (x=0;x<=strTablesToLink.Length-1;x++)
					{
						if (strTablesToLink[x] == null) break;
						if (oDao.TableExists(strFile1,strTablesToLink[x].Trim() + "_temp"))
							oDao.DeleteTableFromMDB(strFile1,strTablesToLink[x].Trim() + "_temp");
							  
						oDao.CreateTableLink(strFile1,strTablesToLink[x].Trim() + "_temp",strFile2,strTablesToLink[x].Trim());


					}
					oAdo.OpenConnection(oAdo.getMDBConnString(strFile1,"",""));
					for (x=0;x<=strTablesToLink.Length-1;x++)
					{
						if (strTablesToLink[x] == null) break;
						strSql="";
						switch (strTablesToLink[x].Trim().ToUpper())
						{
							case "SCENARIO":
								strSql = "INSERT INTO scenario SELECT * FROM scenario_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_COSTS":
								strDestColArray = oAdo.getFieldNamesArray(oAdo.m_OleDbConnection,"SELECT * FROM scenario_costs");
								strSourceColArray=oAdo.getFieldNamesArray(oAdo.m_OleDbConnection,"SELECT * FROM scenario_costs_temp");
								strColumns="";
								for (y=0;y<=strDestColArray.Length - 1;y++)
								{
									for (z=0;z<=strSourceColArray.Length-1;z++)
									{
										if (strSourceColArray[z].Trim().ToUpper() == 
											strDestColArray[y].Trim().ToUpper())
										{
											strColumns=strColumns + strSourceColArray[z].Trim() + ",";
											break;
										}
									}
								}
								if (strColumns.Trim().Length > 0)
								{
									strColumns=strColumns.Substring(0,strColumns.Length - 1);
									strSql = "INSERT INTO scenario_costs " + 
										     "(" + strColumns + ") " + 
										     "SELECT " + strColumns + " " + 
										     "FROM scenario_costs_temp " + 
											 "WHERE scenario_id IS NOT NULL AND " + 
											 "LEN(TRIM(scenario_id)) > 0";
								}

								break;
							case "SCENARIO_DATASOURCE":
                                strSql = "INSERT INTO scenario_datasource SELECT * FROM scenario_datasource_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0 AND TRIM(UCASE(TABLE_TYPE)) NOT IN ('FIRE AND FUEL EFFECTS','ADDITIONAL HARVEST COSTS','TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS','TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES')";
								break;
							case "SCENARIO_HARVEST_COST_COLUMNS":
								strSql = "INSERT INTO scenario_harvest_cost_columns SELECT * FROM scenario_harvest_cost_columns_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_LAND_OWNER_GROUPS":
								strSql = "INSERT INTO scenario_land_owner_groups SELECT * FROM scenario_land_owner_groups_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_PLOT_FILTER":
								strSql = "INSERT INTO scenario_plot_filter SELECT * FROM scenario_plot_filter_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_PLOT_FILTER_MISC":
								strSql = "INSERT INTO scenario_cond_filter_misc SELECT * FROM scenario_plot_filter_misc_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_COND_FILTER":
								strSql = "INSERT INTO scenario_cond_filter SELECT * FROM scenario_cond_filter_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_COND_FILTER_MISC":
								strSql = "INSERT INTO scenario_cond_filter_misc SELECT * FROM scenario_cond_filter_misc_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_PSITES":
								strSql = "INSERT INTO scenario_psites SELECT * FROM scenario_psites_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
							case "SCENARIO_RX_INTENSITY":
								strSql = "INSERT INTO scenario_rx_intensity SELECT * FROM scenario_rx_intensity_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
								break;
                            case "SCENARIO_PROCESSOR_SCENARIO_SELECT":
                                strSql = "INSERT INTO scenario_processor_scenario_select SELECT * FROM scenario_processor_scenario_select_temp WHERE scenario_id IS NOT NULL AND LEN(TRIM(scenario_id)) > 0";
                                break;
						}
						if (strSql.Trim().Length > 0)
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,strSql);

					}

					strSql = "UPDATE scenario SET file='scenario_core_rule_definitions.mdb'";
				    oAdo.SqlNonQuery(oAdo.m_OleDbConnection,strSql);

					//remove the linked tables
					for (x=0;x<=strTablesToLink.Length-1;x++)
					{
						if (strTablesToLink[x]==null) break;
						strSql="";
						if (oAdo.TableExist(oAdo.m_OleDbConnection,strTablesToLink[x].Trim() + "_temp"))
						{
							strSql = "DROP TABLE " + strTablesToLink[x].Trim() + "_temp";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,strSql);
						}

					}
					oAdo.CloseConnection(oAdo.m_OleDbConnection);
					//reset path variables in the scenario db
					this.ReferenceMainForm.frmProject.uc_project1.SetProjectPathEnvironmentVariables();
					

				}
			}
			oDao.m_DaoWorkspace.Close();

			//check and update table structures
			//create temp access table that will contain links
			//create a temp file that will contain a temporary copy of 
			//the latest tables
			oDao = new dao_data_access();
			strTempDbFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");
			oDao.CreateMDB(strTempDbFile);
			oDao.m_DaoWorkspace.Close();
			oConn = new System.Data.OleDb.OleDbConnection();
			oConn.ConnectionString = oAdo.getMDBConnString(strTempDbFile,"","");
			oConn.Open();
			oAdo.OpenConnection(oAdo.getMDBConnString(strFile1,"",""));

			strAppTablesArray = oAdo.getTableNames(oAdo.m_OleDbConnection);
			for (x=0;x<=strAppTablesArray.Length-1;x++)
			{
				for (y=0;y<=this.m_strScenarioRuleDefinitionsTableArray.Length - 1;y++)
				{
					if (strAppTablesArray[x].Trim().ToUpper()==this.m_strScenarioRuleDefinitionsTableArray[y].Trim().ToUpper())
					{
						strTableName="";
						switch (m_strScenarioRuleDefinitionsTableArray[y])
						{
							case "SCENARIO":
								strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableName;
								frmMain.g_oTables.m_oScenario.CreateScenarioTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_COSTS":
								strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioCostsTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_DATASOURCE":
								strTableName=Tables.Scenario.DefaultScenarioDatasourceTableName;
								frmMain.g_oTables.m_oScenario.CreateScenarioDatasourceTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_HARVEST_COST_COLUMNS":
								strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioHarvestCostColumnsTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioHarvestCostColumnsTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_LAND_OWNER_GROUPS":
								strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioLandOwnerGroupsTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_PLOT_FILTER":
								strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioPlotFilterTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_PLOT_FILTER_MISC":
								strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterMiscTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioPlotFilterMiscTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_COND_FILTER":
								strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioCondFilterTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_COND_FILTER_MISC":
								strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioCondFilterMiscTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_PSITES":
								strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioPSitesTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_RX_INTENSITY":
                                strTableName = Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLastTieBreakRankTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioLastTieBreakRankTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_FVS_VARIABLES_TIEBREAKER":
								strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesTieBreakerTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_FVS_VARIABLES_OPTIMIZATION":
								strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesOptimizationTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_FVS_VARIABLES_OVERALL_EFFECTIVE":
								strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesOverallEffectiveTable(oAdo,oConn,strTableName);
								break;
							case "SCENARIO_FVS_VARIABLES":
								strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesTable(oAdo,oConn,strTableName);
								break;
                            case "SCENARIO_PROCESSOR_SCENARIO_SELECT":
                                strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName;
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioProcessorScenarioSelectTable(oAdo,oConn,strTableName);

                                break;


						}
						if (strTableName.Trim().Length > 0)
						{
							string strEmptyTable = strTableName + "_work_temp";
							oAdo.m_strSQL = "SELECT COUNT(*) FROM (SELECT TOP 1 * FROM " + strTableName + ")";
							//check if the projects version of the table has records
							if ((int)oAdo.getRecordCount(oAdo.m_OleDbConnection,oAdo.m_strSQL, strTableName) > 0)
							{
								bool bModified=false;
								if (oAdo.TableExist(oAdo.m_OleDbConnection,strEmptyTable))
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + strEmptyTable);
								//create a temp copy with no records (empty table) of the projects version of the table
								oAdo.m_strSQL = "SELECT TOP 1 * INTO " + strEmptyTable + " FROM " + strTableName;
								oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
								oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DELETE FROM " + strEmptyTable);
								//compare the application version of the table with the 
								//projects version of the table and 
								//modify the projects version of the table structure if needs be.
								bModified = oAdo.ReconcileTableColumns(oAdo.m_OleDbConnection,
									strEmptyTable,
									oConn,strTableName);
								if (bModified)
								{
									//create the indexes for the empty table structure
									//oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent,oAdoCurrent.m_OleDbConnection,
									//	oDs.m_strDataSource[x,Datasource.TABLETYPE].Trim(),strEmptyTable.Trim());
									//insert all the records into our new empty table
									oAdo.m_strSQL = "INSERT INTO " + strEmptyTable + " SELECT * FROM " + strTableName;
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
									//drop the current table
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + strTableName);
									//create the table structure giving it the current table name
									oAdo.m_strSQL = "SELECT TOP 1 * INTO " + strTableName + " FROM " + strEmptyTable;
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
									//delete the 1 record in the current table
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DELETE FROM " + strTableName);
									//insert all the records from the new empty table to the current table
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"INSERT INTO " + strTableName + " SELECT *  FROM " + strEmptyTable.Trim());
									//drop the empty table
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + strEmptyTable);
								}
								else
								{
									oAdo.SqlNonQuery(oAdo.m_OleDbConnection,"DROP TABLE " + strEmptyTable);
								}
							}
							else
							{
								if (oAdo.ReconcileTableColumns(oAdo.m_OleDbConnection,
									strTableName,
									oConn,strTableName))
								{
								}
							}
						}
					}
				}
			}
			//create any tables that do not exist
			for (y=0;y<=this.m_strScenarioRuleDefinitionsTableArray.Length - 1;y++)
			{
				if (!oAdo.TableExist(oAdo.m_OleDbConnection,this.m_strScenarioRuleDefinitionsTableArray[y].Trim().ToUpper()))
				{
					strTableName="";

					switch (m_strScenarioRuleDefinitionsTableArray[y])
					{
						case "SCENARIO":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableName;
							frmMain.g_oTables.m_oScenario.CreateScenarioTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_COSTS":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioCostsTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_DATASOURCE":
							strTableName=Tables.Scenario.DefaultScenarioTableName;
							frmMain.g_oTables.m_oScenario.CreateScenarioDatasourceTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_HARVEST_COST_COLUMNS":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioHarvestCostColumnsTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioHarvestCostColumnsTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_LAND_OWNER_GROUPS":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioLandOwnerGroupsTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_PLOT_FILTER":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioPlotFilterTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_PLOT_FILTER_MISC":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterMiscTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioPlotFilterMiscTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_COND_FILTER":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioCondFilterTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_COND_FILTER_MISC":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioCondFilterMiscTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_PSITES":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioPSitesTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_RX_INTENSITY":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLastTieBreakRankTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioLastTieBreakRankTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_FVS_VARIABLES_TIEBREAKER":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesTieBreakerTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_FVS_VARIABLES_OPTIMIZATION":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesOptimizationTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_FVS_VARIABLES_OVERALL_EFFECTIVE":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesOverallEffectiveTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
						case "SCENARIO_FVS_VARIABLES":
							strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesTable(oAdo,oAdo.m_OleDbConnection,strTableName);
							break;
                        case "SCENARIO_PROCESSOR_SCENARIO_SELECT":
                            strTableName=Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName;
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioProcessorScenarioSelectTable(oAdo,oConn,strTableName);
                            break;
					}
				}
			}

			oAdo.CloseConnection(oAdo.m_OleDbConnection);
			oAdo=null;
			oDao=null;
			
		}

		/// <summary>
		/// Check and update the project datasource table with the latest version of table type entries
		/// </summary>
		private void CheckProjectDatasourceTableRecords()
		{
		    
			int x;
			int y;
			FIA_Biosum_Manager.Datasource oDs = new Datasource();
			oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
			oDs.m_strDataSourceTableName = "datasource";
			oDs.m_strScenarioId="";
			oDs.LoadTableColumnNamesAndDataTypes=false;
			oDs.LoadTableRecordCount=false;
			oDs.populate_datasource_array();

            

			ado_data_access oAdo=new ado_data_access();
			string strDbFile="";
            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));
			//
            //drop obsolete table types 
            //
            //additional harvest costs table type
            oAdo.m_strSQL = "DELETE FROM datasource WHERE TRIM(UCASE(table_type)) IN ('ADDITIONAL HARVEST COSTS','TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES')";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            //
			//
			//make sure all the current table type data sources are accounted for
			//
			for (x=0;x<=Datasource.g_strProjectDatasourceTableTypesArray.Length - 1;x++)
			{
				//check if the latest datasource table type exists 
				y = oDs.getDataSourceTableNameRow(Datasource.g_strProjectDatasourceTableTypesArray[x].Trim());
				if (y==-1)
				{
					//table type does not exist so create it
					switch (Datasource.g_strProjectDatasourceTableTypesArray[x].Trim().ToUpper())
					{
						case "PLOT":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\db",
								frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableDbFile),
                                frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName);
							break;
						case "CONDITION":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\db",
								frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableDbFile),
								frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName);
							break;
						case "TREE":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\db",
								frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableDbFile),
								frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName);
							break;
						
						case "OWNER GROUPS":  
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\db",
								"ref_master.mdb",
								"owner_groups");
							break;
						case "TREATMENT PRESCRIPTIONS":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Treatment Prescriptions'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'" + strDbFile.Trim() + "'," + 
								"'rx');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "TREE SPECIES":
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Tree Species'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'ref_master.mdb'," + 
								"'tree_species');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "FVS TREE SPECIES":
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('FVS Tree Species'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'ref_master.mdb'," + 
								"'fvs_tree_species');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						
						case "TRAVEL TIMES":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\gis\\db",
								frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.TravelTime.DefaultTravelTimeTableDbFile),
                                Tables.TravelTime.DefaultTravelTimeTableName);
							break;
						case "PROCESSING SITES":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\gis\\db",
								frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oTravelTime.DefaultProcessingSiteTableDbFile),
                                Tables.TravelTime.DefaultProcessingSiteTableName);
							break;
						case "FIADB FVS VARIANT":
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('FIADB FVS Variant'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'ref_master.mdb'," + 
								"'fiadb_fvs_variant');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						
						case "PLOT AND CONDITION RECORD AUDIT":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\db",
								frmMain.g_oUtils.getFileName(Tables.Audit.DefaultCondAuditTableDbFile),
                                Tables.Audit.DefaultCondAuditTableName);
							break;
						case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
							oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
								Datasource.g_strProjectDatasourceTableTypesArray[x].Trim(),
								ReferenceProjectDirectory.Trim() + "\\db",
                                frmMain.g_oUtils.getFileName(Tables.Audit.DefaultCondRxAuditTableDbFile),
                                Tables.Audit.DefaultCondRxAuditTableName);
							break;
						case "POPULATION EVALUATION":
							strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Population Evaluation'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'" + strDbFile.Trim() + "'," + 
								"'pop_eval');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "POPULATION ESTIMATION UNIT":
							strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Population Estimation Unit'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'" + strDbFile.Trim() + "'," + 								
								"'pop_estn_unit');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "POPULATION STRATUM":
							strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Population Stratum'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'" + strDbFile.Trim() + "'," + 		
								"'pop_stratum');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "POPULATION PLOT STRATUM ASSIGNMENT":
							strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Population Plot Stratum Assignment'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'" + strDbFile.Trim() + "'," + 		
								"'pop_plot_stratum_assgn');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
						case "SITE TREE":
							strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
								"('Site Tree'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'master.mdb'," + 
								"'sitetree');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
							break;
                        
                        //version 5 additions
                        case "TREATMENT PRESCRIPTIONS ASSIGNED FVS COMMANDS":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxFvsCommandTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Prescriptions Assigned FVS Commands'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'fvsmaster.mdb'," + 
								"'rx_fvs_commands');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PRESCRIPTIONS HARVEST COST COLUMNS":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxHarvestCostColumnsTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Prescriptions Harvest Cost Columns'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'fvsmaster.mdb'," +
                                "'rx_harvest_cost_columns');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PRESCRIPTION CATEGORIES":
                            //strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(frmMain.g_oTables.m_oReference.DefaultRxHarvestCostTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Prescription Categories'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'ref_master.mdb'," +
                                "'fvs_rx_category');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PRESCRIPTION SUBCATEGORIES":
                            oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Prescription Subcategories'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'ref_master.mdb'," +
                                "'fvs_rx_subcategory');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PACKAGES":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxPackageTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Packages'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'fvsmaster.mdb'," +
                                "'rxpackage');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PACKAGE ASSIGNED FVS COMMANDS":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxPackageFvsCommandTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Package Assigned FVS Commands'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'fvsmaster.mdb'," +
                                "'rxpackage_fvs_commands');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PACKAGE MEMBERS":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxPackageMembersTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Package Members'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'fvsmaster.mdb'," +
                                "'rxpackage_members');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "TREATMENT PACKAGE FVS COMMANDS ORDER":
                            strDbFile = frmMain.g_oUtils.getFileNameUsingLastIndexOf(Tables.FVS.DefaultRxPackageFvsCommandsOrderTableDbFile);
							oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Treatment Package FVS Commands Order'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'fvsmaster.mdb'," +
                                "'rxpackage_fvs_commands_order');";
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            break;
                        case "FVS COMMANDS":
                            oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('FVS Commands'," + 
								"'" + ReferenceProjectDirectory.Trim() + "\\db'," + 
								"'ref_fvscommands.mdb'," +
                                "'fvs_commands');";
							oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
                            break;
                        case "HARVEST METHODS":
                            oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('Harvest Methods'," +
                                "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                                "'ref_master.mdb'," +
                                "'harvest_methods');";
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            break;
                        //version 5.2.1 additions
                        case "FVS WESTERN TREE SPECIES TRANSLATOR":
                            oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('FVS Western Tree Species Translator'," +
                                "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                                "'ref_master.mdb'," +
                                "'FVS_WesternTreeSpeciesTranslator');";
                             oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            break;
                        case "FVS EASTERN TREE SPECIES TRANSLATOR":
                            oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                                "('FVS Eastern Tree Species Translator'," +
                                "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                                "'ref_master.mdb'," +
                                "'FVS_EasternTreeSpeciesTranslator');";
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            break;


                        //case "ADDITIONAL HARVEST COSTS":
                        //    oAdo.m_strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                        //        "('Additional Harvest Costs'," +
                        //        "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                        //        "'master.mdb'," +
                        //        "'additional_harvest_costs');";
                        //    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                        //    break;
					}
				}
			}
			oAdo.CloseConnection(oAdo.m_OleDbConnection);
		}
		/// <summary>
		/// Project table type Db files, tables and table columns are updated to the current version.
		/// Each record in the datasource table is checked to make sure the db file exists,
		/// the table exists, and all the columns exist. 
		/// </summary>
		private void CheckProjectDatasourceTables()
		{

			int x;
			FIA_Biosum_Manager.Datasource oDs = new Datasource();
			oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
			oDs.m_strDataSourceTableName = "datasource";
			oDs.m_strScenarioId="";
			oDs.LoadTableColumnNamesAndDataTypes=false;
			oDs.LoadTableRecordCount=false;
			oDs.populate_datasource_array();

			ado_data_access oAdo=new ado_data_access();
			ado_data_access oAdoCurrent=null;

			dao_data_access oDao=null;
			string strCurrDbFile="";
			string strDbFile="";
			string strTempDbFile="";
			string strTempTableName="";
			System.Data.OleDb.OleDbConnection oConn=null;
			
			

			//open the project db file
			oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" + 
				frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile,"",""));

            


			oAdo.m_strSQL = "SELECT * FROM datasource ORDER BY path,file";
			oAdo.SqlQueryReader(oAdo.m_OleDbConnection,oAdo.m_strSQL);

			if (oAdo.m_OleDbDataReader.HasRows)
			{
				//sequentially read each record
				while (oAdo.m_OleDbDataReader.Read())
				{
					if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value &&
						oAdo.m_OleDbDataReader["file"] != System.DBNull.Value)
					{
					    //get the array number of the current datasource table type
						x=oDs.getDataSourceTableNameRow(oAdo.m_OleDbDataReader["table_type"].ToString().Trim());

						
						strDbFile = oAdo.m_OleDbDataReader["path"].ToString().Trim() + "\\" + 
							        oAdo.m_OleDbDataReader["file"].ToString().Trim();

						//check if the DB file has changed
						if (strDbFile.Trim().ToUpper() != strCurrDbFile.Trim().ToUpper())
						{
							//close the current open DB file
							if (oAdoCurrent!=null) oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
							else oAdoCurrent = new ado_data_access();

							//make sure the db file exists
							if (System.IO.File.Exists(strDbFile)==false)
							{
								//create the DB file
								if (oDao==null) oDao = new dao_data_access();
								oDao.CreateMDB(strDbFile);
							}
							//open a connection to the DB file
							oAdoCurrent.OpenConnection(oAdo.getMDBConnString(strDbFile,"",""));
							strCurrDbFile=strDbFile;

						}
						//check the table
						if (oDs.m_strDataSource[x,Datasource.TABLESTATUS]=="NF") //NF=table not found
						{
							//create the table
							switch (oDs.m_strDataSource[x,Datasource.TABLETYPE].Trim().ToUpper())
							{
								
								case "PLOT":
									frmMain.g_oTables.m_oFIAPlot.CreatePlotTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName);
									break;
								case "CONDITION":
									frmMain.g_oTables.m_oFIAPlot.CreateConditionTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName);
									break;
								case "TREE":
									frmMain.g_oTables.m_oFIAPlot.CreateTreeTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName);
									break;
								case "OWNER GROUPS":  
									frmMain.g_oTables.m_oReference.CreateOwnerGroupsTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.Reference.DefaultOwnerGroupsTableName);
									break;
								case "TREATMENT PRESCRIPTIONS":
									frmMain.g_oTables.m_oFvs.CreateRxTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.FVS.DefaultRxTableName);
									break;
								case "TREE SPECIES":
									frmMain.g_oTables.m_oReference.CreateTreeSpeciesTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.Reference.DefaultFVSTreeSpeciesTableName);
									break;
								case "FVS TREE SPECIES":
									frmMain.g_oTables.m_oReference.CreateFVSTreeSpeciesTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.Reference.DefaultFVSTreeSpeciesTableName);
									break;
								
								case "TRAVEL TIMES":
									frmMain.g_oTables.m_oTravelTime.CreateTravelTimeTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.TravelTime.DefaultTravelTimeTableName);
									break;
								case "PROCESSING SITES":
                                    frmMain.g_oTables.m_oTravelTime.CreateProcessingSiteTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.TravelTime.DefaultProcessingSiteTableName);
									break;
								case "FVS TREE LIST FOR PROCESSOR":
									frmMain.g_oTables.m_oFvs.CreateFVSOutProcessorIn(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.FVS.DefaultFVSTreeTableName);
									break;
								case "FIADB FVS VARIANT":
									frmMain.g_oTables.m_oReference.CreateFiadbFVSVariantTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.Reference.DefaultFiadbFVSVariantTableName);
									break;
								
								case "PLOT AND CONDITION RECORD AUDIT":
                                    frmMain.g_oTables.m_oAudit.CreateCondAuditTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.Audit.DefaultCondAuditTableName);
									break;
								case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
                                    frmMain.g_oTables.m_oAudit.CreatePlotCondRxAuditTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.Audit.DefaultCondRxAuditTableName);
									break;
								case "POPULATION EVALUATION":
									frmMain.g_oTables.m_oFIAPlot.CreatePopEvalTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName);
									break;
								case "POPULATION ESTIMATION UNIT":
									frmMain.g_oTables.m_oFIAPlot.CreatePopEstnUnitTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName);
									break;
								case "POPULATION STRATUM":
									frmMain.g_oTables.m_oFIAPlot.CreatePopStratumTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName);
									break;
								case "POPULATION PLOT STRATUM ASSIGNMENT":
									frmMain.g_oTables.m_oFIAPlot.CreatePopPlotStratumAssgnTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName);
									break;
								case "SITE TREE":
									frmMain.g_oTables.m_oFIAPlot.CreateSiteTreeTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName);
									break;
                                //version 5 additions
                                case "TREATMENT PRESCRIPTIONS ASSIGNED FVS COMMANDS":
                                    frmMain.g_oTables.m_oFvs.CreateRxFvsCommandsTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.FVS.DefaultRxFvsCommandTableName);
                                    break;
                                case "TREATMENT PRESCRIPTIONS HARVEST COST COLUMNS":
                                    frmMain.g_oTables.m_oFvs.CreateRxHarvestCostColumnTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection,Tables.FVS.DefaultRxHarvestCostColumnsTableName);
                                    break;
                                case "TREATMENT PRESCRIPTION CATEGORIES":
                                    frmMain.g_oTables.m_oReference.CreateRxCategoryTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.Reference.DefaultRxCategoryTableName);
                                    break;
                                case "TREATMENT PRESCRIPTION SUBCATEGORIES":
                                    frmMain.g_oTables.m_oReference.CreateRxSubCategoryTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.Reference.DefaultRxSubCategoryTableName);
                                    break;
                                case "TREATMENT PACKAGES":
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.FVS.DefaultRxPackageTableName);
                                    break;
                                case "TREATMENT PACKAGE ASSIGNED FVS COMMANDS":
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageFvsCommandsTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.FVS.DefaultRxPackageFvsCommandTableName);
                                    break;
                                case "TREATMENT PACKAGE MEMBERS":
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageMembersTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.FVS.DefaultRxPackageMembersTableName);
                                    break;
                                case "TREATMENT PACKAGE FVS COMMANDS ORDER":
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageFvsCommandsOrderTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.FVS.DefaultRxPackageFvsCommandsOrderTableName);
                                    break;
                               
                                case "HARVEST METHODS":
                                    frmMain.g_oTables.m_oReference.CreateHarvestMethodsTable(oAdoCurrent,oAdoCurrent.m_OleDbConnection, Tables.Reference.DefaultHarvestMethodsTableName);
                                    break;
                                //5.2.1 additions
                                case "FVS WESTERN TREE SPECIES TRANSLATOR":
                                    frmMain.g_oTables.m_oReference.CreateFVSWesternSpeciesTranslatorTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.Reference.DefaultFVSWesternTreeSpeciesTableName);
                                    break;
                                case "FVS EASTERN TREE SPECIES TRANSLATOR":
                                    frmMain.g_oTables.m_oReference.CreateFVSEasternSpeciesTranslatorTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, Tables.Reference.DefaultFVSEasternTreeSpeciesTableName);
                                    break;
                                

                               
							}
						}
						else
						{
							//update columns in the table
							if (oDao==null)
							{ 
								//create a temp file that will contain a temporary copy of 
								//the latest tables
								oDao = new dao_data_access();
								strTempDbFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");
								oDao.CreateMDB(strTempDbFile);
								oConn = new System.Data.OleDb.OleDbConnection();
								oConn.ConnectionString = oAdoCurrent.getMDBConnString(strTempDbFile,"","");
								oConn.Open();
							}

							strTempTableName="";
							//create an empty table structure of the table
							switch (oDs.m_strDataSource[x,Datasource.TABLETYPE].Trim().ToUpper())
							{
								case "PLOT":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName;
									frmMain.g_oTables.m_oFIAPlot.CreatePlotTable(oAdoCurrent,oConn,
										     strTempTableName);
									break;
								case "CONDITION":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName;
									frmMain.g_oTables.m_oFIAPlot.CreateConditionTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "TREE":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName;
									frmMain.g_oTables.m_oFIAPlot.CreateTreeTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "OWNER GROUPS":  
									strTempTableName = Tables.Reference.DefaultOwnerGroupsTableName;
									frmMain.g_oTables.m_oReference.CreateOwnerGroupsTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "TREATMENT PRESCRIPTIONS":
									strTempTableName = Tables.FVS.DefaultRxTableName;
									frmMain.g_oTables.m_oFvs.CreateRxTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "TREE SPECIES":
									strTempTableName = Tables.Reference.DefaultTreeSpeciesTableName;
									frmMain.g_oTables.m_oReference.CreateTreeSpeciesTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "FVS TREE SPECIES":
									strTempTableName = Tables.Reference.DefaultFVSTreeSpeciesTableName;
									frmMain.g_oTables.m_oReference.CreateFVSTreeSpeciesTable(oAdoCurrent,oConn,strTempTableName);
									break;
							
								case "TRAVEL TIMES":
                                    strTempTableName = Tables.TravelTime.DefaultTravelTimeTableName;
									frmMain.g_oTables.m_oTravelTime.CreateTravelTimeTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "PROCESSING SITES":
                                    strTempTableName = Tables.TravelTime.DefaultProcessingSiteTableName;
									frmMain.g_oTables.m_oTravelTime.CreateProcessingSiteTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "FVS TREE LIST FOR PROCESSOR":
									strTempTableName = Tables.FVS.DefaultFVSTreeTableName;
									frmMain.g_oTables.m_oFvs.CreateFVSOutProcessorIn(oAdoCurrent,oConn,strTempTableName);
									break;
								case "FIADB FVS VARIANT":
									strTempTableName = Tables.Reference.DefaultFiadbFVSVariantTableName;
									frmMain.g_oTables.m_oReference.CreateFiadbFVSVariantTable(oAdoCurrent,oConn,strTempTableName);
									break;
								
								case "PLOT AND CONDITION RECORD AUDIT":
                                    strTempTableName = Tables.Audit.DefaultCondAuditTableName;
									frmMain.g_oTables.m_oAudit.CreateCondAuditTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
                                    strTempTableName = Tables.Audit.DefaultCondRxAuditTableName;
									frmMain.g_oTables.m_oAudit.CreatePlotCondRxAuditTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "POPULATION EVALUATION":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName;
									frmMain.g_oTables.m_oFIAPlot.CreatePopEvalTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "POPULATION ESTIMATION UNIT":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName;
									frmMain.g_oTables.m_oFIAPlot.CreatePopEstnUnitTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "POPULATION STRATUM":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName;
									frmMain.g_oTables.m_oFIAPlot.CreatePopStratumTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "POPULATION PLOT STRATUM ASSIGNMENT":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName;
									frmMain.g_oTables.m_oFIAPlot.CreatePopPlotStratumAssgnTable(oAdoCurrent,oConn,strTempTableName);
									break;
								case "SITE TREE":
									strTempTableName = frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName;
									frmMain.g_oTables.m_oFIAPlot.CreateSiteTreeTable(oAdoCurrent,oConn,strTempTableName);
									break;
                                //version 5 additions
                                case "TREATMENT PRESCRIPTIONS ASSIGNED FVS COMMANDS":
                                    strTempTableName = Tables.FVS.DefaultRxFvsCommandTableName;
                                    frmMain.g_oTables.m_oFvs.CreateRxFvsCommandsTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PRESCRIPTIONS HARVEST COST COLUMNS":
                                    strTempTableName = Tables.FVS.DefaultRxHarvestCostColumnsTableName;
                                    frmMain.g_oTables.m_oFvs.CreateRxHarvestCostColumnTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PRESCRIPTION CATEGORIES":
                                    strTempTableName = Tables.Reference.DefaultRxCategoryTableName;
                                    frmMain.g_oTables.m_oReference.CreateRxCategoryTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PRESCRIPTION SUBCATEGORIES":
                                    strTempTableName = Tables.Reference.DefaultRxSubCategoryTableName;
                                    frmMain.g_oTables.m_oReference.CreateRxSubCategoryTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PACKAGES":
                                    strTempTableName = Tables.FVS.DefaultRxPackageTableName;
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PACKAGE ASSIGNED FVS COMMANDS":
                                    strTempTableName = Tables.FVS.DefaultRxPackageFvsCommandTableName;
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageFvsCommandsTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PACKAGE MEMBERS":
                                    strTempTableName = Tables.FVS.DefaultRxPackageMembersTableName;
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageMembersTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "TREATMENT PACKAGE FVS COMMANDS ORDER":
                                    strTempTableName = Tables.FVS.DefaultRxPackageFvsCommandsOrderTableName;
                                    frmMain.g_oTables.m_oFvs.CreateRxPackageFvsCommandsOrderTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                
                                case "HARVEST METHODS":
                                    strTempTableName = Tables.Reference.DefaultHarvestMethodsTableName;
                                    frmMain.g_oTables.m_oReference.CreateHarvestMethodsTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                //5.2.1 additions
                                case "FVS WESTERN TREE SPECIES TRANSLATOR":
                                    strTempTableName = Tables.Reference.DefaultFVSWesternTreeSpeciesTableName;
                                    frmMain.g_oTables.m_oReference.CreateFVSWesternSpeciesTranslatorTable(oAdoCurrent, oConn, strTempTableName);
                                    break;
                                case "FVS EASTERN TREE SPECIES TRANSLATOR":
                                    strTempTableName = Tables.Reference.DefaultFVSEasternTreeSpeciesTableName;
                                    frmMain.g_oTables.m_oReference.CreateFVSEasternSpeciesTranslatorTable(oAdoCurrent, oConn, strTempTableName);

                                    break;

							}
							if (strTempTableName.Trim().Length > 0)
							{
								
								
								string strEmptyTable = oDs.m_strDataSource[x,Datasource.TABLE].Trim() + "_work_temp";

								oAdoCurrent.m_strSQL = "SELECT COUNT(*) FROM (SELECT TOP 1 * FROM " + oDs.m_strDataSource[x,Datasource.TABLE].Trim() + ")";
								
								//check if the projects version of the table has records
								if ((int)oAdoCurrent.getRecordCount(oAdoCurrent.m_OleDbConnection,oAdoCurrent.m_strSQL, oDs.m_strDataSource[x,Datasource.TABLE].Trim()) > 0)
								{

									bool bModified=false;

									if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,strEmptyTable))
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE " + strEmptyTable);

									//create a temp copy with no records (empty table) of the projects version of the table
									oAdoCurrent.m_strSQL = "SELECT TOP 1 * INTO " + strEmptyTable + " FROM " + oDs.m_strDataSource[x,Datasource.TABLE].Trim();
									oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,oAdoCurrent.m_strSQL);
								    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DELETE FROM " + strEmptyTable);
									

									//compare the application version of the table with the 
									//projects version of the table and 
									//modify the projects version of the table structure if needs be.
									bModified = oAdoCurrent.ReconcileTableColumns(oAdoCurrent.m_OleDbConnection,
													strEmptyTable,
													oConn,strTempTableName);
									if (bModified)
									{
										
										//insert all the records into our new empty table
										oAdoCurrent.m_strSQL = "INSERT INTO " + strEmptyTable + " SELECT * FROM " + oDs.m_strDataSource[x,Datasource.TABLE].Trim();
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,oAdoCurrent.m_strSQL);
										//drop the current table
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE " + oDs.m_strDataSource[x,Datasource.TABLE].Trim());
										//create the table structure giving it the current table name
										oAdoCurrent.m_strSQL = "SELECT TOP 1 * INTO " + oDs.m_strDataSource[x,Datasource.TABLE].Trim() + " FROM " + strEmptyTable;
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,oAdoCurrent.m_strSQL);
										//delete the 1 record in the current table
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DELETE FROM " + oDs.m_strDataSource[x,Datasource.TABLE].Trim());
                                        //insert all the records from the new empty table to the current table
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"INSERT INTO " + oDs.m_strDataSource[x,Datasource.TABLE].Trim() + " SELECT *  FROM " + strEmptyTable.Trim());
										//drop the empty table
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE " + strEmptyTable);
                                        
                                        if (oAdoCurrent.ColumnExist(oAdo.m_OleDbConnection,oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),"RX") &&
                                            oAdoCurrent.ColumnExist(oAdo.m_OleDbConnection,oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),"RXPACKAGE") &&
                                            oAdoCurrent.ColumnExist(oAdo.m_OleDbConnection,oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),"RXCYCLE"))
                                              this.ConvertRxAndRxPackageAndRxCycle(oAdoCurrent,oAdoCurrent.m_OleDbConnection,oDs.m_strDataSource[x,Datasource.TABLETYPE].Trim());

                                        //create the indexes for the empty table structure
                                        if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, strEmptyTable))
                                        {
                                            oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                                oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(), strEmptyTable.Trim());
                                        }
                                        else
                                        {
                                            if (oAdoCurrent.TableExist(oConn, strEmptyTable))
                                            {
                                                oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oConn,
                                                    oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(), strEmptyTable.Trim());
                                            }
                                        }

									}
									else
									{
										oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE " + strEmptyTable);
									}

								

								}
								else
								{
									if (oAdoCurrent.ReconcileTableColumns(oAdoCurrent.m_OleDbConnection,
										oDs.m_strDataSource[x,Datasource.TABLE],
										oConn,strTempTableName))
									{
									}
								}
								
							}
						}
					}
				}
			}
			if (oAdoCurrent!=null) oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
			oAdo.CloseConnection(oAdo.m_OleDbConnection);
			if (oConn != null)
			{
			   oAdo.CloseConnection(oConn);
			}
			if (oDao!=null)
			{
				oDao.m_DaoWorkspace.Close();
				oDao=null; 
			}
			oAdoCurrent=null;
			oAdo=null;
		}

		private void CheckProjectReferenceDatasourceTables()
		{

			int x,y;
			FIA_Biosum_Manager.Datasource oDs = new Datasource();
			oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
			oDs.m_strDataSourceTableName = "datasource";
			oDs.m_strScenarioId="";
			oDs.LoadTableColumnNamesAndDataTypes=false;
			oDs.LoadTableRecordCount=false;
			oDs.populate_datasource_array();

			ado_data_access oAdo=new ado_data_access();
			ado_data_access oAdoCurrent=null;

			dao_data_access oDao= new dao_data_access();
			string strCurrDbFile="";
			string strDbFile="";
			string strTempDbFile="";
			string strTempTableName="";
			System.Data.OleDb.OleDbConnection oConn=null;


			//copy ref_master.mdb to project db directory if file does not exist
			if (System.IO.File.Exists(ReferenceProjectDirectory.Trim() + "\\db\\ref_master.mdb")==false)
				 System.IO.File.Copy(frmMain.g_oEnv.strAppDir + "\\db\\ref_master.mdb",this.ReferenceProjectDirectory + "\\db\\ref_master.mdb",true);


			
			

			//open the project db file
			oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" + 
				frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile,"",""));

			oAdo.m_strSQL = "SELECT * FROM datasource " + 
				            "WHERE table_type IS NOT NULL AND " + 
                                  "UCASE(TRIM(table_type)) " + 
                                  "IN ('OWNER GROUPS'," + 
                                      "'TREE SPECIES'," + 
                                      "'FVS TREE SPECIES'," + 
                                      "'FIADB FVS VARIANT'," + 
                                      "'INVENTORIES'," + 
                                      "'TREATMENT PRESCRIPTION CATEGORIES'," + 
                                      "'TREATMENT PRESCRIPTION SUBCATEGORIES'," + 
                                      "'HARVEST METHODS'," + 
                                      "'FVS WESTERN TREE SPECIES TRANSLATOR'," +
                                      "'FVS EASTERN TREE SPECIES TRANSLATOR'" + 
                                      ") ORDER BY path,file";
			oAdo.SqlQueryReader(oAdo.m_OleDbConnection,oAdo.m_strSQL);

			if (oAdo.m_OleDbDataReader.HasRows)
			{
				while (oAdo.m_OleDbDataReader.Read())
				{
					if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value &&
						oAdo.m_OleDbDataReader["file"] != System.DBNull.Value)
					{
						x=oDs.getDataSourceTableNameRow(oAdo.m_OleDbDataReader["table_type"].ToString().Trim());


						
						strDbFile = oAdo.m_OleDbDataReader["path"].ToString().Trim() + "\\" + 
							oAdo.m_OleDbDataReader["file"].ToString().Trim();

						if (strDbFile.Trim().ToUpper() != strCurrDbFile.Trim().ToUpper())
						{
							if (oAdoCurrent!=null) 
							{  
								if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"owner_groups_temp"))
									 oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE owner_groups_temp");
								if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"tree_species_temp"))
									oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE tree_species_temp");
								if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"fvs_tree_species_temp"))
									oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE fvs_tree_species_temp");
								if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"fiadb_fvs_variant_temp"))
									oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE fiadb_fvs_variant_temp");
								if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"inventories_temp"))
									oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE inventories_temp");
                                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "harvest_methods_temp"))
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE harvest_methods_temp");
                                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "fvs_rx_category_temp"))
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE fvs_rx_category_temp");
                                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "fvs_rx_subcategory_temp"))
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE fvs_rx_subcategory_temp");
                                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "FVS_WesternTreeSpeciesTranslator_temp"))
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE FVS_WesternTreeSpeciesTranslator_temp");
                                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "FVS_EasternTreeSpeciesTranslator_temp"))
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE FVS_EasternTreeSpeciesTranslator_temp");

								oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
							}
							else 
							{
								strTempDbFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");
								oDao.CreateMDB(strTempDbFile);
								System.IO.File.Copy(frmMain.g_oEnv.strAppDir.Trim() + "\\db\\ref_master.mdb",strTempDbFile,true);

								oAdoCurrent = new ado_data_access();
							}

							//make sure the db file exists
							if (System.IO.File.Exists(strDbFile)==false)
							{
								oDao.CreateMDB(strDbFile);
							}
							
							//create a link to all the pre-populated reference tables
							oDao.CreateTableLink(strDbFile,"owner_groups_temp",strTempDbFile,"owner_groups",true);
							oDao.CreateTableLink(strDbFile,"tree_species_temp",strTempDbFile,"tree_species",true);
							oDao.CreateTableLink(strDbFile,"fvs_tree_species_temp",strTempDbFile,"fvs_tree_species",true);
							oDao.CreateTableLink(strDbFile,"fiadb_fvs_variant_temp",strTempDbFile,"fiadb_fvs_variant",true);
							oDao.CreateTableLink(strDbFile,"inventories_temp",strTempDbFile,"inventories",true);
                            oDao.CreateTableLink(strDbFile, "harvest_methods_temp", strTempDbFile, "harvest_methods", true);
                            oDao.CreateTableLink(strDbFile, "fvs_rx_category_temp", strTempDbFile, "fvs_rx_category",true);
                            oDao.CreateTableLink(strDbFile, "fvs_rx_subcategory_temp", strTempDbFile, "fvs_rx_subcategory",true);
                            oDao.CreateTableLink(strDbFile, "FVS_WesternTreeSpeciesTranslator_temp", strTempDbFile, "FVS_WesternTreeSpeciesTranslator", true);
                            oDao.CreateTableLink(strDbFile, "FVS_EasternTreeSpeciesTranslator_temp", strTempDbFile, "FVS_EasternTreeSpeciesTranslator", true);
							oAdoCurrent.OpenConnection(oAdo.getMDBConnString(strDbFile,"",""));
							strCurrDbFile=strDbFile;

						}
                        if (oDs.m_strDataSource[x, Datasource.TABLESTATUS] == "NF") //NF=table not found
                        {
                            //create the table
                            switch (oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim().ToUpper())
                            {

                                case "OWNER GROUPS":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM owner_groups_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "TREE SPECIES":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM tree_species_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "FVS TREE SPECIES":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM fvs_tree_species_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "FIADB FVS VARIANT":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM fiadb_fvs_variant_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "INVENTORIES":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM inventories_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                //version 5 additions
                                case "TREATMENT PRESCRIPTION CATEGORIES":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM fvs_rx_category_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "TREATMENT PRESCRIPTION SUBCATEGORIES":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM fvs_rx_subcategory_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "HARVEST METHODS":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM harvest_methods_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "FVS WESTERN TREE SPECIES TRANSLATOR":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM FVS_WesternTreeSpeciesTranslator_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "FVS EASTERN TREE SPECIES TRANSLATOR":
                                    oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM FVS_EasternTreeSpeciesTranslator_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;

                            }
                        }
                        else
                        {
                            //update columns and data
                            strTempTableName = "";
                            strInsertSql = "";
                            strUpdateSql = "";
                            string[] strColumnsArray = null;
                            string strColumnsList = "";
                            switch (oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim().ToUpper())
                            {
                                case "OWNER GROUPS":
                                    oAdoCurrent.m_strSQL = "DROP TABLE " + oDs.m_strDataSource[x, Datasource.TABLE];
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    strInsertSql = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE] + " FROM owner_groups_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                            oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                            oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "INVENTORIES":
                                    oAdoCurrent.m_strSQL = "DROP TABLE " + oDs.m_strDataSource[x, Datasource.TABLE];
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    strInsertSql = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE] + " FROM inventories_temp";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                        oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                        oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                                    break;
                                case "TREE SPECIES":
                                    strTempTableName = "tree_species_temp";


                                    //insert any new tree species records
                                    strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        if (strColumnsArray[y].Trim().ToUpper() != "ID")
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strInsertSql = "INSERT INTO " +
                                                        oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                                    "SELECT " + strColumnsList + " " +
                                                    "FROM tree_species_temp a " +
                                                    "WHERE  a.fvs_variant IS NOT NULL AND " +
                                                           "LEN(TRIM(a.fvs_variant)) > 0 AND " +
                                                           "NOT EXISTS " +
                                                          "(SELECT b.spcd,b.fvs_variant  " +
                                                           "FROM " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " b " +
                                                           "WHERE a.spcd=b.spcd AND a.fvs_variant=b.fvs_variant)";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    //update any null values
                                    strColumnsList = "";
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        if (strColumnsArray[y].Trim().ToUpper() != "ID")
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() +
                                                "=IIF(a." + strColumnsArray[y].Trim() + " IS NULL," +
                                                     "b." + strColumnsArray[y].Trim() + "," +
                                                     "a." + strColumnsArray[y].Trim() + "),";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strUpdateSql = "UPDATE " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " a " +
                                                   "INNER JOIN tree_species_temp b " +
                                                   "ON a.spcd=b.spcd AND a.fvs_variant=b.fvs_variant " +
                                                   "SET " + strColumnsList;
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strUpdateSql);

                                    break;
                                case "FVS TREE SPECIES":
                                    strTempTableName = "fvs_tree_species_temp";

                                    //add/update columns
                                    //oAdoCurrent.ReconcileTableColumns(oAdoCurrent.m_OleDbConnection,
                                    //	oDs.m_strDataSource[x,Datasource.TABLE],oAdoCurrent.m_OleDbConnection,
                                    //	strTempTableName);

                                    //insert any new tree species records
                                    strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        if (strColumnsArray[y].Trim().ToUpper() != "ID")
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strInsertSql = "INSERT INTO " +
                                        oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                        "SELECT " + strColumnsList + " " +
                                        "FROM fvs_tree_species_temp a " +
                                        "WHERE  a.fvs_variant IS NOT NULL AND " +
                                        "LEN(TRIM(a.fvs_variant)) > 0 AND " +
                                        "NOT EXISTS " +
                                        "(SELECT b.spcd,b.fvs_variant  " +
                                        "FROM " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " b " +
                                        "WHERE a.spcd=b.spcd AND a.fvs_variant=b.fvs_variant)";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    //update any null values
                                    strColumnsList = "";
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        if (strColumnsArray[y].Trim().ToUpper() != "ID")
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() +
                                                "=IIF(a." + strColumnsArray[y].Trim() + " IS NULL," +
                                                "b." + strColumnsArray[y].Trim() + "," +
                                                "a." + strColumnsArray[y].Trim() + "),";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strUpdateSql = "UPDATE " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " a " +
                                        "INNER JOIN fvs_tree_species_temp b " +
                                        "ON a.spcd=b.spcd AND a.fvs_variant=b.fvs_variant " +
                                        "SET " + strColumnsList;
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strUpdateSql);
                                    break;
                                case "FIADB FVS VARIANT":
                                    strTempTableName = "fiadb_fvs_variant_temp";

                                    //insert any new tree species records
                                    strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        if (strColumnsArray[y].Trim().ToUpper() != "ID")
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strInsertSql = "INSERT INTO " +
                                        oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                        "SELECT " + strColumnsList + " " +
                                        "FROM " + strTempTableName + " a " +
                                        "WHERE  a.fvs_variant IS NOT NULL AND " +
                                        "LEN(TRIM(a.fvs_variant)) > 0 AND " +
                                        "a.statecd IS NOT NULL AND " +
                                        "a.countycd IS NOT NULL AND " +
                                        "a.plot IS NOT NULL AND " +
                                        "NOT EXISTS " +
                                        "(SELECT b.statecd,b.countycd,b.plot,b.fvs_variant  " +
                                        "FROM " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " b " +
                                        "WHERE a.statecd=b.statecd AND " +
                                              "a.countycd=b.countycd AND " +
                                              "a.plot=b.plot AND " +
                                              "a.fvs_variant=b.fvs_variant)";
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);

                                    //update any null values
                                    strColumnsList = "";
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        if (strColumnsArray[y].Trim().ToUpper() != "ID")
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() +
                                                "=IIF(a." + strColumnsArray[y].Trim() + " IS NULL," +
                                                "b." + strColumnsArray[y].Trim() + "," +
                                                "a." + strColumnsArray[y].Trim() + "),";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strUpdateSql = "UPDATE " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " a " +
                                        "INNER JOIN " + strTempTableName + " b " +
                                        "ON a.statecd=b.statecd AND " +
                                           "a.countycd=b.countycd AND " +
                                           "a.plot=b.plot AND " +
                                           "a.fvs_variant=b.fvs_variant " +
                                        "SET " + strColumnsList;
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strUpdateSql);
                                    break;
                                //version 5 additions
                                case "TREATMENT PRESCRIPTION CATEGORIES":
                                    if ((int)oAdoCurrent.getRecordCount(oAdoCurrent.m_OleDbConnection, "SELECT COUNT(*) FROM " + oDs.m_strDataSource[x, Datasource.TABLE].Trim(), "TEMP") == 0)
                                    {
                                        strTempTableName = "fvs_rx_category_temp";
                                        strColumnsList="";
                                        //insert any new tree species records
                                        strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                        for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                        {
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                        }
                                        strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                        strInsertSql = "INSERT INTO " +
                                            oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                            "SELECT " + strColumnsList + " " +
                                            "FROM " + strTempTableName + " a ";
                                           
                                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                          
                                        
                                    }
                                    break;
                                case "TREATMENT PRESCRIPTION SUBCATEGORIES":
                                     if ((int)oAdoCurrent.getRecordCount(oAdoCurrent.m_OleDbConnection, "SELECT COUNT(*) FROM " + oDs.m_strDataSource[x, Datasource.TABLE].Trim(), "TEMP") == 0)
                                    {
                                        strTempTableName = "fvs_rx_subcategory_temp";
                                        strColumnsList="";
                                        //insert any new tree species records
                                        strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                        for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                        {
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                        }
                                        strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                        strInsertSql = "INSERT INTO " +
                                            oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                            "SELECT " + strColumnsList + " " +
                                            "FROM " + strTempTableName + " a ";
                                           
                                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                          
                                            
                                    
                                    }
                                    break;
                                case "HARVEST METHODS":
                                        oAdoCurrent.m_strSQL = "DROP TABLE " + oDs.m_strDataSource[x, Datasource.TABLE];
                                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                        frmMain.g_oTables.m_oReference.CreateHarvestMethodsTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, oDs.m_strDataSource[x, Datasource.TABLE]);
                                        strTempTableName = "harvest_methods_temp";
                                        strColumnsList="";
                                        //insert any new tree species records
                                        strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                        for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                        {
                                            strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                        }
                                        strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                        strInsertSql = "INSERT INTO " +
                                            oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                            "SELECT " + strColumnsList + " " +
                                            "FROM " + strTempTableName + " a ";
                                           
                                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    break;
                                case "FVS WESTERN TREE SPECIES TRANSLATOR":
                                    oAdoCurrent.m_strSQL = "DROP TABLE " + oDs.m_strDataSource[x, Datasource.TABLE];
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    frmMain.g_oTables.m_oReference.CreateFVSWesternSpeciesTranslatorTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, oDs.m_strDataSource[x, Datasource.TABLE]);
                                    strTempTableName = "FVS_WesternTreeSpeciesTranslator_temp";
                                    strColumnsList = "";
                                    //insert any new tree species records
                                    strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strInsertSql = "INSERT INTO " +
                                        oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                        "SELECT " + strColumnsList + " " +
                                        "FROM " + strTempTableName + " a ";

                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    break;
                                case "FVS EASTERN TREE SPECIES TRANSLATOR":
                                    oAdoCurrent.m_strSQL = "DROP TABLE " + oDs.m_strDataSource[x, Datasource.TABLE];
                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                    frmMain.g_oTables.m_oReference.CreateFVSEasternSpeciesTranslatorTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, oDs.m_strDataSource[x, Datasource.TABLE]);
                                    strTempTableName = "FVS_EasternTreeSpeciesTranslator_temp";
                                    strColumnsList = "";
                                    //insert any new tree species records
                                    strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
                                    for (y = 0; y <= strColumnsArray.Length - 1; y++)
                                    {
                                        strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
                                    }
                                    strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
                                    strInsertSql = "INSERT INTO " +
                                        oDs.m_strDataSource[x, Datasource.TABLE] + " " +
                                        "SELECT " + strColumnsList + " " +
                                        "FROM " + strTempTableName + " a ";

                                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
                                    break;
                            }
                        }
					}
				}
			}
			if (oAdoCurrent!=null) 
			{
				if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"owner_groups_temp"))
					oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE owner_groups_temp");
				if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"tree_species_temp"))
					oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE tree_species_temp");
				if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"fvs_tree_species_temp"))
					oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE fvs_tree_species_temp");
				if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"fiadb_fvs_variant_temp"))
					oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE fiadb_fvs_variant_temp");
				if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"inventories_temp"))
					oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE inventories_temp");
                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "harvest_methods_temp"))
                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE harvest_methods_temp");
                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "fvs_rx_category_temp"))
                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE fvs_rx_category_temp");
                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "fvs_rx_subcategory_temp"))
                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE fvs_rx_subcategory_temp");
                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "FVS_WesternTreeSpeciesTranslator_temp"))
                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE FVS_WesternTreeSpeciesTranslator_temp");
                if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "FVS_EasternTreeSpeciesTranslator_temp"))
                    oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE FVS_EasternTreeSpeciesTranslator_temp");
				oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
			}
			oAdo.CloseConnection(oAdo.m_OleDbConnection);
			if (oConn != null)
			{
				oAdo.CloseConnection(oConn);
			}
			if (oDao!=null)
			{
				oDao.m_DaoWorkspace.Close();
				oDao=null; 
			}
			oAdoCurrent=null;
			oAdo=null;
            strCurrDbFile = "";
            strDbFile = "";
            strTempDbFile = "";
            strTempTableName = "";
            oConn = null;
            //
            //fvs commands table
            //
            

            
            oAdo = new ado_data_access();
            oDao = new dao_data_access();
            string strNewDbFile = "";
            string strOldDbFile = "";

            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));

            oAdo.m_strSQL = "SELECT * FROM datasource " +
                            "WHERE table_type IS NOT NULL AND " +
                                  "UCASE(TRIM(table_type)) " +
                                  "IN ('FVS COMMANDS') ORDER BY path,file";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value &&
                        oAdo.m_OleDbDataReader["file"] != System.DBNull.Value)
                    {
                        x = oDs.getDataSourceTableNameRow(oAdo.m_OleDbDataReader["table_type"].ToString().Trim());



                        strOldDbFile = oAdo.m_OleDbDataReader["path"].ToString().Trim() + "\\" +
                            oAdo.m_OleDbDataReader["file"].ToString().Trim();

                        if (System.IO.File.Exists(strOldDbFile))
                        {
                            oAdoCurrent = new ado_data_access();
                            //
                            //get all table names in the old file
                            //
                            string[] strTablesArray = null;
                            oDao.getTableNames(strOldDbFile, ref strTablesArray);
                            //
                            //get the new file name
                            //
                            strNewDbFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
                            //
                            //create the new file and open it
                            //
                            oDao.CreateMDB(strNewDbFile);
                            oDao.OpenDb(strNewDbFile);
                            //
                            //create a link in the new file to the new table
                            //
                            oDao.CreateTableLink(oDao.m_DaoDatabase, "fvs_commands_temp", frmMain.g_oEnv.strAppDir + "\\db\\ref_fvscommands.mdb", "fvs_commands");
                            //
                            //create a link to all tables from the old file to the new file
                            //
                            if (strTablesArray != null)
                            {
                               
                                for (y = 0; y <= strTablesArray.Length - 1; y++)
                                {
                                    if (strTablesArray[y] == null) break;
                                    if (strTablesArray[y].Trim().Length > 0)
                                    {
                                        //do not link to the table we are replacing
                                        if (strTablesArray[y].Trim().ToUpper() !=
                                            oDs.m_strDataSource[x, Datasource.TABLE].Trim().ToUpper())
                                        {
                                            oDao.CreateTableLink(oDao.m_DaoDatabase, strTablesArray[y].Trim() + "_temp", strOldDbFile, strTablesArray[y].Trim());
                                        }


                                    }
                                }
                                oDao.m_DaoWorkspace.Close();
                                oDao = null;
                                //open oledb connection to new file
                                oAdoCurrent.OpenConnection(oAdoCurrent.getMDBConnString(strNewDbFile,"",""));
                                for (y = 0; y <= strTablesArray.Length - 1; y++)
                                {
                                    if (strTablesArray[y] == null) break;
                                    //import linked tables
                                    if (strTablesArray[y].Trim().ToUpper() == oDs.m_strDataSource[x, Datasource.TABLE].Trim().ToUpper())
                                    {
                                        
                                    }
                                    else
                                    {
                                        oAdoCurrent.m_strSQL = "SELECT * INTO " + strTablesArray[y].Trim() + " " +
                                                               "FROM " + strTablesArray[y].Trim() + "_temp";
                                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                                        //drop the linked table
                                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE " + strTablesArray[y].Trim() + "_temp");
                                    }
                                }
                               
                            }
                            //
                            //import linked fvs_commands table
                            //
                            if (oAdoCurrent.m_OleDbConnection.State == System.Data.ConnectionState.Closed)
                                oAdoCurrent.OpenConnection(oAdoCurrent.getMDBConnString(strNewDbFile, "", ""));
                            oAdoCurrent.m_strSQL = "SELECT * " + 
                                                   "INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " " + 
                                                   "FROM fvs_commands_temp";
                            oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
                            oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
                                oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
                                oDs.m_strDataSource[x, Datasource.TABLE].Trim());
                            oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE fvs_commands_temp");
                            oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
                            //
                            //delete old file and copy new file
                            //
                            System.IO.File.Delete(strOldDbFile);
                            System.IO.File.Copy(strNewDbFile,strOldDbFile,true);

                        }
                        else
                        {
                            //copy ref_master.mdb to project db directory if file does not exist
                            if (System.IO.File.Exists(ReferenceProjectDirectory.Trim() + "\\db\\ref_fvscommands.mdb") == false)
                                System.IO.File.Copy(frmMain.g_oEnv.strAppDir + "\\db\\ref_fvscommands.mdb", this.ReferenceProjectDirectory + "\\db\\ref_fvscommands.mdb", true);
                        }

                        
                    }
                }
            }
            
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            if (oConn != null)
            {
                oAdo.CloseConnection(oConn);
            }
            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
            oAdoCurrent = null;
            oAdo = null;
		}



        private void UpdateFVSPlotVariantAssignmentsTable()
        {

            int x;
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();

            ado_data_access oAdo = new ado_data_access();
            ado_data_access oAdoCurrent = null;

            dao_data_access oDao = new dao_data_access();
            string strCurrDbFile = "";
            string strCurrTableName="";

            string strSourceDbFile="";
            string strSourceTableName = "";


            frmMain.g_sbpInfo.Text = "Version Update: Update FIADB FVS Plot Variant Reference Table...Stand by";

            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));

            //get the MDB file name and table name
            oAdo.m_strSQL = "SELECT * FROM datasource " +
                            "WHERE table_type IS NOT NULL AND " +
                                  "UCASE(TRIM(table_type)) " +
                                  "IN ('FIADB FVS VARIANT') " + 
                            "ORDER BY path,file";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value &&
                        oAdo.m_OleDbDataReader["file"] != System.DBNull.Value)
                    {



                        x = oDs.getDataSourceTableNameRow(oAdo.m_OleDbDataReader["table_type"].ToString().Trim());

                        strCurrDbFile = oAdo.m_OleDbDataReader["path"].ToString().Trim() + "\\" +
                                        oAdo.m_OleDbDataReader["file"].ToString().Trim();

                        strCurrTableName = oDs.m_strDataSource[x, Datasource.TABLE].Trim();

                        oDao = new dao_data_access();
                        //
                        //CREATE THE MDB IF NOT EXIST
                        //This table no longer lives in ref_master.mdb; In shared appdata directory
                        //
                        //if (oDs.m_strDataSource[x, Datasource.FILESTATUS] == "NF") //NF=table not found
                        //{
                        //    oDao.CreateMDB(strCurrDbFile);
                        //}
                        //strSourceDbFile = frmMain.g_oEnv.strAppDir.Trim() + "\\" + Tables.Reference.DefaultFiadbFVSVariantTableDbFile;
                        //strSourceTableName = Tables.Reference.DefaultFiadbFVSVariantTableName;
                        //
                        //
                        //DELETE ANY OLD TABLES
                        //
                        if (oDao.TableExists(strCurrDbFile, strCurrTableName + "_temp"))
                            oDao.DeleteTableFromMDB(strCurrDbFile, strCurrTableName + "_temp");
                        if (oDao.TableExists(strCurrDbFile, strCurrTableName + "_temp2"))
                            oDao.DeleteTableFromMDB(strCurrDbFile, strCurrTableName + "_temp2");
                        //
                        //CREATE TABLE LINK
                        //
                        oDao.CreateTableLink(strCurrDbFile, strCurrTableName + "_temp", strSourceDbFile, strSourceTableName);
                        oDao.m_DaoWorkspace.Close();
                        oDao = null;
                        System.Threading.Thread.Sleep(4000);
                        //
                        //INSTANTIATE NEW ADO OBJECT
                        //
                        oAdoCurrent = new ado_data_access();
                        //
                        //OPEN ADO CONNECTION
                        //
                        oAdoCurrent.OpenConnection(oAdoCurrent.getMDBConnString(strCurrDbFile, "", ""));
                        //
                        //COPY RECORDS FROM LINK TABLE
                        //
                        oAdo.m_strSQL = "SELECT * INTO " + strCurrTableName + "_temp2 FROM " + strCurrTableName + "_temp";
                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdo.m_strSQL);
                        //
                        //DROP THE LINK
                        //
                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE " + strCurrTableName + "_temp");
                        //
                        //DROP THE OLD TABLE
                        //
                        if (oDs.m_strDataSource[x, Datasource.TABLESTATUS] == "F")
                            oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE " + strCurrTableName);
                        //
                        //COPY TEMP TABLE TO PRODUCTION
                        // 
                        oAdo.m_strSQL = "SELECT * INTO " + strCurrTableName + " FROM " + strCurrTableName + "_temp2";
                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdo.m_strSQL);
                        //
                        //DROP TEMP TABLE
                        //
                        oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE " + strCurrTableName + "_temp2");
                        //
                        //CLOSE ADOCURRENT CONNECTION
                        //
                        oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
                        oAdoCurrent = null;
                        break;
                    }
                }
            }
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
        }

		/// <summary>
		/// Check and update the scenario datasource table with the latest version of table type entries
		/// </summary>
		private void CheckCoreScenarioDatasourceTableRecords()
		{
		    
			int x;
			int y;
			bool bCore;

			FIA_Biosum_Manager.Datasource oDsScenario = new Datasource();
			FIA_Biosum_Manager.Datasource oDsProject = new Datasource();
			ado_data_access oAdoScenario=new ado_data_access();
			ado_data_access oAdoProject = new ado_data_access();
			

			oDsScenario.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\core\\db\\scenario_core_rule_definitions.mdb";
			
			string strDbFile="";
			string strSQL="";
			string strScenarioList="";
			string[] strScenarioArray=null;

			//open the core scenario db file
			oAdoScenario.OpenConnection(oAdoScenario.getMDBConnString(oDsScenario.m_strDataSourceMDBFile,"",""));
            //remove obsolete table types
            oAdoScenario.m_strSQL = "DELETE FROM scenario_datasource WHERE " + 
                "TRIM(UCASE(table_type)) IN (" +
                "'FIRE AND FUEL EFFECTS'," +
                "'TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS'," +
                "'HARVEST COSTS')";
            oAdoScenario.SqlNonQuery(oAdoScenario.m_OleDbConnection, oAdoScenario.m_strSQL);
			//get all the scenarios
			oAdoScenario.SqlQueryReader(oAdoScenario.m_OleDbConnection,"SELECT scenario_id FROM scenario");
			if (oAdoScenario.m_OleDbDataReader.HasRows)
			{
				while (oAdoScenario.m_OleDbDataReader.Read())
				{
					if (oAdoScenario.m_OleDbDataReader["scenario_id"] != System.DBNull.Value)
					{
						strScenarioList = strScenarioList + oAdoScenario.m_OleDbDataReader["scenario_id"].ToString().Trim() + ",";
					}
				}
				
			}
			oAdoScenario.m_OleDbDataReader.Close();

			if (strScenarioList.Trim().Length > 0)
			{
				strScenarioList = strScenarioList.Substring(0,strScenarioList.Trim().Length - 1);
				strScenarioArray = frmMain.g_oUtils.ConvertListToArray(strScenarioList,",");
				//open the project db file
				oAdoProject.OpenConnection(oAdoProject.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" + 
						frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile,"",""));

				//cycle through each scenario datasource to ensure all the latest datasources exist
				for (x=0;x<=strScenarioArray.Length-1;x++)
				{
					//make sure all the latest datasources are in the scenario datasource
					//get the project datasources
					oAdoProject.SqlQueryReader(oAdoProject.m_OleDbConnection,"select * from datasource");
					if (oAdoProject.m_intError==0)
					{
						//load the scenario datasource
						oDsScenario.m_strDataSourceTableName = "scenario_datasource";
						oDsScenario.m_strScenarioId=strScenarioArray[x].Trim();
						oDsScenario.LoadTableColumnNamesAndDataTypes=false;
						oDsScenario.LoadTableRecordCount=false;
						oDsScenario.populate_datasource_array();
						while (oAdoProject.m_OleDbDataReader.Read())
						{
							bCore=false;
							switch (oAdoProject.m_OleDbDataReader["table_type"].ToString().Trim().ToUpper())
							{
								case "PLOT":
									bCore=true;
									break;
								case "CONDITION":
									bCore = true;
									break;
								//case "HARVEST COSTS":
								//	bCore = true;
								//	break;
								case "TREATMENT PRESCRIPTIONS":
									bCore = true;
									break;
								//case "TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS":
								//	bCore = true;
								//	break;
								case "TRAVEL TIMES":
									bCore = true;
									break;
								case "PROCESSING SITES":
									bCore = true;
									break;
								//case "TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES":
								//	bCore = true;
								//	break;
								case "PLOT AND CONDITION RECORD AUDIT":
									bCore = true;
									break;
								case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
									bCore = true;
									break;
                                case "TREATMENT PACKAGES":
                                    bCore = true;
                                    break;

								default:
									break;
							}
							if (bCore == true)
							{
								//see if the project datasource table type exists in the scenario datasource table
								y=oDsScenario.getValidTableNameRow(oAdoProject.m_OleDbDataReader["table_type"].ToString().Trim().ToUpper());
								if (y==-1)
								{
									strSQL = "INSERT INTO scenario_datasource (scenario_id,table_type,Path,file,table_name) VALUES " + "('" + strScenarioArray[x].Trim() + "'," + 
										"'" + oAdoProject.m_OleDbDataReader["table_type"].ToString().Trim() + "'," + 
										"'" + oAdoProject.m_OleDbDataReader["path"].ToString().Trim() + "'," + 
										"'" + oAdoProject.m_OleDbDataReader["file"].ToString().Trim() + "'," +  
										"'" + oAdoProject.m_OleDbDataReader["table_name"].ToString().Trim() + "');";
									oAdoScenario.SqlNonQuery(oAdoScenario.m_OleDbConnection,strSQL);
								}
							}

						}
						oAdoProject.m_OleDbDataReader.Close();
					}
				}
				oAdoProject.CloseConnection(oAdoProject.m_OleDbConnection);
			}
			oAdoScenario.CloseConnection(oAdoScenario.m_OleDbConnection);
		}
        /// <summary>
        /// Check and update the scenario datasource table with the latest version of table type entries
        /// </summary>
        private void CheckProcessorScenarioDatasourceTableRecords()
        {

            int x;
            int y;
            bool bCore;

            FIA_Biosum_Manager.Datasource oDsScenario = new Datasource();
            FIA_Biosum_Manager.Datasource oDsProject = new Datasource();
            ado_data_access oAdoScenario = new ado_data_access();
            ado_data_access oAdoProject = new ado_data_access();


            oDsScenario.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\Processor\\db\\scenario_processor_rule_definitions.mdb";
            if (System.IO.File.Exists(oDsScenario.m_strDataSourceMDBFile) == false)
            {
                frmMain.g_oFrmMain.frmProject.uc_project1.CreateProcessorScenarioRuleDefinitionDbAndTables(oDsScenario.m_strDataSourceMDBFile);

            }

            string strDbFile = "";
            string strSQL = "";
            string strScenarioList = "";
            string[] strScenarioArray = null;

            //open the core scenario db file
            oAdoScenario.OpenConnection(oAdoScenario.getMDBConnString(oDsScenario.m_strDataSourceMDBFile, "", ""));
            
            //remove obsolete table types
            oAdoScenario.m_strSQL = "DELETE FROM scenario_datasource WHERE " +
                "TRIM(UCASE(table_type)) IN (" +
                "'FIRE AND FUEL EFFECTS'," +
                "'TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS'," +
                "'HARVEST COSTS'," +
                "'PLOT AND CONDITION RECORD AUDIT'," +
                "'PLOT, CONDITION AND TREATMENT RECORD AUDIT'," + 
                "'FVS TREE LIST FOR PROCESSOR'," +
                "'ADDITIONAL HARVEST COSTS'," +
                "'TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES')";
            oAdoScenario.SqlNonQuery(oAdoScenario.m_OleDbConnection, oAdoScenario.m_strSQL);
            //get all the scenarios
            oAdoScenario.SqlQueryReader(oAdoScenario.m_OleDbConnection, "SELECT scenario_id FROM scenario");
            if (oAdoScenario.m_OleDbDataReader.HasRows)
            {
                while (oAdoScenario.m_OleDbDataReader.Read())
                {
                    if (oAdoScenario.m_OleDbDataReader["scenario_id"] != System.DBNull.Value)
                    {
                        strScenarioList = strScenarioList + oAdoScenario.m_OleDbDataReader["scenario_id"].ToString().Trim() + ",";
                    }
                }

            }
            oAdoScenario.m_OleDbDataReader.Close();

            if (strScenarioList.Trim().Length > 0)
            {
                strScenarioList = strScenarioList.Substring(0, strScenarioList.Trim().Length - 1);
                strScenarioArray = frmMain.g_oUtils.ConvertListToArray(strScenarioList, ",");
                //open the project db file
                oAdoProject.OpenConnection(oAdoProject.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                        frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));

                //cycle through each scenario datasource to ensure all the latest datasources exist
                for (x = 0; x <= strScenarioArray.Length - 1; x++)
                {
                    //make sure all the latest datasources are in the scenario datasource
                    //get the project datasources
                    oAdoProject.SqlQueryReader(oAdoProject.m_OleDbConnection, "select * from datasource");
                    if (oAdoProject.m_intError == 0)
                    {
                        //load the scenario datasource
                        oDsScenario.m_strDataSourceTableName = "scenario_datasource";
                        oDsScenario.m_strScenarioId = strScenarioArray[x].Trim();
                        oDsScenario.LoadTableColumnNamesAndDataTypes = false;
                        oDsScenario.LoadTableRecordCount = false;
                        oDsScenario.populate_datasource_array();
                        while (oAdoProject.m_OleDbDataReader.Read())
                        {
                            bCore = false;
                            switch (oAdoProject.m_OleDbDataReader["table_type"].ToString().Trim().ToUpper())
                            {
                                case "PLOT":
                                    bCore = true;
                                    break;
                                case "CONDITION":
                                    bCore = true;
                                    break;
                                //case "HARVEST COSTS":
                                //	bCore = true;
                                //	break;
                                case "TREATMENT PRESCRIPTIONS":
                                    bCore = true;
                                    break;
                                //case "TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS":
                                //	bCore = true;
                                //	break;
                                case "TRAVEL TIMES":
                                    bCore = true;
                                    break;
                                case "PROCESSING SITES":
                                    bCore = true;
                                    break;
                                //case "TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES":
                                //	bCore = true;
                                //	break;
                                //case "PLOT AND CONDITION RECORD AUDIT":
                                 //   bCore = true;
                                  //  break;
                                //case "PLOT, CONDITION AND TREATMENT RECORD AUDIT":
                                //    bCore = true;
                                //    break;
                                case "TREATMENT PACKAGES":
                                    bCore = true;
                                    break;
                                case "HARVEST METHODS":
                                    bCore = true;
                                    break;
                                //case "ADDITIONAL HARVEST COSTS":
                                //    bCore = true;
                                //    break;
                                case "TREATMENT PRESCRIPTIONS HARVEST COST COLUMNS":
                                    bCore = true;
                                    break;
                                default:
                                    break;
                            }
                            if (bCore == true)
                            {
                                //see if the project datasource table type exists in the scenario datasource table
                                y = oDsScenario.getValidTableNameRow(oAdoProject.m_OleDbDataReader["table_type"].ToString().Trim().ToUpper());
                                if (y == -1)
                                {
                                    strSQL = "INSERT INTO scenario_datasource (scenario_id,table_type,Path,file,table_name) VALUES " + "('" + strScenarioArray[x].Trim() + "'," +
                                        "'" + oAdoProject.m_OleDbDataReader["table_type"].ToString().Trim() + "'," +
                                        "'" + oAdoProject.m_OleDbDataReader["path"].ToString().Trim() + "'," +
                                        "'" + oAdoProject.m_OleDbDataReader["file"].ToString().Trim() + "'," +
                                        "'" + oAdoProject.m_OleDbDataReader["table_name"].ToString().Trim() + "');";
                                    oAdoScenario.SqlNonQuery(oAdoScenario.m_OleDbConnection, strSQL);
                                }
                            }

                        }
                        oAdoProject.m_OleDbDataReader.Close();
                    }
                }
                oAdoProject.CloseConnection(oAdoProject.m_OleDbConnection);
            }
            oAdoScenario.CloseConnection(oAdoScenario.m_OleDbConnection);
        }
        private void CheckRxValues()
        {
            int x,y,z;
            string strRxPackage="";
            string str="";
            string str2 = "";
            string strTable="";
            string strSourceFolder = "";
            string strDestFolder = "";

            ado_data_access oAdo = new ado_data_access();

          
            
            Queries oQueries = new Queries();
            oQueries.m_oFvs.LoadDatasource = true;
            oQueries.m_oFIAPlot.LoadDatasource = true;
            oQueries.m_oProcessor.LoadDatasource = true;
            oQueries.LoadDatasources(true);

           
            oAdo.OpenConnection(oAdo.getMDBConnString(oQueries.m_strTempDbFile, "", ""));
           
            //
            //GET RX VALUES
            //
            string[] strRxArray = frmMain.g_oUtils.ConvertListToArray(oAdo.CreateCommaDelimitedList(oAdo.m_OleDbConnection, "SELECT RX FROM " + oQueries.m_oFvs.m_strRxTable, ""), ",");
            //
            //GET VARIANT VALUES
            //
            string[] strVariantArray = frmMain.g_oUtils.ConvertListToArray(oAdo.CreateCommaDelimitedList(oAdo.m_OleDbConnection, "SELECT DISTINCT FVS_VARIANT FROM " + oQueries.m_oFIAPlot.m_strPlotTable, ""), ",");
            if (strRxArray != null && strRxArray.Length > 0)
            {
                if (strRxArray[0].Trim().Length == 1)
                {
                    //update rx 1 character alpha to 3  numeric character 701 to 725
                    
                    ConvertRx(oAdo, oAdo.m_OleDbConnection, oQueries.m_oFvs.m_strRxTable);
                    //update CATID to 700 series which is custom defined rx
                    oAdo.m_strSQL = "UPDATE " + oQueries.m_oFvs.m_strRxTable + " " +
                                    "SET CATID=4";
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                    //update fvs_tree table rx values
                    ConvertRx(oAdo, oAdo.m_OleDbConnection, oQueries.m_oFvs.m_strFvsTreeTable);
                    //create and assign an rxpackage to each rx
                    string[] strNewRxArray = frmMain.g_oUtils.ConvertListToArray(oAdo.CreateCommaDelimitedList(oAdo.m_OleDbConnection, "SELECT RX FROM " + oQueries.m_oFvs.m_strRxTable, ""), ",");
                    for (x = 0; x <= strNewRxArray.Length - 1; x++)
                    {
                        z=x+1;
                        if (z.ToString().Trim().Length == 1)
                            strRxPackage = "00" + z.ToString().Trim();
                        else if (z.ToString().Trim().Length == 2)
                            strRxPackage = "0" + z.ToString().Trim();
                        else
                            strRxPackage = z.ToString().Trim();
                          
                        oAdo.m_strSQL = "INSERT INTO " +
                                        oQueries.m_oFvs.m_strRxPackageTable + " " +
                                        "(rxpackage,rxcycle_length,simyear1_rx) VALUES " +
                                        "('" + strRxPackage + "','10','" + strNewRxArray[x].Trim() + "')";
                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                    }
                    if (strVariantArray != null && strVariantArray.Length > 0)
                    {
                        //copy any fvs input and output files 
                        for (x = 0; x <= strVariantArray.Length - 1; x++)
                        {
                            if (strVariantArray[x].Trim().Length > 0)
                            {
                                strSourceFolder = ReferenceProjectDirectory + "\\fvs\\db\\in\\" + strVariantArray[x].Trim();
                                strDestFolder = ReferenceProjectDirectory + "\\fvs\\data\\" + strVariantArray[x].Trim();
                                if (System.IO.Directory.Exists(strDestFolder) == false)
                                {
                                    //create fvs cutlist table
                                    //CA_TREE_CUTLIST.MDB
                                    strTable = "fvs_tree_IN_" + strVariantArray[x].Trim() + "_TREE_CUTLIST";
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection,frmMain.g_oTables.m_oFvs.CreateFVSOutProcessorInTableSQL(strTable));

                                    //create fvs file input  folder
                                    System.IO.Directory.CreateDirectory(strDestFolder);
                                    frmMain.g_sbpInfo.Text = "Version Update: Copying FVS Input And Output Files For Variant " + strVariantArray[x].Trim() + "...Stand By";
                                    //copy files
                                    utils.FileCopy(strSourceFolder, strDestFolder, false);
                                    //copy any fvs output files 
                                    strSourceFolder = ReferenceProjectDirectory + "\\fvs\\db\\out";
                                    z=0;
                                    for (y=0;y<=strRxArray.Length - 1;y++)
                                    {

                                        z=y+1;
                                        if (z.ToString().Trim().Length == 1)
                                            strRxPackage = "00" + z.ToString().Trim();
                                        else if (z.ToString().Trim().Length == 2)
                                            strRxPackage = "0" + z.ToString().Trim();
                                        else
                                            strRxPackage = z.ToString().Trim();

                                        if (strRxArray[y] != null && strRxArray[y].Length > 0)
                                        {

                                            //FVSOUT_CA_P000-050-051-000-050.MDB
                                            str = "FVSOUT_" + strVariantArray[x].Trim() + 
                                                   "_" + strRxArray[y].Trim() + ".mdb";
                                            str2 = "FVSOUT_" + strVariantArray[x].Trim() +
                                                   "_P" + strRxPackage + "-" + strNewRxArray[y].Trim() +
                                                   "-000-000-000.mdb";
                                            frmMain.g_sbpInfo.Text = "Version Update: Copy " + str + " to " + str2 + "...Stand By";
                                            if (System.IO.File.Exists(strSourceFolder + "\\" + str) == true &&
                                                System.IO.File.Exists(strDestFolder + "\\" + str2) == false)
                                            {
                                                System.IO.File.Copy(strSourceFolder + "\\" + str,
                                                                    strDestFolder + "\\" + str2,true);
                                            }

                                            frmMain.g_sbpInfo.Text = "Version Update: Insert " + oQueries.m_oFvs.m_strFvsTreeTable + " to " + strTable + " for variant " + strVariantArray[x].Trim() + "...Stand By";
                                            oAdo.m_strSQL= "INSERT INTO " + strTable + " " +
                                                           "SELECT * FROM " + oQueries.m_oFvs.m_strFvsTreeTable + " " + 
                                                           "WHERE fvs_variant='" + strVariantArray[x].Trim() + "' AND " + 
                                                                 "RX='" + strNewRxArray[y].Trim() + "'";

                                          
                                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                                            frmMain.g_sbpInfo.Text = "Version Update: Update RxPackage to " + strRxPackage + " AND RxCycle to 1 for variant " + strVariantArray[x].Trim() + "...Stand By";
                                            oAdo.m_strSQL = "UPDATE " + strTable + " " +
                                                            "SET rxpackage='" + strRxPackage + "'," +
                                                                "rxcycle='1' " +
                                                            "WHERE rx='" + strNewRxArray[y].Trim() + "'";

                                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                                        }
                                    }

                                    
                                }
                                

                            }
                        }
                        

                       

                    }
                    
                    
                    


                }

            }
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            if (strRxArray != null && strRxArray.Length > 0)
            {
                if (strRxArray[0].Trim().Length == 1)
                {
                    dao_data_access oDao = new dao_data_access();
                    //
                    //make sure treelist folder exist
                    //
                    strDestFolder = ReferenceProjectDirectory + "\\fvs\\data\\TreeLists";
                    if (System.IO.Directory.Exists(strDestFolder) == false)
                    {

                        //create tree cutlist folder
                        System.IO.Directory.CreateDirectory(strDestFolder);
                    }
                    //
                    //create links to variant cutlist table
                    //
                    for (x = 0; x <= strVariantArray.Length - 1; x++)
                    {
                        strTable = "fvs_tree_IN_" + strVariantArray[x].Trim() + "_TREE_CUTLIST";
                        //only process variant if the strTable exists
                        if (oDao.TableExists(oQueries.m_strTempDbFile,strTable))
                        {

                            if (strVariantArray[x] != null &&
                                strVariantArray[x].Trim().Length > 0)
                            {
                                //CA_TREE_CUTLIST.MDB
                                strTable = "fvs_tree_IN_" + strVariantArray[x].Trim() + "_TREE_CUTLIST";
                                if (System.IO.File.Exists(strDestFolder + "\\" + strVariantArray[x].Trim() + "_TREE_CUTLIST.MDB"))
                                {
                                    System.IO.File.Delete(strDestFolder + "\\" + strVariantArray[x].Trim() + "_TREE_CUTLIST.MDB");
                                    System.Threading.Thread.Sleep(2000);

                                }
                                oDao.CreateMDB(strDestFolder + "\\" + strVariantArray[x].Trim() + "_TREE_CUTLIST.MDB");
                                oDao.CreateTableLink(strDestFolder + "\\" + strVariantArray[x].Trim() + "_TREE_CUTLIST.MDB", strTable, oQueries.m_strTempDbFile, strTable);
                             }
                        }
                    }
                    oDao.m_DaoWorkspace.Close();
                    oDao = null;
                }
            }
            
            if (strRxArray != null && strRxArray.Length > 0)
            {
                if (strRxArray[0].Trim().Length == 1)
                {
                    
                    //
                    //make sure treelist folder exist
                    //
                    strDestFolder = ReferenceProjectDirectory + "\\fvs\\data\\TreeLists";
                   
                    //
                    //create cutlist mdb and copy table to it
                    //
                    for (x = 0; x <= strVariantArray.Length - 1; x++)
                    {
                        if (System.IO.File.Exists(strDestFolder + "\\" + strVariantArray[x].Trim() + "_TREE_CUTLIST.MDB"))
                        {
                            oAdo.OpenConnection(oAdo.getMDBConnString(strDestFolder + "\\" + strVariantArray[x].Trim() + "_TREE_CUTLIST.MDB", "", ""));
                            strTable = "fvs_tree_IN_" + strVariantArray[x].Trim() + "_TREE_CUTLIST";
                            //only process variant if the strTable exists
                            if (oAdo.TableExist(oAdo.m_OleDbConnection, strTable))
                            {

                                if (strVariantArray[x] != null &&
                                    strVariantArray[x].Trim().Length > 0)
                                {
                                    //CA_TREE_CUTLIST.MDB
                                    if ((int)oAdo.getRecordCount(oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM " + strTable, "TEMP") > 0)
                                    {
                                        frmMain.g_sbpInfo.Text = "Version Update: Copy table " + strTable + " into FVS_Tree table for variant " + strVariantArray[x].Trim() + "...Stand By";
                                        oAdo.m_strSQL = "SELECT * INTO fvs_tree FROM " + strTable;
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + strTable);
                                    }
                                    else
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + strTable);
                                    }





                                }
                            }
                            oAdo.CloseConnection(oAdo.m_OleDbConnection);
                        }
                    }
                    


                }
            }
            









        }
        private void CheckFVSOutPrePostValues()
        {
            int x,y;
            
            string strAlpha="A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z";
            string[] strAlphaArray = frmMain.g_oUtils.ConvertListToArray(strAlpha, ",");
            string strSourceFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\FVS\\DB\\biosum_fvsout_prepost_rx.mdb";
            if (System.IO.File.Exists(strSourceFile) == false) return;

            Queries oQueries = new Queries();
            oQueries.m_oFvs.LoadDatasource = true;
            oQueries.LoadDatasources(true);
            
            ado_data_access oAdo = new ado_data_access();

            //
            //get the number of treatments
            //
            int intRxCount = (int)oAdo.getRecordCount(oAdo.getMDBConnString(oQueries.m_strTempDbFile, "", ""), "SELECT COUNT(*) FROM " + oQueries.m_oFvs.m_strRxTable, "temp");
            
            oAdo.OpenConnection(oAdo.getMDBConnString(strSourceFile, "", ""));



            string[] strTableArray = oAdo.getTableNames(oAdo.m_OleDbConnection);

            if (strTableArray != null)
            {
                
                for (x = 0; x <= strTableArray.Length - 1; x++)
                {
                    if (strTableArray[x] != null &&
                        strTableArray[x].Trim().Length >= 7)
                    {
                        if (strTableArray[x].Substring(0, 7).ToUpper() == "PRE_FVS")
                        {
                            if (oAdo.ColumnExist(oAdo.m_OleDbConnection, strTableArray[x], "RX"))
                            {
                            }
                            else
                            {
                                oAdo.m_strSQL = "ALTER TABLE " + strTableArray[x] + " ADD COLUMN RXPACKAGE TEXT (3)";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                oAdo.m_strSQL = "ALTER TABLE " + strTableArray[x] + " ADD COLUMN RX TEXT (3)";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                oAdo.m_strSQL = "ALTER TABLE " + strTableArray[x] + " ADD COLUMN RXCYCLE TEXT (1)";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                oAdo.m_strSQL = "UPDATE " + strTableArray[x] + " SET RX='A'";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                for (y = 2; y <= intRxCount; y++)
                                {
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, "PRETEMP"))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE PRETEMP");
                                    }
                                    oAdo.m_strSQL = "SELECT * INTO PRETEMP FROM " + strTableArray[x] + " WHERE TRIM(RX)='A'";
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                    oAdo.m_strSQL = "UPDATE PRETEMP SET RX='" + strAlphaArray[y - 1] + "'";
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                    oAdo.m_strSQL = "INSERT INTO " + strTableArray[x] + " SELECT * FROM PRETEMP";
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                                }
                                this.ConvertRxAndRxPackageAndRxCycle(oAdo, oAdo.m_OleDbConnection, strTableArray[x]);

                            }

                        }
                        else
                        {
                            if (strTableArray[x] != null &&
                                strTableArray[x].Trim().Length > 7)
                            {
                                if (strTableArray[x].Substring(0, 8).ToUpper() == "POST_FVS")
                                {
                                    if (oAdo.ColumnExist(oAdo.m_OleDbConnection, strTableArray[x], "RXPACKAGE"))
                                    {
                                    }
                                    else
                                    {
                                        oAdo.m_strSQL = "ALTER TABLE " + strTableArray[x] + " ALTER COLUMN RX TEXT (3)";
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                        oAdo.m_strSQL = "ALTER TABLE " + strTableArray[x] + " ADD COLUMN RXPACKAGE TEXT (3)";
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                        oAdo.m_strSQL = "ALTER TABLE " + strTableArray[x] + " ADD COLUMN RXCYCLE TEXT (1)";
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                        this.ConvertRxAndRxPackageAndRxCycle(oAdo, oAdo.m_OleDbConnection, strTableArray[x]);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            oAdo.CloseConnection(oAdo.m_OleDbConnection);




        }
        private void ConvertRx(ado_data_access p_oAdo, 
                               System.Data.OleDb.OleDbConnection p_oConn,
                               string p_strTableName)
        {
            p_oAdo.m_strSQL = "UPDATE " + p_strTableName + " " +
                                    "SET RX = IIF(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)) > 0," +
                                       "'7' + IIF(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)) <10," +
                                       "'0' + CSTR(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)))," +
                                       "CSTR(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)))),RX)";

            p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);



        }

        private void ConvertRxAndRxPackageAndRxCycle(ado_data_access p_oAdo,
                               System.Data.OleDb.OleDbConnection p_oConn,
                               string p_strTableName)
        {
            p_oAdo.m_strSQL = "UPDATE " + p_strTableName + " " +
                              "SET RXPACKAGE = " +
                                      "IIF(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)) > 0," +
                                         "'0' + IIF(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)) <10," +
                                         "'0' + CSTR(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)))," +
                                         "CSTR(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)))),RXPACKAGE)," +
                                  "RXCYCLE=IIF(LEN(TRIM(RXCYCLE)) > 0,RXCYCLE,'1')," +
                                  "RX=IIF(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)) > 0," +
                                         "'7' + IIF(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)) <10," +
                                         "'0' + CSTR(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)))," +
                                         "CSTR(INSTR(1,'ABCDEFGHIJKLMNOPQRSTUVWXYZ',TRIM(rx)))),RX)";
            p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);

        }
        private void CleanUp()
        {
            ado_data_access oAdo = new ado_data_access();

            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));

            //remove obsolete table types
            //TREE VOL VAL AND HARVEST COSTS ARE NOW SCENARIO BASED IN THE PROCESSOR
            oAdo.m_strSQL = "DELETE FROM datasource WHERE " +
                "TRIM(UCASE(table_type)) IN (" +
                "'FIRE AND FUEL EFFECTS'," +
                "'TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS'," +
                "'HARVEST COSTS'," +
                "'FVS TREE LIST FOR PROCESSOR')";

            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
        }
        private void UpgradeFVSOutTreeListFiles()
        {
            int x,y;
            string strFile="";
            List<string> strVariant = new List<string>(); //hold the variants that have an existing variant_tree_cutlist.mdb
            string[] strPackageArray = null;
            
            string[] strFiles = System.IO.Directory.GetFiles(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\FVS\\Data\\TreeLists","*.mdb");
            if (strFiles != null && strFiles.Length > 0)
            {
                frmMain.g_sbpInfo.Text = "Version Update: Update FVS_CUTLIST...Stand By";
                ado_data_access oAdo = new ado_data_access();
                RxTools oRxTools = new RxTools();
                Queries oQueries = new Queries();
                RxPackageItem_Collection oRxPackageItemCollection = new RxPackageItem_Collection();
			    oQueries.m_oFvs.LoadDatasource=true;
			    oQueries.m_oFIAPlot.LoadDatasource=true;
			    oQueries.LoadDatasources(true);
                //
                //open the work file containing the links
                //
                oAdo.OpenConnection(oAdo.getMDBConnString(oQueries.m_strTempDbFile,"",""));
                //
                //get all the variants
                //
			    string strVariantsList = oRxTools.GetListOfFVSVariantsInPlotTable(oAdo,oAdo.m_OleDbConnection,oQueries.m_oFIAPlot.m_strPlotTable);
                string[] strVariantsArray=frmMain.g_oUtils.ConvertListToArray(strVariantsList,",");
                //
                //get all the packages
                //
                oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(oAdo, oAdo.m_OleDbConnection, oQueries, oRxPackageItemCollection);
                //
                //close the work file containing the links
                //
			    oAdo.CloseConnection(oAdo.m_OleDbConnection);
                //
                //create DAO instance
                //
			    dao_data_access oDao = new dao_data_access();
                //
                //create temp fvscutlist file and table
                //
                string strTempFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "ACCDB");
                oDao.CreateMDB(strTempFile);
                oAdo.OpenConnection(oAdo.getMDBConnString(strTempFile, "", ""));
                frmMain.g_oTables.m_oFvs.CreateFVSOutProcessorIn(oAdo, oAdo.m_OleDbConnection, Tables.FVS.DefaultFVSTreeTableName);
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                //
                //create links to the source (old) fvs tree cutlist tables
                //
			    string strSourceTreeListDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()  + "\\fvs\\data\\TreeLists";
			    string strDestinationTreeListLinkTableName="";
                string strSourceTreeListLinkTableName="";
			    for (x=0;x<=strVariantsArray.Length-1;x++)
			    {
                    //make sure BiosumCalc directory exists
                    if (System.IO.Directory.Exists(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\data\\" + strVariantsArray[x].Trim() + "\\BiosumCalc") == false)
                    {
                        System.IO.Directory.CreateDirectory(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\data\\" + strVariantsArray[x].Trim() + "\\BiosumCalc");
                    }
                    //source table link
                    strSourceTreeListLinkTableName = strVariantsArray[x].Trim() + "_TREE_CUTLIST";
                    if (System.IO.File.Exists(strSourceTreeListDir + "\\" + strSourceTreeListLinkTableName.Trim() + ".MDB"))
                    {
                        oDao.CreateTableLink(oQueries.m_strTempDbFile, "fvs_tree_IN_" + strSourceTreeListLinkTableName.Trim(), strSourceTreeListDir + "\\" + strSourceTreeListLinkTableName.Trim() + ".MDB", "fvs_tree", true);
                        strVariant.Add(strVariantsArray[x].Trim());
                        
                    }
                    
			
			    }
                //
                //remove DAO instance
                //
                oDao.m_DaoWorkspace.Close();
                oDao=null;
                //
                //reopen the work file containing the links
                //
                oAdo.OpenConnection(oAdo.getMDBConnString(oQueries.m_strTempDbFile,"",""));
                //
                //create the variant + package cut list mdb file (new)
                //
                for (x = 0; x <= strVariant.Count - 1; x++)
                {
                    //get a list of packages in the old cutlist mdb
                    strSourceTreeListLinkTableName = "fvs_tree_IN_" + strVariant[x].Trim() + "_TREE_CUTLIST";
                    string strPackageList = oAdo.CreateCommaDelimitedList(oAdo.m_OleDbConnection, "SELECT DISTINCT rxpackage FROM " + strSourceTreeListLinkTableName, ",");
                    if (strPackageList.Trim().Length > 0)
                    {
                        strPackageArray = frmMain.g_oUtils.ConvertListToArray(strPackageList, ",");
                        if (strPackageArray != null && strPackageArray.Length > 0)
                        {
                            for (y = 0; y <= strPackageArray.Length - 1; y++)
                            {
                                if (strPackageArray[y] != null && strPackageArray[y].Trim().Length > 0)
                                {
                                    //destination table creation
                                    strFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\data\\" + strVariant[x].Trim() + "\\BiosumCalc\\" + strVariant[x].Trim() + "_P" + strPackageArray[y].Trim() + "_TREE_CUTLIST.MDB";
                                    if (System.IO.File.Exists(strFile) == false)
                                    {
                                        System.IO.File.Copy(strTempFile, strFile, false);
                                    }
                                }

                            }
                        }
                    }
                }
                //
                //close the work file containing the links
                //
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                //
                //create links to variant\BiosumCalc\variant_package_tree_cutlist.mdb files
                //
                // 29-JUL-2022: No longer valid method
                //oRxTools.CreateTableLinksToFVSOutTreeListTables(oQueries, oQueries.m_strTempDbFile);
                //
                //reopen the work file containing the links
                //
                oAdo.OpenConnection(oAdo.getMDBConnString(oQueries.m_strTempDbFile, "", ""));
                //
                //get processing date and time
                //
                System.DateTime oDate = System.DateTime.Now;
                string strDateFormat = "yyyy-MM-dd_HH-mm-ss";
                string strFileDate = oDate.ToString(strDateFormat);
                strFileDate = strFileDate.Replace("/", "_"); strFileDate = strFileDate.Replace(":", "_");
                //
                //append for source trees to the destination trees
                //
                if (strVariant != null)
                {
                    for (x = 0; x <= strVariant.Count - 1; x++)
                    {
                        for (y = 0; y <= oRxPackageItemCollection.Count - 1; y++)
                        {

                            strDestinationTreeListLinkTableName = "fvs_tree_IN_" + strVariant[x].Trim() + "_P" + oRxPackageItemCollection.Item(y).RxPackageId + "_TREE_CUTLIST";
                            strSourceTreeListLinkTableName = "fvs_tree_IN_" + strVariant[x].Trim() + "_TREE_CUTLIST";
                            if (oAdo.TableExist(oAdo.m_OleDbConnection, strSourceTreeListLinkTableName) &&
                                oAdo.TableExist(oAdo.m_OleDbConnection, strDestinationTreeListLinkTableName))
                            {
                                //make a backup of the destination file
                                frmMain.g_sbpInfo.Text = "Version Update: FVS_CUTLIST - Make backup of " + strVariant[x].Trim() + "_P" + oRxPackageItemCollection.Item(y).RxPackageId + "_TREE_CUTLIST.MDB...Stand By";
                                strFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\data\\" + strVariant[x].Trim() + "\\BiosumCalc\\" + strVariant[x].Trim() + "_P" + oRxPackageItemCollection.Item(y).RxPackageId + "_TREE_CUTLIST.MDB";
                                System.IO.File.Copy(strFile, strFile + "_" + strFileDate);
                                //delete any current records from the destination table
                                oAdo.m_strSQL = "DELETE FROM " + strDestinationTreeListLinkTableName;
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                //insert the package
                                frmMain.g_sbpInfo.Text = "Version Update: FVS_CUTLIST - Insert Records into " + strVariant[x].Trim() + "_P" + oRxPackageItemCollection.Item(y).RxPackageId + "_TREE_CUTLIST.FVS_TREE table...Stand By";
                                oAdo.m_strSQL = "INSERT INTO " + strDestinationTreeListLinkTableName + " " +
                                                "SELECT * FROM " + strSourceTreeListLinkTableName + " " +
                                                "WHERE rxpackage = '" + oRxPackageItemCollection.Item(y).RxPackageId + "'";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);



                            }
                        }

                    }
                }
                //
                //close connection
                //
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                //
                //rename the old fvs cutlist files
                //
                for (x = 0; x <= strFiles.Length - 1; x++)
                {
                    if (strFiles[x].Substring(strFiles[x].Length - 4, 4) == ".MDB")
                    {
                        System.IO.File.Copy(strFiles[x].Trim(), strFiles[x].Trim() + "_VersionControlUpgrade_" + strFileDate);
                        System.IO.File.Delete(strFiles[x].Trim());
                    }
                }
                
			

            }
        }
        private void UpgradeToPrePostSeqNumMatrix()
        {
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();

            

           

            ado_data_access oAdo = new ado_data_access();
           
            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));
            uc_fvs_output_prepost_seqnum.InitializePrePostSeqNumTables(oAdo, ReferenceProjectDirectory.Trim() + "\\" + Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile);

           

            oAdo.CloseConnection(oAdo.m_OleDbConnection);

           

            oAdo = null;
        }
        
        private void UpgradeFVSOutPOTFireFiles()
        {
            int y;
           
            string strFVSOutDbFile="";
            string strVariant="";
            string strPackage="";
            
            System.Data.OleDb.OleDbConnection oFVSOutConn;
            string strConn="";
            

           
                frmMain.g_sbpInfo.Text = "Version Update: Update FVS_POTFIRE...Stand By";
                ado_data_access oAdo = new ado_data_access();
                
                
                RxTools oRxTools = new RxTools();
                Queries oQueries = new Queries();
                RxPackageItem_Collection oRxPackageItemCollection = new RxPackageItem_Collection();
                oQueries.m_oFvs.LoadDatasource = true;
                oQueries.m_oFIAPlot.LoadDatasource = true;
                oQueries.LoadDatasources(true);
                //
                //get processing date and time
                //
                System.DateTime oDate = System.DateTime.Now;
                string strDateFormat = "yyyy-MM-dd_HH-mm-ss";
                string strFileDate = oDate.ToString(strDateFormat);
                strFileDate = strFileDate.Replace("/", "_"); strFileDate = strFileDate.Replace(":", "_");
                //
                //open the work file containing the links
                //
                oAdo.OpenConnection(oAdo.getMDBConnString(oQueries.m_strTempDbFile, "", ""));
                //
                //get all the packages
                //
                //oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(oAdo, oAdo.m_OleDbConnection, oQueries, oRxPackageItemCollection);
                //
                //close the work file containing the links
                //
                //oAdo.CloseConnection(oAdo.m_OleDbConnection);
  
                string strSourceDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\data";



                oAdo.m_strSQL = Queries.FVS.GetFVSVariantRxPackageSQL(oQueries.m_oFIAPlot.m_strPlotTable,oQueries.m_oFvs.m_strRxPackageTable);

                System.Data.OleDb.OleDbDataReader oDataReader;

                System.Data.OleDb.OleDbCommand oDbCommand = new System.Data.OleDb.OleDbCommand();

                oDbCommand = oAdo.m_OleDbConnection.CreateCommand();

                oDbCommand.CommandText = oAdo.m_strSQL;

                oDataReader = oDbCommand.ExecuteReader();

                while (oDataReader.Read())
                {


                    strVariant = oDataReader["fvs_variant"].ToString().Trim();
                    strPackage = oDataReader["RxPackage"].ToString().Trim();



                    strFVSOutDbFile = oRxTools.GetRxPackageFvsOutDbFileName(oDataReader);
                    if (System.IO.File.Exists(strSourceDir + "\\" + strVariant + "\\" + strFVSOutDbFile))
                    {
                        
                        strConn = oAdo.getMDBConnString(strSourceDir + "\\" + strVariant + "\\" + strFVSOutDbFile, "", "");
                        oFVSOutConn = new System.Data.OleDb.OleDbConnection();
                        oAdo.OpenConnection(strConn, ref oFVSOutConn);
                        if (oAdo.m_intError == 0)
                        {
                            frmMain.g_sbpInfo.Text = "Version Update: Update POTFIRE " + strFVSOutDbFile + "...Stand By";

                            //check if tables fvs_cases,fvs_potfire exist and the column BIOSUM_PotFireOneYearAdded_YN
                            if (oAdo.TableExist(oFVSOutConn, "FVS_CASES") &&
                                oAdo.ColumnExist(oFVSOutConn, "FVS_CASES", "BIOSUM_PotFireOneYearAdded_YN") &&
                                oAdo.TableExist(oFVSOutConn, "FVS_POTFIRE"))
                            {
                                System.IO.File.Copy(strSourceDir + "\\" + strVariant + "\\" + strFVSOutDbFile,
                                           strSourceDir + "\\" + strVariant + "\\" + strFVSOutDbFile + "_VersionControlUpgrade_" + strFileDate);

                                if (oAdo.TableExist(oFVSOutConn, "min_potfireyear"))
                                    oAdo.SqlNonQuery(oFVSOutConn, "DROP TABLE min_potfireyear");
                                //reset the FVS_POTFIRE to its original non-baseyear state
                                oAdo.m_strSQL = "UPDATE FVS_POTFIRE p " +
                                                "INNER JOIN FVS_CASES c " +
                                                "ON p.caseID = c.caseID AND " +
                                                   "p.standID = c.standID " +
                                                "SET p.year = IIF(c.BIOSUM_PotFireOneYearAdded_YN='Y',p.year - 1,p.year)," +
                                                    "c.BIOSUM_PotFireOneYearAdded_YN='N'," +
                                                    "C.BIOSUM_Append_YN='N'";
                                oAdo.SqlNonQuery(oFVSOutConn, oAdo.m_strSQL);

                                oAdo.m_strSQL = "SELECT b.min_year, a.standid " +
                                                "INTO min_potfireyear FROM FVS_POTFIRE a," +
                                                   "(SELECT MIN(year) as min_year,standid FROM FVS_POTFIRE GROUP BY standid) b " +
                                                "WHERE a.standid=b.standid AND a.year=b.min_year;";

                                oAdo.SqlNonQuery(oFVSOutConn, oAdo.m_strSQL);

                                oAdo.m_strSQL = "DELETE DISTINCTROW a.* " +
                                                "FROM FVS_POTFIRE a " +
                                                "INNER JOIN min_potfireyear b  ON a.standid = b.standid AND a.year = b.min_year";

                                oAdo.SqlNonQuery(oFVSOutConn, oAdo.m_strSQL);

                                oAdo.m_strSQL = "ALTER TABLE FVS_CASES DROP COLUMN BIOSUM_PotFireOneYearAdded_YN";
                                oAdo.SqlNonQuery(oFVSOutConn, oAdo.m_strSQL);

                                if (oAdo.TableExist(oFVSOutConn, "min_potfireyear"))
                                    oAdo.SqlNonQuery(oFVSOutConn, "DROP TABLE min_potfireyear");

                            }

                        }
                        oAdo.CloseConnection(oFVSOutConn);

                    }


                }
                oDataReader.Close();
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
           
            
               


            
        }
        private void UpgradeFVSOutPREPOSTTables()
        {
            int x;
            string strPreTable = "";
            string strPostTable = "";
            string strPREPOSTDbFile = "";
            string strSQL = "";
            frmMain.g_sbpInfo.Text = "Version Update: Update FVS PREPOST tables...Stand By";
            dao_data_access oDao = new dao_data_access();
            string strFVSOutPrePostPathAndDbFile = 
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\db\\biosum_fvsout_prepost_rx.mdb";
            if (System.IO.File.Exists(strFVSOutPrePostPathAndDbFile))
            {

                for (x = 0; x <= Tables.FVS.g_strFVSOutTablesArray.Length - 1; x++)
                {
                    strPreTable = "PRE_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim();
                    strPostTable = "POST_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim();
                    if (oDao.TableExists(strFVSOutPrePostPathAndDbFile, strPreTable) &&
                        oDao.TableExists(strFVSOutPrePostPathAndDbFile, strPostTable))
                    {
                        frmMain.g_sbpInfo.Text = "Version Update: Update PREPOST " + Tables.FVS.g_strFVSOutTablesArray[x].Trim() + "...Stand By";
                        strPREPOSTDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\db\\PREPOST_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim() + ".ACCDB";
                        if (!System.IO.File.Exists(strPREPOSTDbFile))
                        {
                            oDao.CreateMDB(strPREPOSTDbFile);
                        }
                        oDao.CreateTableLink(strPREPOSTDbFile, strPreTable + "_temp", strFVSOutPrePostPathAndDbFile, strPreTable, true);
                        oDao.CreateTableLink(strPREPOSTDbFile, strPostTable + "_temp", strFVSOutPrePostPathAndDbFile, strPostTable, true);
                        oDao.OpenDb(strPREPOSTDbFile);
                        strSQL = "SELECT * INTO " + strPreTable + " FROM " + strPreTable + "_temp";
                        oDao.m_DaoDatabase.Execute(strSQL, null);
                        strSQL = "SELECT * INTO " + strPostTable + " FROM " + strPostTable + "_temp";
                        oDao.m_DaoDatabase.Execute(strSQL, null);
                        oDao.DeleteTableFromMDB(oDao.m_DaoDatabase, strPreTable + "_temp");
                        oDao.DeleteTableFromMDB(oDao.m_DaoDatabase, strPostTable + "_temp");
                        oDao.m_DaoDatabase.Close();


                    }

                }
            }

        }
        private void UpdateAuditDbFile_5_3_0()
        {
            
		    
			int x;
			//
            //MAIN PROJECT DATASOURCE
            //
			FIA_Biosum_Manager.Datasource oDs = new Datasource();
			oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
			oDs.m_strDataSourceTableName = "datasource";
			oDs.m_strScenarioId="";
			oDs.LoadTableColumnNamesAndDataTypes=false;
			oDs.LoadTableRecordCount=false;
			oDs.populate_datasource_array();

            

			ado_data_access oAdo=new ado_data_access();
			
            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));

            oAdo.m_strSQL = "DELETE FROM datasource WHERE TRIM(UCASE(table_type))='PLOT AND CONDITION RECORD AUDIT'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);


			oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
				"Plot And Condition Record Audit",
				ReferenceProjectDirectory.Trim() + "\\db",
                frmMain.g_oUtils.getFileName(Tables.Audit.DefaultCondAuditTableDbFile),
                Tables.Audit.DefaultCondAuditTableName);


            oAdo.m_strSQL = "DELETE FROM datasource WHERE TRIM(UCASE(table_type))='PLOT, CONDITION AND TREATMENT RECORD AUDIT'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

			oDs.InsertDatasourceRecord(oAdo,oAdo.m_OleDbConnection,
				"Plot, Condition And Treatment Record Audit",
				ReferenceProjectDirectory.Trim() + "\\db",
                frmMain.g_oUtils.getFileName(Tables.Audit.DefaultCondRxAuditTableDbFile),
                Tables.Audit.DefaultCondRxAuditTableName);


			oAdo.CloseConnection(oAdo.m_OleDbConnection);

            dao_data_access oDao = new dao_data_access();
            if (!System.IO.File.Exists(ReferenceProjectDirectory.Trim() + "\\db\\Audit.mdb"))
            {
                oDao.CreateMDB(ReferenceProjectDirectory.Trim() + "\\db\\Audit.mdb");
            }
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\db\\Audit.mdb", "", ""));
            if (oAdo.TableExist(oAdo.m_OleDbConnection, "plot_cond_audit"))
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE plot_cond_audit");
            if (oAdo.TableExist(oAdo.m_OleDbConnection, "plot_audit"))
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE  ");
            if (oAdo.TableExist(oAdo.m_OleDbConnection, "plot_cond_rx_audit"))
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE plot_cond_rx_audit");


            frmMain.g_oTables.m_oAudit.CreateCondAuditTable(oAdo, oAdo.m_OleDbConnection, Tables.Audit.DefaultCondAuditTableName);
            frmMain.g_oTables.m_oAudit.CreatePlotCondRxAuditTable(oAdo, oAdo.m_OleDbConnection, Tables.Audit.DefaultCondRxAuditTableName);
            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            //
            //CORE SCENARIO
            //
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\core\\db\\scenario_core_rule_definitions.mdb","",""));
            oAdo.m_strSQL = "UPDATE scenario_datasource " +
                            "SET table_name='" + Tables.Audit.DefaultCondAuditTableName + "'," +
                                "path='" + ReferenceProjectDirectory.Trim() + "\\db' " +
                            "WHERE TRIM(UCASE(table_type))='PLOT AND CONDITION RECORD AUDIT'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            oAdo.m_strSQL = "UPDATE scenario_datasource " +
                            "SET table_name='" + Tables.Audit.DefaultCondRxAuditTableName + "'," +
                                "path='" + ReferenceProjectDirectory.Trim() + "\\db' " +
                            "WHERE TRIM(UCASE(table_type))='PLOT, CONDITION AND TREATMENT RECORD AUDIT'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            //
            //PROCESSOR SCENARIO
            //
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\processor\\db\\scenario_processor_rule_definitions.mdb", "", ""));
            oAdo.m_strSQL = "UPDATE scenario_datasource " +
                            "SET table_name='" + Tables.Audit.DefaultCondAuditTableName + "'," +
                                "path='" + ReferenceProjectDirectory.Trim() + "\\db' " +
                            "WHERE TRIM(UCASE(table_type))='PLOT AND CONDITION RECORD AUDIT'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            oAdo.m_strSQL = "UPDATE scenario_datasource " +
                            "SET table_name='" + Tables.Audit.DefaultCondRxAuditTableName + "'," +
                                "path='" + ReferenceProjectDirectory.Trim() + "\\db' " +
                            "WHERE TRIM(UCASE(table_type))='PLOT, CONDITION AND TREATMENT RECORD AUDIT'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oDao = null;
            oAdo = null;

            //
            //DELETE ANY AUDIT_VARIANT.MDB FILES
            //
            string[] strFiles = System.IO.Directory.GetFiles(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\DB", "AUDIT_*.mdb");
            for (x = 0; x <= strFiles.Length - 1; x++)
            {
                System.IO.File.Delete(strFiles[x].Trim());
            }
            strFiles.Initialize();

		}
        private void UpdateDatasources_5_3_0()
        {

            string strSQL;

            //
            //MAIN PROJECT DATASOURCE
            //
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();



            ado_data_access oAdo = new ado_data_access();

            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));

            oAdo.m_strSQL = "DELETE FROM datasource WHERE TRIM(UCASE(table_type))='FVS PRE-POST SEQNUM DEFINITIONS'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                        "('FVS PRE-POST SeqNum Definitions'," +
                        "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                        "'fvsmaster.mdb'," +
                        "'" + Tables.FVS.DefaultFVSPrePostSeqNumTable + "');";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);

            oAdo.m_strSQL = "DELETE FROM datasource WHERE TRIM(UCASE(table_type))='FVS PRE-POST SEQNUM TREATMENT PACKAGE ASSIGN'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                "('FVS PRE-POST SeqNum Treatment Package Assign'," +
                "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                "'fvsmaster.mdb'," +
                "'" + Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable + "');";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);

            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            oAdo = null;

           
        }
        private void UpdateDatasources_5_6_0()
        {
            dao_data_access oDao = null;
            string strSQL;
            string strPath = "";
            string strFile = "";
            string strTable="";
            System.Data.OleDb.OleDbConnection oMasterConn=null;
            
                

            //
            //MAIN PROJECT DATASOURCE
            //
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();
            



            ado_data_access oAdo = new ado_data_access();

            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));
            //
            //BREAKPOINT DIAMETER DATASOURCE
            //
            if ((int)oAdo.getRecordCount(oAdo.m_OleDbConnection, "SELECT table_type FROM datasource WHERE TRIM(UCASE(table_type))='FIA TREE MACRO PLOT BREAKPOINT DIAMETER'", "DATASOURCE") == 0)
            {
                //append the new datasource
                strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                           "('FIA Tree Macro Plot Breakpoint Diameter'," +
                            "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                           "'ref_master.mdb'," +
                           "'TreeMacroPlotBreakPointDia');";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                //add the new reference table
                oDao = new dao_data_access();
                oDao.CreateTableLink(ReferenceProjectDirectory.Trim() + "\\db\\ref_master.mdb", "breakpointupdate_v560", frmMain.g_oEnv.strAppDir + "\\db\\ref_master.mdb", "TreeMacroPlotBreakPointDia");
                strSQL = "SELECT * INTO TreeMacroPlotBreakPointDia FROM breakpointupdate_v560";
                oDao.OpenDb(ReferenceProjectDirectory.Trim() + "\\db\\ref_master.mdb");
                oDao.m_DaoDatabase.Execute(strSQL,null);
                oDao.m_DaoDatabase.Execute("DROP TABLE breakpointupdate_v560",null);
                oDao.m_DaoDatabase.Close();
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }

            //
            //BIOSUM CALCULATED ADJUSTMENT FACTORS DATASOURCE
            //
            if ((int)oAdo.getRecordCount(oAdo.m_OleDbConnection, "SELECT table_type FROM datasource WHERE TRIM(UCASE(table_type))='BIOSUM POP STRATUM ADJUSTMENT FACTORS'", "DATASOURCE") == 0)
            {
                //append the new datasource
                strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                           "('BIOSUM Pop Stratum Adjustment Factors'," +
                            "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                           "'master.mdb'," +
                           "'" + frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName + "');";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                //create the new table
                strPath = ReferenceProjectDirectory + "\\db\\master.mdb";
                oMasterConn = new System.Data.OleDb.OleDbConnection();
                oAdo.OpenConnection(oAdo.getMDBConnString(strPath, "", ""), ref oMasterConn);
                frmMain.g_oTables.m_oFIAPlot.CreateBiosumPopStratumAdjustmentFactorsTable(oAdo, oMasterConn, frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName);
            }
            strPath = "";
            //
            //ADD PLOT TABLE COLUMN MACRO_BREAKPOINT_DIA
            //
            strSQL = "SELECT * FROM datasource WHERE TRIM(UCASE(table_type)) = 'PLOT'";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                oAdo.m_OleDbDataReader.Read();
                strPath = oAdo.m_OleDbDataReader["path"].ToString().Trim();
                strFile = oAdo.m_OleDbDataReader["file"].ToString().Trim();
                strTable = oAdo.m_OleDbDataReader["table_name"].ToString().Trim();
            }
            oAdo.m_OleDbDataReader.Close();
            if (strPath.Trim().Length > 0)
            {
                strPath = strPath + "\\" + strFile;
                if (System.IO.File.Exists(strPath))
                {
                    if (oMasterConn == null) oMasterConn = new System.Data.OleDb.OleDbConnection();
                    else
                    {
                        if (strPath.Trim().ToUpper() != oMasterConn.DataSource.Trim().ToUpper())
                        {
                            oAdo.CloseConnection(oMasterConn);
                            oAdo.OpenConnection(oAdo.getMDBConnString(strPath, "", ""), ref oMasterConn);

                        }
                    }
                   
                    if (oAdo.TableExist(oMasterConn,strTable) && !oAdo.ColumnExist(oMasterConn,strTable,"MACRO_BREAKPOINT_DIA"))
                    {
                        oAdo.AddColumn(oMasterConn, strTable, "macro_breakpoint_dia", "INTEGER","");
                    }
                    
                }
            }
            strPath = ""; strFile = ""; strTable = "";
            //
            //ADD TREE TABLE COLUMN CONDPROP_SPECIFIC
            //
            strSQL = "SELECT * FROM datasource WHERE TRIM(UCASE(table_type)) = 'TREE'";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                oAdo.m_OleDbDataReader.Read();
                strPath = oAdo.m_OleDbDataReader["path"].ToString().Trim();
                strFile = oAdo.m_OleDbDataReader["file"].ToString().Trim();
                strTable = oAdo.m_OleDbDataReader["table_name"].ToString().Trim();
            }
            oAdo.m_OleDbDataReader.Close();
            if (strPath.Trim().Length > 0)
            {
                strPath = strPath + "\\" + strFile;
                if (System.IO.File.Exists(strPath))
                {
                    if (oMasterConn == null) oMasterConn = new System.Data.OleDb.OleDbConnection();
                    else
                    {
                        if (strPath.Trim().ToUpper() != oMasterConn.DataSource.Trim().ToUpper())
                        {
                            oAdo.CloseConnection(oMasterConn);
                            oAdo.OpenConnection(oAdo.getMDBConnString(strPath, "", ""), ref oMasterConn);

                        }
                    }


                    if (oAdo.TableExist(oMasterConn, strTable) && !oAdo.ColumnExist(oMasterConn, strTable, "CONDPROP_SPECIFIC"))
                    {
                        oAdo.AddColumn(oMasterConn, strTable, "condprop_specific", "DOUBLE", "");
                    }

                }
            }
            strPath = ""; strFile = ""; strTable = "";
            //
            //ADD COND TABLE COLUMNS MICRPROP_UNADJ, SUBPPROP_UNADJ, and MACRPROP_UNADJ
            //
            strSQL = "SELECT * FROM datasource WHERE TRIM(UCASE(table_type)) = 'CONDITION'";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                oAdo.m_OleDbDataReader.Read();
                strPath = oAdo.m_OleDbDataReader["path"].ToString().Trim();
                strFile = oAdo.m_OleDbDataReader["file"].ToString().Trim();
                strTable = oAdo.m_OleDbDataReader["table_name"].ToString().Trim();
            }
            oAdo.m_OleDbDataReader.Close();
            if (strPath.Trim().Length > 0)
            {
                strPath = strPath + "\\" + strFile;
                if (System.IO.File.Exists(strPath))
                {
                    if (oMasterConn == null) oMasterConn = new System.Data.OleDb.OleDbConnection();
                    else
                    {
                        if (strPath.Trim().ToUpper() != oMasterConn.DataSource.Trim().ToUpper())
                        {
                            oAdo.CloseConnection(oMasterConn);
                            oAdo.OpenConnection(oAdo.getMDBConnString(strPath, "", ""), ref oMasterConn);

                        }
                    }


                    if (oAdo.TableExist(oMasterConn, strTable))
                    {
                        if (!oAdo.ColumnExist(oMasterConn, strTable, "MICRPROP_UNADJ"))
                            oAdo.AddColumn(oMasterConn, strTable, "micrprop_unadj", "DOUBLE", "");

                        if (!oAdo.ColumnExist(oMasterConn, strTable, "SUBPPROP_UNADJ"))
                            oAdo.AddColumn(oMasterConn, strTable, "subpprop_unadj", "DOUBLE", "");

                        if (!oAdo.ColumnExist(oMasterConn, strTable, "MACRPROP_UNADJ"))
                            oAdo.AddColumn(oMasterConn, strTable, "macrprop_unadj", "DOUBLE", "");
                    }

                }
            }

            if (oMasterConn != null && oMasterConn.State == System.Data.ConnectionState.Open)
                oAdo.CloseConnection(oMasterConn);
            
            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            oAdo = null;


        }

        /// <summary>
        /// Change these column names in the biosum_pop_stratum_adjustment_factors table:
        /// biosum_adj_factor_macr to pmh_macr
        /// biosum_adj_factor_subp to pmh_sub
        /// biosum_adj_factor_micr to pmh_micr
        /// biosum_adj_factor_cond to pmh_cond
        /// 
        /// </summary>
        private void Update_5_6_2()
        {
            int intTableType = -1;
            string strSQL;
            string strPath = "";
            string strFile = "";
            string strTable = "";
            System.Data.OleDb.OleDbConnection oMasterConn = null;



            //
            //MAIN PROJECT DATASOURCE
            //
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();



            
            ado_data_access oAdo = new ado_data_access();

            //open the project db file
            oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
                frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));
          
            //
            //BIOSUM CALCULATED ADJUSTMENT FACTORS DATASOURCE
            //
            intTableType = oDs.getDataSourceTableNameRow("BIOSUM POP STRATUM ADJUSTMENT FACTORS");
            if (intTableType > -1)
            {
                //even though listed in datasource table, the file and table could possibly not exist
                if (oDs.m_strDataSource[intTableType, Datasource.FILESTATUS] == "NF" ||
                    oDs.m_strDataSource[intTableType, Datasource.TABLESTATUS] == "NF")
                {
                    //okay, recreate the table in the master.mdb and change datasource values
                    
                    //change datasource back to default values
                    strSQL = "UPDATE datasource SET " +
                                 "Path='" + ReferenceProjectDirectory.Trim() + "\\db'," +
                                 "File='master.mdb'," +
                                 "table_name='" + frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName + "' " +
                             "WHERE TRIM(UCASE(table_type))='BIOSUM POP STRATUM ADJUSTMENT FACTORS'";
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                    //create the table
                    strPath = ReferenceProjectDirectory.Trim() + "\\db\\master.mdb";
                    oMasterConn = new System.Data.OleDb.OleDbConnection();
                    oAdo.OpenConnection(oAdo.getMDBConnString(strPath, "", ""), ref oMasterConn);
                    frmMain.g_oTables.m_oFIAPlot.CreateBiosumPopStratumAdjustmentFactorsTable(oAdo, oMasterConn, frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName);
                }
                else
                {
                    //table found so check the column names
                    //open connection to the db container
                    oMasterConn = new System.Data.OleDb.OleDbConnection();
                    strPath = oDs.m_strDataSource[intTableType,Datasource.PATH].Trim() + "\\" + oDs.m_strDataSource[intTableType,Datasource.MDBFILE].Trim();
                    strTable = oDs.m_strDataSource[intTableType, Datasource.TABLE].Trim();
                    oAdo.OpenConnection(oAdo.getMDBConnString(strPath, "", ""), ref oMasterConn);
                    //pmh_macr
                    if (!oAdo.ColumnExist(oMasterConn, strTable, "pmh_macr"))
                    {
                        oAdo.AddColumn(oMasterConn, strTable, "pmh_macr", "DECIMAL", "5,4");
                        //update with the old column name values
                        if (oAdo.ColumnExist(oMasterConn, strTable, "biosum_adj_factor_macr"))
                        {
                            strSQL = "UPDATE " + strTable + " SET " +
                                       "pmh_macr=IIF(biosum_adj_factor_macr IS NOT NULL,biosum_adj_factor_macr,null)";
                            oAdo.SqlNonQuery(oMasterConn,strSQL);
                            //drop the old column
                            strSQL = "ALTER TABLE " + strTable + " DROP COLUMN biosum_adj_factor_macr";
                            oAdo.SqlNonQuery(oMasterConn, strSQL);
                        }
                    }
                    //pmh_sub
                    if (!oAdo.ColumnExist(oMasterConn, strTable, "pmh_sub"))
                    {
                        oAdo.AddColumn(oMasterConn, strTable, "pmh_sub", "DECIMAL", "5,4");
                        //update with the old column name values
                        if (oAdo.ColumnExist(oMasterConn, strTable, "biosum_adj_factor_subp"))
                        {
                            strSQL = "UPDATE " + strTable + " SET " +
                                       "pmh_sub=IIF(biosum_adj_factor_subp IS NOT NULL,biosum_adj_factor_subp,null)";
                            oAdo.SqlNonQuery(oMasterConn, strSQL);
                            //drop the old column
                            strSQL = "ALTER TABLE " + strTable + " DROP COLUMN biosum_adj_factor_subp";
                            oAdo.SqlNonQuery(oMasterConn, strSQL);
                        }
                    }
                    //pmh_micr
                    if (!oAdo.ColumnExist(oMasterConn, strTable, "pmh_micr"))
                    {
                        oAdo.AddColumn(oMasterConn, strTable, "pmh_micr", "DECIMAL", "5,4");
                        //update with the old column name values
                        if (oAdo.ColumnExist(oMasterConn, strTable, "biosum_adj_factor_micr"))
                        {
                            strSQL = "UPDATE " + strTable + " SET " +
                                       "pmh_micr=IIF(biosum_adj_factor_micr IS NOT NULL,biosum_adj_factor_micr,null)";
                            oAdo.SqlNonQuery(oMasterConn, strSQL);
                            //drop the old column
                            strSQL = "ALTER TABLE " + strTable + " DROP COLUMN biosum_adj_factor_micr";
                            oAdo.SqlNonQuery(oMasterConn, strSQL);
                        }
                    }
                    //pmh_cond
                    if (!oAdo.ColumnExist(oMasterConn, strTable, "pmh_cond"))
                    {
                        oAdo.AddColumn(oMasterConn, strTable, "pmh_cond", "DECIMAL", "5,4");
                        //update with the old column name values
                        if (oAdo.ColumnExist(oMasterConn, strTable, "biosum_adj_factor_cond"))
                        {
                            strSQL = "UPDATE " + strTable + " SET " +
                                       "pmh_cond=IIF(biosum_adj_factor_cond IS NOT NULL,biosum_adj_factor_cond,null)";
                            oAdo.SqlNonQuery(oMasterConn, strSQL);
                            //drop the old column
                            strSQL = "ALTER TABLE " + strTable + " DROP COLUMN biosum_adj_factor_cond";
                            oAdo.SqlNonQuery(oMasterConn, strSQL);
                        }
                    }
                }

            }
            else
            {
                //does not exist
                //append the new datasource
                strSQL = "INSERT INTO datasource (table_type,Path,file,table_name) VALUES " +
                            "('BIOSUM Pop Stratum Adjustment Factors'," +
                            "'" + ReferenceProjectDirectory.Trim() + "\\db'," +
                            "'master.mdb'," +
                            "'" + frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName + "');";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                //create the new table
                strPath = ReferenceProjectDirectory + "\\db\\master.mdb";
                oMasterConn = new System.Data.OleDb.OleDbConnection();
                oAdo.OpenConnection(oAdo.getMDBConnString(strPath, "", ""), ref oMasterConn);
                frmMain.g_oTables.m_oFIAPlot.CreateBiosumPopStratumAdjustmentFactorsTable(oAdo, oMasterConn, frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName);
            }
            
           

            if (oMasterConn != null && oMasterConn.State == System.Data.ConnectionState.Open)
                oAdo.CloseConnection(oMasterConn);

            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            oAdo = null;


        }
        private void UpdateDatasources_5_7_0()
        {
            frmMain.g_sbpInfo.Text = "Version Update: Creating and Updating Placeholder column...Stand by";
            
            ado_data_access oAdo = new ado_data_access();

            string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\processor\\db";
            //open the scenario_processor_rule_definitions.mdb file
            oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioDir + "\\scenario_processor_rule_definitions.mdb", "", ""));

            //retrieve paths for all scenarios in the project and put them in list
            List<string> lstScenarioDb = new List<string>();
            oAdo.m_strSQL = "SELECT path from scenario";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    string strPath = "";
                    if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value)
                        strPath = oAdo.m_OleDbDataReader["path"].ToString().Trim();
                    if (!String.IsNullOrEmpty(strPath))
                    {
                        //Check to see if the .mdb exists before adding it to the list
                        string strPathToMdb = strPath + "\\db\\scenario_results.mdb";
                        //sample path: C:\\workspace\\BioSum\\biosum_data\\bluemountains\\processor\\scenario1\\db\\scenario_results.mdb
                        if (System.IO.File.Exists(strPathToMdb))
                            lstScenarioDb.Add(strPathToMdb);
                    }
                }
                oAdo.m_OleDbDataReader.Close();
            }

            // Loop through the scenario_results.mdb looking for tree_vol_val_by_species_diam_groups table
            string strPlaceholder = "place_holder";
            string strBcVolCf = "bc_vol_cf";
            string strBcWtGt = "bc_wt_gt";
            foreach (string strPath in lstScenarioDb)
            {
                // Add columns to tree_vol_val_by_species_diam_groups table
                oAdo.OpenConnection(oAdo.getMDBConnString(strPath, "", ""));
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, strPlaceholder))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, strPlaceholder, "CHAR", "1", "N");

                    // Set place_holder field to 'N' for new column
                    oAdo.m_strSQL = "UPDATE " + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + " " +
                        "SET place_holder = IIF(place_holder IS NULL,'N',place_holder)";
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                }
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, strBcVolCf))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, strBcVolCf, "DOUBLE", "");
                }
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, strBcWtGt))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, strBcWtGt, "DOUBLE", "");
                }

                // Add placeholder column to harvest_costs table
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, strPlaceholder))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, strPlaceholder, "CHAR", "1", "N");

                    // Set place_holder field to 'N' for new column
                    oAdo.m_strSQL = "UPDATE " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + " " +
                        "SET place_holder = IIF(place_holder IS NULL,'N',place_holder)";
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                }
            }

            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            oAdo = null;

        }
        private void Update_5_7_0()
        {
            frmMain.g_sbpInfo.Text = "Version Update: Updating to version 5.7.0...Stand by";
            string strTreeFile = "";
            string strTreeTable = "";
            int intTreeTableType;





            //
            //MAIN PROJECT DATASOURCE
            //
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();

            ado_data_access oAdo = new ado_data_access();

            intTreeTableType = oDs.getDataSourceTableNameRow("TREE");

            if (intTreeTableType > -1)
            {
                //even though listed in datasource table, the file and table could possibly not exist
                if (oDs.m_strDataSource[intTreeTableType, Datasource.FILESTATUS] == "F" &&
                    oDs.m_strDataSource[intTreeTableType, Datasource.TABLESTATUS] == "F")
                {
                    strTreeTable = oDs.m_strDataSource[intTreeTableType, Datasource.TABLE].Trim();

                    strTreeFile = oDs.m_strDataSource[intTreeTableType, Datasource.PATH].Trim() + "\\" +
                                oDs.m_strDataSource[intTreeTableType, Datasource.MDBFILE].Trim();


                    //open the project db file
                    oAdo.OpenConnection(oAdo.getMDBConnString(strTreeFile, "", ""));
                    if (oAdo.ColumnExist(oAdo.m_OleDbConnection, strTreeTable, "voltsgrs") == false)
                    {
                        oAdo.AddColumn(oAdo.m_OleDbConnection, strTreeTable, "voltsgrs", "DOUBLE", "");
                    }
                    oAdo.CloseConnection(oAdo.m_OleDbConnection);
                }
            }
            //
            //BACKUP AND DELETE THE CUTLIST TABLES
            //
            System.DateTime oDate = System.DateTime.Now;
            string strDateFormat = "yyyy-MM-dd_HH-mm-ss";
            string strFileDate = oDate.ToString(strDateFormat);
            strFileDate = strFileDate.Replace("/", "_"); strFileDate = strFileDate.Replace(":", "_");
            List<string> strFiles = new List<string>();
            string strFile = "";
            string strFolder = "";
            RxTools oRxTools = new RxTools();
            Queries oQueries = new Queries();
            RxPackageItem_Collection oRxPackageItemCollection = new RxPackageItem_Collection();
            oQueries.m_oFvs.LoadDatasource = true;
            oQueries.m_oFIAPlot.LoadDatasource = true;
            oQueries.LoadDatasources(true);
            //
            //open the work file containing the links
            //
            oAdo.OpenConnection(oAdo.getMDBConnString(oQueries.m_strTempDbFile, "", ""));
            //
            //get all the variants
            //
            string strVariantsList = oRxTools.GetListOfFVSVariantsInPlotTable(oAdo, oAdo.m_OleDbConnection, oQueries.m_oFIAPlot.m_strPlotTable);
            string[] strVariantsArray = frmMain.g_oUtils.ConvertListToArray(strVariantsList, ",");
            //
            //get all the packages
            //
            oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(oAdo, oAdo.m_OleDbConnection, oQueries, oRxPackageItemCollection);
            //
            //close the work file containing the links
            //
            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            for (int x = 0; x <= strVariantsArray.Length - 1; x++)
            {
                strFolder = ReferenceProjectDirectory + "\\FVS\\Data\\" + strVariantsArray[x].Trim() + "\\BiosumCalc";

                for (int y = 0; y <= oRxPackageItemCollection.Count - 1; y++)
                {
                    strFile = strFolder + "\\" + strVariantsArray[x].Trim() + "_P" + oRxPackageItemCollection.Item(y).RxPackageId + "_TREE_CUTLIST.MDB";
                    if (System.IO.File.Exists(strFile))
                    {
                        strFiles.Add(strFile);
                    }
                }



            }

            if (strFiles != null && strFiles.Count > 0)
            {

                if (strFiles[0] != null & strFiles[0].Trim().Length > 0)
                {

                    //
                    //rename the old fvs cutlist files
                    //
                    for (int x = 0; x <= strFiles.Count - 1; x++)
                    {
                        if (strFiles[x] != null && strFiles[x].Trim().Length > 0)
                        {
                            System.IO.File.Copy(strFiles[x].Trim(), strFiles[x].Trim() + "_VersionControlUpgrade_" + strFileDate);
                            System.IO.File.Delete(strFiles[x].Trim());
                        }

                    }
                }
            }
        }

        private void UpdateDatasources_5_7_7()
        {
            frmMain.g_sbpInfo.Text = "Version Update: Copying new harvest methods table to project...Stand by";

            // Load project data sources table
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();

            // Extract table properties from data sources table; Assume still under the old name
            int intHarvestMethodTable = oDs.getValidTableNameRow("FRCS System Harvest Method");
            string strDirectoryPath = oDs.m_strDataSource[intHarvestMethodTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = oDs.m_strDataSource[intHarvestMethodTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            string strFileStatus = oDs.m_strDataSource[intHarvestMethodTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            string strSourceTableName = oDs.m_strDataSource[intHarvestMethodTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strTableStatus = oDs.m_strDataSource[intHarvestMethodTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();

            ado_data_access oAdo = new ado_data_access();

            // Create table link of the application db ref_master harvest method table and tree species table
            if (strFileStatus == "F" && strTableStatus == "F")
            {
                dao_data_access oDao = new dao_data_access();
                string strDestinationDbFile = strDirectoryPath + "\\" + strFileName;
                string strDestinationTableName = "harvestmethod_worktable";
                string strDestinationSpeciesTableName = "tree_species_577";
                string strSourceDbFile=frmMain.g_oEnv. strAppDir.Trim() + "\\db\\ref_master.mdb";

                // Harvest Methods table
                oDao.CreateTableLink(strDestinationDbFile, strDestinationTableName, strSourceDbFile, Tables.Reference.DefaultHarvestMethodsTableName);
                // Tree Species table
                oDao.CreateTableLink(strDestinationDbFile, strDestinationSpeciesTableName + "_worktable", strSourceDbFile, Tables.Reference.DefaultTreeSpeciesTableName);
                oDao.m_DaoWorkspace.Close();
                

                //open connection to destination database
                oAdo.OpenConnection(oAdo.getMDBConnString(strDestinationDbFile, "", ""));
                //drop existing harvest methods table
                string strSql = "DROP TABLE " + strSourceTableName;
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSql);
                //copy contents of new harvest methods table into place
                strSql = "SELECT * INTO " + Tables.Reference.DefaultHarvestMethodsTableName + " FROM " + strDestinationTableName;
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSql);
                //copy contents of new tree species table into place
                strSql = "SELECT * INTO " + strDestinationSpeciesTableName + " FROM " + strDestinationSpeciesTableName + "_worktable";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSql);

                //drop the harvest methods table link
                if (oAdo.TableExist(oAdo.m_OleDbConnection, strDestinationTableName))
                {
                    strSql = "DROP TABLE " + strDestinationTableName;
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSql);
                }
                //drop the tree species table link
                if (oAdo.TableExist(oAdo.m_OleDbConnection, strDestinationSpeciesTableName))
                {
                    strSql = "DROP TABLE " + strDestinationSpeciesTableName + "_worktable";
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSql);
                }


                //open connection to DATASOURce database
                oAdo.OpenConnection(oAdo.getMDBConnString(oDs.m_strDataSourceMDBFile, "", ""));
                strSql = "UPDATE " + oDs.m_strDataSourceTableName + " SET table_type = 'Harvest Methods', " +
                         "table_name = '" + Tables.Reference.DefaultHarvestMethodsTableName + "' " +
                         "WHERE TRIM(table_type) = 'FRCS System Harvest Method'";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSql);
                
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
            }


            frmMain.g_sbpInfo.Text = "Version Update: Creating 3 new columns in scenario_harvest_method table...Stand by";

            string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\processor\\db";
            //open the scenario_processor_rule_definitions.mdb file
            oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioDir + "\\scenario_processor_rule_definitions.mdb", "", ""));

            //update the woodland as percent of total
            string strWoodlandMarchPctCol = "WoodlandMerchAsPercentOfTotalVol";
            if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName, strWoodlandMarchPctCol))
            {
                oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName, strWoodlandMarchPctCol, "INTEGER", "");

                // Set adj factor to 80 for new column
                oAdo.m_strSQL = "UPDATE " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName + " " +
                    "SET " + strWoodlandMarchPctCol + " = 80";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            }

            string strSaplingMerchPctCol = "SaplingMerchAsPercentOfTotalVol";
            if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName, strSaplingMerchPctCol))
            {
                oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName, strSaplingMerchPctCol, "INTEGER", "");

                // Set adj factor to 80 for new column
                oAdo.m_strSQL = "UPDATE " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName + " " +
                    "SET " + strSaplingMerchPctCol + " = 60";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            }

            string strCullPctThreshold = "CullPctThreshold";
            if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName, strCullPctThreshold))
            {
                oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName, strCullPctThreshold, "INTEGER", "");

                // Set adj factor to 80 for new column
                oAdo.m_strSQL = "UPDATE " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName + " " +
                    "SET " + strCullPctThreshold + " = 50";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            }
            //rename the FRCS System Harvest Method table in scenario_processor_rule_definitions.mdb
            oAdo.m_strSQL = "UPDATE " + Tables.Scenario.DefaultScenarioDatasourceTableName + " SET table_type = 'Harvest Methods', " +
                            "table_name = '" + Tables.Reference.DefaultHarvestMethodsTableName + "' " +
                            "WHERE TRIM(table_type) = 'FRCS System Harvest Method'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            // Add new move-in costs table in scenario_processor_rule_definitions.mdb if it is missing
            if (!oAdo.TableExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName))
            {
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioMoveInCostsTable(oAdo, oAdo.m_OleDbConnection,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName);
            }

            frmMain.g_sbpInfo.Text = "Version Update: Creating stand_residue_wt_gt column in tree vol val table(s)...Stand by";

            //retrieve paths for all scenarios in the project and put them in list
            List<string> lstScenarioDb = new List<string>();
            oAdo.m_strSQL = "SELECT path from scenario";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    string strPath = "";
                    if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value)
                        strPath = oAdo.m_OleDbDataReader["path"].ToString().Trim();
                    if (!String.IsNullOrEmpty(strPath))
                    {
                        //Check to see if the .mdb exists before adding it to the list
                        string strPathToMdb = strPath + "\\db\\scenario_results.mdb";
                        //sample path: C:\\workspace\\BioSum\\biosum_data\\bluemountains\\processor\\scenario1\\db\\scenario_results.mdb
                        if (System.IO.File.Exists(strPathToMdb))
                            lstScenarioDb.Add(strPathToMdb);
                    }
                }
                oAdo.m_OleDbDataReader.Close();
            }

            // Loop through the scenario_results.mdb looking for tree_vol_val_by_species_diam_groups table
            string strStandResidueWtGt = "stand_residue_wt_gt";
            foreach (string strPath in lstScenarioDb)
            {
                // Add columns to tree_vol_val_by_species_diam_groups table
                oAdo.OpenConnection(oAdo.getMDBConnString(strPath, "", ""));
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, strStandResidueWtGt))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, strStandResidueWtGt, "DOUBLE", "");
                }
            }

            oAdo.CloseConnection(oAdo.m_OleDbConnection);
        }

        private void UpdateDatasources_5_7_8()
        {
            frmMain.g_sbpInfo.Text = "Version Update: Adding new columns to harvest costs tables...Stand by";

            ado_data_access oAdo = new ado_data_access();

            string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\processor\\db";
            //open the scenario_processor_rule_definitions.mdb file
            oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioDir + "\\scenario_processor_rule_definitions.mdb", "", ""));
            
            //retrieve paths for all scenarios in the project and put them in list
            List<string> lstScenarioDb = new List<string>();
            oAdo.m_strSQL = "SELECT path from scenario";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    string strPath = "";
                    if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value)
                        strPath = oAdo.m_OleDbDataReader["path"].ToString().Trim();
                    if (!String.IsNullOrEmpty(strPath))
                    {
                        //Check to see if the .mdb exists before adding it to the list
                        string strPathToMdb = strPath + "\\db\\scenario_results.mdb";
                        //sample path: C:\\workspace\\BioSum\\biosum_data\\bluemountains\\processor\\scenario1\\db\\scenario_results.mdb
                        if (System.IO.File.Exists(strPathToMdb))
                            lstScenarioDb.Add(strPathToMdb);
                    }
                }
                oAdo.m_OleDbDataReader.Close();
            }
            // Loop through the scenario_results.mdb looking for harvest_costs table
            foreach (string strPath in lstScenarioDb)
            {
                // Add columns to harvest_costs table
                oAdo.OpenConnection(oAdo.getMDBConnString(strPath, "", ""));
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "chip_cpa"))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "chip_cpa", "DOUBLE", "");
                }
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "assumed_movein_cpa"))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "assumed_movein_cpa", "DOUBLE", "");
                }
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "ideal_complete_cpa"))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "ideal_complete_cpa", "DOUBLE", "");
                }
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "ideal_harvest_cpa"))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "ideal_harvest_cpa", "DOUBLE", "");
                }
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "ideal_chip_cpa"))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "ideal_chip_cpa", "DOUBLE", "");
                }
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "ideal_assumed_movein_cpa"))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "ideal_assumed_movein_cpa", "DOUBLE", "");
                }
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "override_YN"))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, "override_YN", "CHAR", "1", "N");
                }
            }

            frmMain.g_sbpInfo.Text = "Version Update: Making modifications to scenario_harvest_method table ...Stand by";

            //open the scenario_processor_rule_definitions.mdb file
            oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioDir + "\\scenario_processor_rule_definitions.mdb", "", ""));
            
            if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName, "HarvestMethodSelection"))
            {
                oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName, "HarvestMethodSelection", "CHAR", "15");

                // Populate column from old UseRxDefaultHarvestMethodYN column
                oAdo.m_strSQL = "UPDATE " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName + " " +
                                "SET HarvestMethodSelection = IIF(UseRxDefaultHarvestMethodYN = 'Y', 'RX','SPECIFIED')";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            }

            // Drop the old UseRxDefaultHarvestMethodYN column
            if (oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName, "UseRxDefaultHarvestMethodYN"))
            {
               oAdo.m_strSQL = "ALTER TABLE " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName + " DROP COLUMN UseRxDefaultHarvestMethodYN";
               oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            }

            oAdo.CloseConnection(oAdo.m_OleDbConnection);
        }

        private void UpdateDatasources_5_7_9()
        {
            frmMain.g_sbpInfo.Text = "Version Update: Moving tree groupings to Processor scenario database...Stand by";
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();
            string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\processor\\db";
            
            //open the scenario_processor_rule_definitions.mdb file so we can add the new tree groupings tables
            oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioDir + "\\scenario_processor_rule_definitions.mdb", "", ""));

            // SCENARIO_TREE_DIAM_GROUPS
            if (!oAdo.TableExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName))
            {
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioTreeDiamGroupsTable(oAdo, oAdo.m_OleDbConnection,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName);
            }
            // SCENARIO_TREE_SPECIES_GROUPS_LIST
            if (!oAdo.TableExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName))
            {
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioTreeSpeciesGroupsListTable(oAdo, oAdo.m_OleDbConnection,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName);
            }
            // SCENARIO_TREE_SPECIES_GROUPS
            if (!oAdo.TableExist(oAdo.m_OleDbConnection, Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName))
            {
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioTreeSpeciesGroupsTable(oAdo, oAdo.m_OleDbConnection,
                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName);
            }

            //retrieve paths for all scenarios in the project and put them in list
            List<string> lstScenarioDb = new List<string>();
            oAdo.m_strSQL = "SELECT path from scenario";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    string strPath = "";
                    if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value)
                        strPath = oAdo.m_OleDbDataReader["path"].ToString().Trim();
                    if (!String.IsNullOrEmpty(strPath))
                    {
                         if (System.IO.Directory.Exists(strPath))
                            lstScenarioDb.Add(strPath);
                    }
                }
            }

            // Loop through the scenario_results.mdb transferring data for each existing scenario

            foreach (string strPath in lstScenarioDb)
            {
                string[] strSplit = strPath.Split(new Char[] {'\\'});
                string strScenarioId = strSplit[strSplit.Length - 1];
                string strTreeDiamGroupsPath = "";
                string strTreeDiamGroupsMdb = "";
                string strTreeDiamGroupsTable = "";
                string strSpeciesGroupsPath = "";
                string strSpeciesGroupsMdb = "";
                string strSpeciesGroupsTable = "";
                // Unlike other old tables, TREE_SPECIES_GROUPS_LIST is not stored in scenario_datasource table
                string strSpeciesGroupsListPath = ReferenceProjectDirectory.Trim() + "\\db\\master.mdb";
                string strSpeciesGroupsListTable = "TREE_SPECIES_GROUPS_LIST";
                string strTreeSpeciesPath = "";
                string strTreeSpeciesMdb = "";
                string strTreeSpeciesTable = "";

                
                // Get paths to old tables from scenario_datasource table
                oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioDir + "\\scenario_processor_rule_definitions.mdb", "", ""));
                string strSQL = "SELECT * FROM " + Tables.Scenario.DefaultScenarioDatasourceTableName +
                    " WHERE TRIM(UCASE(scenario_id))='" + strScenarioId.Trim().ToUpper() + "'";
                oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
                if (oAdo.m_OleDbDataReader.HasRows)
                {
                    while (oAdo.m_OleDbDataReader.Read())
                    {
                        string strTableType = oAdo.m_OleDbDataReader["table_type"].ToString().Trim();
                        switch (strTableType.ToUpper())
                        {
                            case "TREE DIAMETER GROUPS":
                                strTreeDiamGroupsPath = oAdo.m_OleDbDataReader["path"].ToString().Trim();
                                strTreeDiamGroupsMdb = oAdo.m_OleDbDataReader["file"].ToString().Trim();
                                strTreeDiamGroupsTable = oAdo.m_OleDbDataReader["table_name"].ToString().Trim();
                                break;
                            case "TREE SPECIES":
                                strTreeSpeciesPath = oAdo.m_OleDbDataReader["path"].ToString().Trim();
                                strTreeSpeciesMdb = oAdo.m_OleDbDataReader["file"].ToString().Trim();
                                strTreeSpeciesTable = oAdo.m_OleDbDataReader["table_name"].ToString().Trim();
                                break;
                            case "TREE SPECIES GROUPS":
                                strSpeciesGroupsPath = oAdo.m_OleDbDataReader["path"].ToString().Trim();
                                strSpeciesGroupsMdb = oAdo.m_OleDbDataReader["file"].ToString().Trim();
                                strSpeciesGroupsTable = oAdo.m_OleDbDataReader["table_name"].ToString().Trim();
                                break;
                        }
                    }
                }

                // Rename USER_SPC_GROUP column to USER_SPC_GROUP_578 to avoid dao errors later
                string strUserSpcGroup = "USER_SPC_GROUP_578";
                oAdo.OpenConnection(m_oAdo.getMDBConnString(strTreeSpeciesPath + "\\" + strTreeSpeciesMdb, "", ""));
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, strTreeSpeciesTable, strUserSpcGroup))
                {
                    oDao.RenameField(strTreeSpeciesPath + "\\" + strTreeSpeciesMdb, strTreeSpeciesTable, "USER_SPC_GROUP", strUserSpcGroup);
                }

                // Read tree diameter groups into memory so we can transfer them to the new table
                oAdo.OpenConnection(oAdo.getMDBConnString(strTreeDiamGroupsPath + "\\" + strTreeDiamGroupsMdb, "", ""));
                if (oAdo.TableExist(oAdo.m_OleDbConnection, strTreeDiamGroupsTable))
                {
                    ProcessorScenarioItem.TreeDiamGroupsItem_Collection _objTreeDiamCollection = new ProcessorScenarioItem.TreeDiamGroupsItem_Collection();

                    strSQL = "SELECT * FROM " + strTreeDiamGroupsTable + " ORDER BY MIN_DIAM";
                    oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
                    if (oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (oAdo.m_OleDbDataReader.Read())
                        {
                            ProcessorScenarioItem.TreeDiamGroupsItem _objTreeDiamItem = new ProcessorScenarioItem.TreeDiamGroupsItem();
                            _objTreeDiamItem.DiamGroup = Convert.ToString(oAdo.m_OleDbDataReader["diam_group"]).Trim();
                            _objTreeDiamItem.DiamClass = Convert.ToString(oAdo.m_OleDbDataReader["diam_class"]).Trim();
                            _objTreeDiamItem.MinDiam = Convert.ToString(oAdo.m_OleDbDataReader["min_diam"]).Trim();
                            _objTreeDiamItem.MaxDiam = Convert.ToString(oAdo.m_OleDbDataReader["max_diam"]).Trim();
                            _objTreeDiamCollection.Add(_objTreeDiamItem);
                        }
                    }

                    // Switch connection back to the scenario db so we can write the diameter groups
                    oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioDir + "\\scenario_processor_rule_definitions.mdb", "", ""));

                    for (int x = 0; x <= _objTreeDiamCollection.Count - 1; x++)
                    {
                        FIA_Biosum_Manager.ProcessorScenarioItem.TreeDiamGroupsItem oItem = _objTreeDiamCollection.Item(x);
                        string strId = oItem.DiamGroup;
                        string strMin = oItem.MinDiam;
                        string strMax = oItem.MaxDiam;
                        string strDef = oItem.DiamClass;

                        oAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName + " " +
                            "(diam_group,diam_class,min_diam,max_diam,scenario_id) VALUES " +
                            "(" + strId + ",'" + strDef.Trim() + "'," +
                            strMin + "," + strMax + ",'" + strScenarioId.Trim() + "')";
                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Unable to locate the previous TREE_DIAM_GROUPS table " +
                                     "in " + strTreeDiamGroupsPath + "\\" + strTreeDiamGroupsMdb +
                                     ". The Tree Diameter Groups could not be transferred.",
                                     "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                                     System.Windows.Forms.MessageBoxIcon.Error);

                }
                // Add link for tree_species table to tree_species_group_list directory to prep for querying
                string strLinkedTreeSpeciesTable = "tree_species_worktable";
                oDao.CreateTableLink(strSpeciesGroupsListPath, strLinkedTreeSpeciesTable,
                    strTreeSpeciesPath + "\\" + strTreeSpeciesMdb, strTreeSpeciesTable);

                // Read tree diameter groups into memory so we can transfer them to the new table
                oAdo.OpenConnection(oAdo.getMDBConnString(strSpeciesGroupsListPath, "", ""));
                if (oAdo.TableExist(oAdo.m_OleDbConnection, strSpeciesGroupsListTable))
                {
                    ProcessorScenarioItem.SpcGroupListItemCollection _objSpcGroupListItemCollection = new ProcessorScenarioItem.SpcGroupListItemCollection();

                    strSQL = "SELECT DISTINCT l.species_group,  l.common_name, t.spcd " + 
                        "FROM " + strSpeciesGroupsListTable + " l, " + strLinkedTreeSpeciesTable + " t " +
                        "WHERE t." + strUserSpcGroup + " = l.species_group AND TRIM(l.common_name) = TRIM(t.common_name)";
                    oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
                    if (oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (oAdo.m_OleDbDataReader.Read())
                        {
                            ProcessorScenarioItem.SpcGroupListItem _objSpcGroupListItem = new ProcessorScenarioItem.SpcGroupListItem();
                            _objSpcGroupListItem.CommonName = Convert.ToString(oAdo.m_OleDbDataReader["common_name"]).Trim();
                            _objSpcGroupListItem.SpeciesGroup = Convert.ToInt32(oAdo.m_OleDbDataReader["species_group"]);
                            _objSpcGroupListItem.SpeciesCode = Convert.ToInt32(oAdo.m_OleDbDataReader["spcd"]);
                            _objSpcGroupListItemCollection.Add(_objSpcGroupListItem);
                        }
                    }

                    // Switch connection back to the scenario db so we can write the diameter groups
                    oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioDir + "\\scenario_processor_rule_definitions.mdb", "", ""));

                    for (int x = 0; x <= _objSpcGroupListItemCollection.Count - 1; x++)
                    {
                        FIA_Biosum_Manager.ProcessorScenarioItem.SpcGroupListItem oItem = _objSpcGroupListItemCollection.Item(x);

                        oAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName + " " +
                                        "(SPECIES_GROUP,common_name,SCENARIO_ID,SPCD) VALUES " +
                                        "(" + oItem.SpeciesGroup + ",'" + oItem.CommonName + "','" + strScenarioId + "', " +
                                        oItem.SpeciesCode + " )";
                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Unable to locate the previous TREE_SPECIES_GROUPS_LIST table " +
                                     "in " + strSpeciesGroupsListPath +
                                     ". The Tree Species Groups could not be transferred.",
                                     "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                                     System.Windows.Forms.MessageBoxIcon.Error);

                }
                //drop the tree species table link
                oAdo.OpenConnection(oAdo.getMDBConnString(strSpeciesGroupsListPath, "", ""));
                if (oAdo.TableExist(oAdo.m_OleDbConnection, strLinkedTreeSpeciesTable))
                {
                    strSQL = "DROP TABLE " + strLinkedTreeSpeciesTable;
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                }

                // Read tree species groups into memory so we can transfer them to the new table
                oAdo.OpenConnection(oAdo.getMDBConnString(strSpeciesGroupsPath + "\\" + strSpeciesGroupsMdb, "", ""));
                if (oAdo.TableExist(oAdo.m_OleDbConnection, strSpeciesGroupsTable))
                {
                    ProcessorScenarioItem.SpcGroupItemCollection _objSpcGroupCollection = new ProcessorScenarioItem.SpcGroupItemCollection();

                    strSQL = "SELECT * FROM " + strSpeciesGroupsTable;
                    oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
                    if (oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (oAdo.m_OleDbDataReader.Read())
                        {
                            ProcessorScenarioItem.SpcGroupItem _objSpcGroupItem = new ProcessorScenarioItem.SpcGroupItem();
                            _objSpcGroupItem.SpeciesGroup = Convert.ToInt32(oAdo.m_OleDbDataReader["species_group"]);
                            _objSpcGroupItem.SpeciesGroupLabel = Convert.ToString(oAdo.m_OleDbDataReader["species_label"]).Trim();
                            _objSpcGroupCollection.Add(_objSpcGroupItem);
                        }
                    }

                    // Switch connection back to the scenario db so we can write the diameter groups
                    oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioDir + "\\scenario_processor_rule_definitions.mdb", "", ""));

                    for (int x = 0; x <= _objSpcGroupCollection.Count - 1; x++)
                    {
                        FIA_Biosum_Manager.ProcessorScenarioItem.SpcGroupItem oItem = _objSpcGroupCollection.Item(x);  
                        oAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName + " " +
                                       "(SPECIES_GROUP,SPECIES_LABEL,SCENARIO_ID) VALUES " +
                                       "(" + oItem.SpeciesGroup + ",'" + oItem.SpeciesGroupLabel + "','" + strScenarioId.Trim() + "')";
                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Unable to locate the previous TREE_SPECIES_GROUPS table " +
                                     "in " + strSpeciesGroupsPath + "\\" + strSpeciesGroupsMdb +
                                     ". The Tree Species Groups could not be transferred.",
                                     "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                                     System.Windows.Forms.MessageBoxIcon.Error);

                }
            }

            frmMain.g_sbpInfo.Text = "Version Update: Update Core Analysis data sources table...Stand by";
            string strCoreMdb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
                Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioDatasourceTableDbFile;
            oAdo.OpenConnection(oAdo.getMDBConnString(strCoreMdb, "", ""));
            oAdo.m_strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioDatasourceTableName +
                " WHERE TRIM(UCASE(table_type)) = 'TREE DIAMETER GROUPS' OR" +
                " TRIM(UCASE(table_type)) = 'TREE SPECIES GROUPS'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            frmMain.g_sbpInfo.Text = "Version Update: Update Project data sources table...Stand by";
            string strProjectMdb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
                Tables.Project.DefaultProjectDatasourceTableDbFile;
            oAdo.OpenConnection(oAdo.getMDBConnString(strProjectMdb, "", ""));
            oAdo.m_strSQL = "DELETE * FROM " + Tables.Project.DefaultProjectDatasourceTableName +
                " WHERE TRIM(UCASE(table_type)) = 'TREE DIAMETER GROUPS' OR" +
                " TRIM(UCASE(table_type)) = 'TREE SPECIES GROUPS' OR" +
                " TRIM(UCASE(table_type)) = 'TREE SPECIES GROUPS LIST'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            
            if (oAdo != null)
            {
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }

            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
        }

        private void UpdateDatasources_5_8_0()
        {
            frmMain.g_sbpInfo.Text = "Version Update: Renaming obsolete tree species diameter and groups tables ...Stand by";
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();

            // Query for the paths/tables from scenario_datasource that we need to rename
            string strTableSuffix = "_ver_control_" + DateTime.Now.ToString("MMddyyyy");
            string strScenarioMdb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\processor\\db\\scenario_processor_rule_definitions.mdb";
            string strSQL = "SELECT distinct path, file, table_name " +
                            "FROM scenario_datasource " +
                            "WHERE table_type in ('TREE DIAMETER GROUPS','TREE SPECIES GROUPS')";
            oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMdb, "", ""));
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    string strPathAndFile = Convert.ToString(oAdo.m_OleDbDataReader["path"]).Trim() +
                                            "\\" + Convert.ToString(oAdo.m_OleDbDataReader["file"]).Trim();
                    string strTable = Convert.ToString(oAdo.m_OleDbDataReader["table_name"]).Trim();
                    if (oDao.TableExists(strPathAndFile, strTable))
                    {
                        oDao.RenameTable(strPathAndFile, strTable, strTable + strTableSuffix, true, false);
                    }
                }
            }

            // Delete entries from scenario_datasource after renaming tables
            oAdo.m_strSQL = "DELETE FROM " + Tables.Scenario.DefaultScenarioDatasourceTableName +
                     " WHERE TRIM(UCASE(table_type)) = 'TREE DIAMETER GROUPS' OR" +
                     " TRIM(UCASE(table_type)) = 'TREE SPECIES GROUPS'";

            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            // Rename tree groups list table in master.mdb if it exists; It isn't managed in the scenario_datasource table
            string strTargetMdb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\master.mdb";
            string strTargetTable = "tree_species_groups_list";
            if (oDao.TableExists(strTargetMdb, strTargetTable))
            {
                oDao.RenameTable(strTargetMdb, strTargetTable, strTargetTable + strTableSuffix, true, false);
            }

            //rename and replace ref_master.mdb tree_species table
            frmMain.g_sbpInfo.Text = "Version Update: Rename and replace tree species table ...Stand by";

            // Load project data sources table
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();

            // Extract table properties from data sources table; Assume still under the old name
            int intTreeSpeciesTable = oDs.getValidTableNameRow("Tree Species");
            string strDirectoryPath = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            string strFileStatus = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            strTargetTable = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strTableStatus = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();

            strTargetMdb = strDirectoryPath + "\\" + strFileName;
            if (strFileStatus == "F" && strTableStatus == "F")
            {
                oDao.RenameTable(strTargetMdb, strTargetTable, strTargetTable + strTableSuffix, true, false);
            }

            //Rename and copy new tree species table into place
            string strSourceDbFile = frmMain.g_oEnv.strAppDir.Trim() + "\\db\\ref_master.mdb";
            oDao.CreateTableLink(strTargetMdb, strTargetTable + "_worktable", strSourceDbFile, strTargetTable);

            //copy contents of new tree species table into place
            strSQL = "SELECT * INTO " + strTargetTable + " FROM " + strTargetTable + "_worktable";
            oAdo.OpenConnection(oAdo.getMDBConnString(strTargetMdb, "", ""));
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);

            //drop the tree species table link
            if (oAdo.TableExist(oAdo.m_OleDbConnection, strTargetTable + "_worktable"))
            {
                strSQL = "DROP TABLE " + strTargetTable + "_worktable";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
            }

            //rename fvs_tree_species table and re-map to %appData%
            frmMain.g_sbpInfo.Text = "Version Update: Rename and remap fvs tree species table ...Stand by";

            int intFvsTreeSpeciesTable = oDs.getValidTableNameRow("FVS Tree Species");
            strDirectoryPath = oDs.m_strDataSource[intFvsTreeSpeciesTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strFileName = oDs.m_strDataSource[intFvsTreeSpeciesTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            strFileStatus = oDs.m_strDataSource[intFvsTreeSpeciesTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            strTargetTable = oDs.m_strDataSource[intFvsTreeSpeciesTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            strTableStatus = oDs.m_strDataSource[intFvsTreeSpeciesTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
            strTargetMdb = strDirectoryPath + "\\" + strFileName;
            if (strFileStatus == "F" && strTableStatus == "F")
            {
                oDao.RenameTable(strTargetMdb, strTargetTable, strTargetTable + strTableSuffix, true, false);
            }

            string strDataSourceMdb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\project.mdb";
            oAdo.OpenConnection(oAdo.getMDBConnString(strDataSourceMdb, "", ""));
            strSQL = "UPDATE datasource " +
                     "SET PATH = '@@appdata@@\\fiabiosum', file = '" + Tables.Reference.DefaultBiosumReferenceDbFile + "' " +
                     "WHERE TABLE_TYPE = '" + Datasource.TableTypes.FvsTreeSpecies + "'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);

            //new datasource table entries for fia_tree_species_ref table
            frmMain.g_sbpInfo.Text = "Version Update: Add datasource entries for new FIA Tree Species Reference table ...Stand by";

            int intTreeSpeciesRef = oDs.getValidTableNameRow(Datasource.TableTypes.FiaTreeSpeciesReference);
            if (intTreeSpeciesRef < 1)
            {
                strSQL = "INSERT INTO datasource " +
                         "(table_type,Path,file,table_name) " +
                         "VALUES ('" + Datasource.TableTypes.FiaTreeSpeciesReference + "','@@appdata@@\\fiabiosum', " +
                         "'" + Tables.Reference.DefaultBiosumReferenceDbFile + "', '" +
                         Tables.ProcessorScenarioRun.DefaultFiaTreeSpeciesRefTableName + "')";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);

                // Refresh datasource array after change
                oDs.populate_datasource_array();
            }

            //new datasource table entries for each scenario
            //retrieve paths for all scenarios in the project and put them in list
            string strProcessorMdb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\processor\\db\\scenario_processor_rule_definitions.mdb";
            oAdo.OpenConnection(oAdo.getMDBConnString(strProcessorMdb,"",""));
            oAdo.m_strSQL = "SELECT distinct scenario_id from scenario";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    string strScenario = "";
                    if (oAdo.m_OleDbDataReader["scenario_id"] != System.DBNull.Value)
                        strScenario = oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim();
                    if (!String.IsNullOrEmpty(strScenario))
                    {
                        // Load scenario data sources table
                        FIA_Biosum_Manager.Datasource oScenarioDs = new Datasource();
                        oDs.m_strDataSourceMDBFile = strProcessorMdb;
                        oDs.m_strDataSourceTableName = "scenario_datasource";
                        oDs.m_strScenarioId = strScenario;
                        oDs.LoadTableColumnNamesAndDataTypes = false;
                        oDs.LoadTableRecordCount = false;
                        oDs.populate_datasource_array();
                        int intFiaSpeciesRef = oDs.getValidTableNameRow(Datasource.TableTypes.FiaTreeSpeciesReference);
                        if (intFiaSpeciesRef < 1)
                        {

                            strSQL = "INSERT INTO scenario_datasource (table_type, path, file, table_name, scenario_id) " +
                                     "VALUES ('" + Datasource.TableTypes.FiaTreeSpeciesReference + "','@@appdata@@\\fiabiosum', " +
                                     "'" + Tables.Reference.DefaultBiosumReferenceDbFile + "', '" +
                                     Tables.ProcessorScenarioRun.DefaultFiaTreeSpeciesRefTableName + "', '" + strScenario + "')";
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                        }
                    }
                }
                oAdo.m_OleDbDataReader.Close();
            }

            if (oAdo != null)
            {
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }

            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
        }

        private void UpdateDatasources_5_8_4()
        {
            frmMain.g_sbpInfo.Text = "Version Update: Updating Core Analysis Configurations ...Stand by";
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();
            string strCoreMdb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\core\\db\\scenario_core_rule_definitions.mdb";            
            oAdo.OpenConnection(oAdo.getMDBConnString(strCoreMdb, "", ""));
            oAdo.m_strSQL = "UPDATE scenario_fvs_variables_tiebreaker " +
                            "SET tiebreaker_method = 'Last Tie-Break Rank' " +
                            "WHERE tiebreaker_method = 'Treatment Intensity'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            oAdo.m_strSQL = "UPDATE scenario_fvs_variables_tiebreaker " +
                            "SET tiebreaker_method = 'Stand Attribute' " +
                            "WHERE tiebreaker_method = 'FVS Variable'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            oAdo.m_strSQL = "UPDATE scenario_fvs_variables_optimization " +
                            "SET optimization_variable = 'Stand Attribute' " +
                            "WHERE optimization_variable = 'FVS Variable'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);


            frmMain.g_sbpInfo.Text = "Version Update: Updating Harvest Methods and Tree Species tables ...Stand by";
            
            //Rename existing harvest_methods table
            // Load project data sources table
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();

            // Extract table properties from data sources table; Assume still under the old name
            string strTableSuffix = "_ver_control_" + DateTime.Now.ToString("MMddyyyy");
            int intHarvestMethodsTable = oDs.getValidTableNameRow(Datasource.TableTypes.HarvestMethods);
            string strDirectoryPath = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            string strFileStatus = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            string strTargetTable = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strTableStatus = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();

            if (strFileStatus == "F" && strTableStatus == "F")
            {
                oDao.RenameTable(strDirectoryPath + "\\" + strFileName, strTargetTable, strTargetTable + strTableSuffix, true, false);
            }

            // Copying the updated harvest_methods table into ref_master.accdb
            string strHarvestWorkTableName = "harvestmethod_worktable";
            string strSourceDbFile = frmMain.g_oEnv.strAppDir.Trim() + "\\" + Tables.Reference.DefaultHarvestMethodsTableDbFile;
            string strTargetDbFile = ReferenceProjectDirectory.Trim() + "\\" + Tables.Reference.DefaultHarvestMethodsTableDbFile;
            // Harvest Methods table
            oDao.CreateTableLink(strTargetDbFile, strHarvestWorkTableName, strSourceDbFile, strTargetTable);

            //copy contents of new harvest methods table into place
            oAdo.OpenConnection(oAdo.getMDBConnString(strTargetDbFile, "", ""));
            oAdo.m_strSQL = "SELECT * INTO " + strTargetTable + " FROM " + strHarvestWorkTableName;
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            //drop the harvest methods table link
            if (oAdo.TableExist(oAdo.m_OleDbConnection, strHarvestWorkTableName))
            {
                oAdo.m_strSQL = "DROP TABLE " + strHarvestWorkTableName;
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            }

            int intTreeSpeciesTable = oDs.getValidTableNameRow("Tree Species");
            strDirectoryPath = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strFileName = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            strFileStatus = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            strTargetTable = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            strTableStatus = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();

            if (strFileStatus == "F" && strTableStatus == "F")
            {
                oDao.RenameTable(strDirectoryPath + "\\" + strFileName, strTargetTable, strTargetTable + strTableSuffix, true, false);
            }

            // Copying the updated tree_species table into ref_master.accdb
            string strTreeSpeciesWorkTableName = "treespecies_worktable";
            // Tree species table
            oDao.CreateTableLink(strTargetDbFile, strTreeSpeciesWorkTableName, strSourceDbFile, strTargetTable);

            //copy contents of new harvest methods table into place
            oAdo.OpenConnection(oAdo.getMDBConnString(strTargetDbFile, "", ""));
            oAdo.m_strSQL = "SELECT * INTO " + strTargetTable + " FROM " + strTreeSpeciesWorkTableName;
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            //drop the tree species table link
            if (oAdo.TableExist(oAdo.m_OleDbConnection, strTreeSpeciesWorkTableName))
            {
                oAdo.m_strSQL = "DROP TABLE " + strTreeSpeciesWorkTableName;
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            }

            frmMain.g_sbpInfo.Text = "Version Update: Updating Reference Tables ...Stand by";

            string strSourceFile = frmMain.g_oEnv.strAppDir + "\\db\\" + Tables.Reference.DefaultBiosumReferenceDbFile;
            string strDestFile = frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceDbFile;
            if (System.IO.File.Exists(strDestFile) == true)
            {
                string strBackupFileName = System.IO.Path.GetFileNameWithoutExtension(strSourceFile) + strTableSuffix + ".accdb";
                if (System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                    frmMain.g_strBiosumDataDir + "\\" + strBackupFileName) == false)
                {
                    System.IO.File.Move(strDestFile, frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                    frmMain.g_strBiosumDataDir + "\\" + strBackupFileName);
                }
            }
            System.IO.File.Copy(strSourceFile, strDestFile, true);

            strSourceFile = frmMain.g_oEnv.strAppDir + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile;
            strDestFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                          "\\" + Tables.Reference.DefaultOpCostReferenceDbFile;
            if (System.IO.File.Exists(strDestFile) == false)
            {
                System.IO.File.Copy(strSourceFile, strDestFile);
            }

            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
            if (oAdo != null)
            {
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }
        }

        private void UpdateDatasources_5_8_5()
        {
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();

            frmMain.g_sbpInfo.Text = "Version Update: Updating tree species tables ...Stand by";
            // Load project data sources table
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();

            string strTableSuffix = "_ver_control_" + DateTime.Now.ToString("MMddyyyy");
            int intTreeSpeciesTable = oDs.getValidTableNameRow("Tree Species");
            string strDirectoryPath = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            string strFileStatus = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            string strTargetTable = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strTableStatus = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();

            if (strFileStatus == "F" && strTableStatus == "F")
            {
                oDao.RenameTable(strDirectoryPath + "\\" + strFileName, strTargetTable, strTargetTable + strTableSuffix, true, false);
            }

            // Copying the updated tree_species table into ref_master.accdb
            string strTreeSpeciesWorkTableName = "treespecies_worktable";
            string strSourceFile = frmMain.g_oEnv.strAppDir.Trim() + "\\db\\ref_master.mdb";
            // Tree species table
            oDao.CreateTableLink(strDirectoryPath + "\\" + strFileName, strTreeSpeciesWorkTableName,
                                 strSourceFile, strTargetTable);

            //copy contents of new tree_species table into place
            oAdo.OpenConnection(oAdo.getMDBConnString(strDirectoryPath + "\\" + strFileName, "", ""));
            oAdo.m_strSQL = "SELECT * INTO " + strTargetTable + " FROM " + strTreeSpeciesWorkTableName;
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            //drop the tree species table link
            if (oAdo.TableExist(oAdo.m_OleDbConnection, strTreeSpeciesWorkTableName))
            {
                oAdo.m_strSQL = "DROP TABLE " + strTreeSpeciesWorkTableName;
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            }

            //refresh biosum_ref.accdb from application directory
            strSourceFile = frmMain.g_oEnv.strAppDir.Trim() + "\\db\\biosum_ref.accdb";
            string strDestFile = frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceDbFile;
            if (System.IO.File.Exists(strDestFile) == true)
            {
                // Create backup copy of biosum_ref.accdb if one wasn't already made today
                string strBackupFileName = System.IO.Path.GetFileNameWithoutExtension(strSourceFile) + strTableSuffix + ".accdb";
                if (!System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                    frmMain.g_strBiosumDataDir + "\\" + strBackupFileName))
                {
                    System.IO.File.Move(strDestFile, frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                    frmMain.g_strBiosumDataDir + "\\" + strBackupFileName);
                }
                else
                {
                    // Delete the file if we already have a backup. Otherwise the copy will fail
                    System.IO.File.Delete(strDestFile);
                }
            }
            System.IO.File.Copy(strSourceFile, strDestFile);

            frmMain.g_sbpInfo.Text = "Version Update: Add DWM field to Cond table ...Stand by";
            // New column on cond table for dwm
            int intCondTable = oDs.getValidTableNameRow("Condition");
            strDirectoryPath = oDs.m_strDataSource[intCondTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strFileName = oDs.m_strDataSource[intCondTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            strFileStatus = oDs.m_strDataSource[intCondTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            strTargetTable = oDs.m_strDataSource[intCondTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            strTableStatus = oDs.m_strDataSource[intCondTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();

            if (strFileStatus == "F" && strTableStatus == "F")
            {
                oAdo.OpenConnection(oAdo.getMDBConnString(strDirectoryPath + "\\" + strFileName, "", ""));
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, strTargetTable, "dwm_fuelbed_typcd"))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, strTargetTable, "dwm_fuelbed_typcd", "TEXT", "3");
                }
            }

            frmMain.g_sbpInfo.Text = "Version Update: Creating empty DWM tables ...Stand by";
            strDestFile = ReferenceProjectDirectory.Trim() + "\\" + frmMain.g_oTables.m_oFIAPlot.DefaultDWMDbFile;
            if (!System.IO.File.Exists(strDestFile))
            {
                oDao.CreateMDB(strDestFile);
            }
            oAdo.OpenConnection(oAdo.getMDBConnString(strDestFile, "", ""));
            if (!oAdo.TableExist(oAdo.m_OleDbConnection, frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName))
            {
                frmMain.g_oTables.m_oFIAPlot.CreateDWMCoarseWoodyDebrisTable(oAdo, oAdo.m_OleDbConnection, 
                    frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName);
                frmMain.g_oTables.m_oFIAPlot.CreateDWMDuffLitterFuelTable(oAdo, oAdo.m_OleDbConnection,
                    frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName);
                frmMain.g_oTables.m_oFIAPlot.CreateDWMFineWoodyDebrisTable(oAdo, oAdo.m_OleDbConnection, 
                    frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName);
                frmMain.g_oTables.m_oFIAPlot.CreateDWMTransectSegmentTable(oAdo, oAdo.m_OleDbConnection,
                    frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName);
            }

            //rename fvs_tree_species table and re-map to %appData%
            frmMain.g_sbpInfo.Text = "Version Update: Move  table ...Stand by";

            int intFvsVariantTable = oDs.getValidTableNameRow("FIADB FVS Variant");
            strDirectoryPath = oDs.m_strDataSource[intFvsVariantTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strFileName = oDs.m_strDataSource[intFvsVariantTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            strFileStatus = oDs.m_strDataSource[intFvsVariantTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            strTargetTable = oDs.m_strDataSource[intFvsVariantTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            strTableStatus = oDs.m_strDataSource[intFvsVariantTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
            if (strFileStatus == "F" && strTableStatus == "F")
            {
                oDao.RenameTable(strDirectoryPath + "\\" + strFileName, strTargetTable, strTargetTable + strTableSuffix, true, false);
            }

            frmMain.g_sbpInfo.Text = "Version Update: Moving FVS Variant table ...Stand by";
            string strDataSourceMdb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\project.mdb";
            oAdo.OpenConnection(oAdo.getMDBConnString(strDataSourceMdb, "", ""));
            oAdo.m_strSQL = "UPDATE datasource " +
                            "SET PATH = '@@appdata@@\\fiabiosum', file = '" + Tables.Reference.DefaultBiosumReferenceDbFile + "' " +
                            "WHERE TABLE_TYPE = '" + Datasource.TableTypes.FVSVariant + "'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            oAdo.m_OleDbConnection.Close();

            frmMain.g_sbpInfo.Text = "Version Update: Adding fvsloccode to Plot table ...Stand by";
            int intPlotTable = oDs.getValidTableNameRow("Plot");
            strDirectoryPath = oDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strFileName = oDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            strFileStatus = oDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            strTargetTable = oDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            strTableStatus = oDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
            if (strFileStatus == "F" && strTableStatus == "F")
            {
                string strLocCodeFieldName = "fvsloccode";
                oAdo.OpenConnection(oAdo.getMDBConnString(strDirectoryPath + "\\" + strFileName, "", ""));
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, strTargetTable, strLocCodeFieldName))
                {
                    oAdo.AddColumn(oAdo.m_OleDbConnection, strTargetTable, strLocCodeFieldName, "INTEGER", "");
                }
                oAdo.m_OleDbConnection.Close();
            }

            frmMain.g_sbpInfo.Text = "Version Update: Updating OPCOST configuration database ...Stand by";
            strSourceFile = frmMain.g_oEnv.strAppDir + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile;
            strDestFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                            "\\" + Tables.Reference.DefaultOpCostReferenceDbFile;
            if (System.IO.File.Exists(strDestFile) == true)
            {
                System.IO.File.Delete(strDestFile);
            }
            System.IO.File.Copy(strSourceFile, strDestFile);

            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
            if (oAdo != null)
            {
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }
        }
        
        private void UpdateDatasources_5_8_6()
        {
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();

            string strTableSuffix = "_ver_control_" + DateTime.Now.ToString("MMddyyyy");
            frmMain.g_sbpInfo.Text = "Version Update: Creating new Treatment Optimizer databases ...Stand by";
            // Rename core folder to optimizer
            System.IO.Directory.Move(ReferenceProjectDirectory.Trim() + "\\core", ReferenceProjectDirectory.Trim() + "\\optimizer");
            string strSourceFile = frmMain.g_oEnv.strAppDir.Trim() + "\\db\\optimizer_definitions.accdb";
            string strDestFile = ReferenceProjectDirectory.Trim() + "\\" + Tables.OptimizerDefinitions.DefaultDbFile;
            if (!System.IO.File.Exists(strDestFile))
            {
                System.IO.File.Copy(strSourceFile, strDestFile);
            }
            strDestFile = ReferenceProjectDirectory.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile;
            if (!System.IO.File.Exists(strDestFile))
            {
                oDao.CreateMDB(strDestFile);
            }

            frmMain.g_sbpInfo.Text = "Version Update: Updating file structure for OPTIMIZER name change ...Stand by";
            string strRuleDefinitionsMdb = ReferenceProjectDirectory.Trim() + "\\optimizer\\db\\scenario_optimizer_rule_definitions.mdb";
            System.IO.File.Move(ReferenceProjectDirectory.Trim() + "\\optimizer\\db\\scenario_core_rule_definitions.mdb",
                strRuleDefinitionsMdb);
            string strRenameConn = m_oAdo.getMDBConnString(strRuleDefinitionsMdb, "", "");
            using (var oRenameConn = new OleDbConnection(strRenameConn))
            {
                oRenameConn.Open();
                oAdo.m_strSQL = "SELECT SCENARIO_ID FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName;
                oAdo.SqlQueryReader(oRenameConn, oAdo.m_strSQL);
                if (oAdo.m_OleDbDataReader.HasRows)
                {
                    while (oAdo.m_OleDbDataReader.Read())
                    {
                        string strScenario = "";
                        if (oAdo.m_OleDbDataReader["scenario_id"] != System.DBNull.Value)
                        {
                            strScenario = oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim();
                            string strUpdate = "UPDATE " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableName +
                                " SET PATH = '" + ReferenceProjectDirectory.Trim() + "\\optimizer\\" + strScenario +
                                "', FILE = 'scenario_optimizer_rule_definitions.mdb'" +
                                " WHERE SCENARIO_ID = '" + strScenario + "'";
                            oAdo.SqlNonQuery(oRenameConn, strUpdate);
                        }
                    }
                }
            }

            frmMain.g_sbpInfo.Text = "Version Update: Updating OPTIMIZER scenario configuration tables ...Stand by";

            strDestFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                            "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableDbFile;
            //open the scenario_optimizer_rule_definitions.mdb file
            oAdo.OpenConnection(oAdo.getMDBConnString(strDestFile, "", ""));
            //add new revenue_attribute field if it is missing
            if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName,
                "revenue_attribute"))
            {
                oAdo.AddColumn(oAdo.m_OleDbConnection, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName,
                    "revenue_attribute", "CHAR", "100");
            }
            //remove filter fields from scenario_fvs_variables_overall_effective
            if (oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName,
                "nr_dpa_filter_enabled_yn"))
            {
                string[] arrFieldsToDelete = new string[] { "nr_dpa_filter_enabled_yn", "nr_dpa_filter_operator", "nr_dpa_filter_value" };
                oDao.DeleteField(strDestFile, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName,
                    arrFieldsToDelete);
            }
            //replace scenario_rx_intensity with scenario_last_tiebreak_rank 
            if (oAdo.TableExist(oAdo.m_OleDbConnection, "scenario_rx_intensity"))
            {
                oDao.RenameTable(strDestFile, "scenario_rx_intensity", "scenario_rx_intensity" + strTableSuffix, true, false);
            }
            frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioLastTieBreakRankTable(oAdo, oAdo.m_OleDbConnection,
                Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLastTieBreakRankTableName);
            //populate scenario_last_tiebreak_rank with packages for each scenario            
            string strConn="";
            string strRxMDBFile = "";
            string strRxPackageTableName = "";
            string strRxConn = "";
            oAdo.getScenarioConnStringAndMDBFile(ref strSourceFile,
                              ref strConn, frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim());
            oAdo.OpenConnection(strConn);

            //retrieve paths for all scenarios in the project and put them in list
            List<string> lstScenario = new List<string>();
            oAdo.m_strSQL = "SELECT path, scenario_id from scenario";
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    string strPath = "";
                    if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value)
                        strPath = oAdo.m_OleDbDataReader["path"].ToString().Trim();
                    if (!String.IsNullOrEmpty(strPath))
                    {
                        if (System.IO.Directory.Exists(strPath))
                            lstScenario.Add(oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim());
                    }
                }
                oAdo.m_OleDbDataReader.Close();
            }

            foreach (string strScenarioId in lstScenario)
            {
                
                /*************************************************************************
                 **get the treatment prescription mdb file,table, and connection strings
                 *************************************************************************/
                oAdo.getScenarioDataSourceConnStringAndTable(ref strRxMDBFile,
                                                ref strRxPackageTableName, ref strRxConn,
                                                "Treatment Packages",
                                                strScenarioId,
                                                oAdo.m_OleDbConnection);

                oAdo.OpenConnection(strRxConn);
                if (oAdo.m_intError != 0)
                {
                    oAdo.m_OleDbConnection.Close();
                    oAdo.m_OleDbConnection = null;
                    return;
                }
                oAdo.m_strSQL = "select * from " + strRxPackageTableName;
                oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                /********************************************************************************
                 **insert records into the scenario_last_tiebreak_rank table from the master rxpackage table
                 ********************************************************************************/
                List<string> lstRxPackages = new List<string>();
                if (oAdo.m_intError == 0)
                {
                    if (oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (oAdo.m_OleDbDataReader.Read())
                        {
                            string strRxPackage = "";
                            if (oAdo.m_OleDbDataReader["rxpackage"] != System.DBNull.Value)
                                strRxPackage = oAdo.m_OleDbDataReader["rxpackage"].ToString().Trim();
                            if (!String.IsNullOrEmpty(strRxPackage))
                            {
                                lstRxPackages.Add(strRxPackage);
                            }
                        }
                        oAdo.m_OleDbDataReader.Close();

                        oAdo.OpenConnection(strConn);
                        foreach (string strRxPackage in lstRxPackages)
                        {
                            oAdo.m_strSQL = "INSERT INTO scenario_last_tiebreak_rank (scenario_id," +
                            "rxpackage) VALUES " +
                            "('" + strScenarioId + "'," +
                            "'" + strRxPackage + "')";
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                        }
                    }
                }
            }

            frmMain.g_sbpInfo.Text = "Version Update: Renaming frcs_harvest_costs_yn columns in audit tables ...Stand by";
            string[] arrDatabases = System.IO.Directory.GetFiles(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db");
            string strOldColumnName = "frcs_harvest_costs_yn";
            string strNewColumnName = "harvest_costs_yn";
            foreach (string strDatabase in arrDatabases)
            {
                string strDatabaseName = System.IO.Path.GetFileName(strDatabase);
                if (strDatabaseName.StartsWith("audit"))
                {
                    strRenameConn = m_oAdo.getMDBConnString(strDatabase, "", "");
                    using (var oRenameConn = new OleDbConnection(strRenameConn))
                    {
                        oRenameConn.Open();
                        if (oAdo.ColumnExist(oRenameConn, Tables.Audit.DefaultCondAuditTableName, strOldColumnName)) ;
                        {
                            oDao.RenameField(strDatabase, Tables.Audit.DefaultCondAuditTableName, strOldColumnName, strNewColumnName);
                        }
                        if (oAdo.ColumnExist(oRenameConn, Tables.Audit.DefaultCondRxAuditTableName, strOldColumnName)) ;
                        {
                            oDao.RenameField(strDatabase, Tables.Audit.DefaultCondRxAuditTableName, strOldColumnName, strNewColumnName);
                        }
                    }
                }
            }

            //frmMain.g_sbpInfo.Text = "Version Update: Creating empty GRM tables ...Stand by";
            //strDestFile = ReferenceProjectDirectory.Trim() + "\\" + frmMain.g_oTables.m_oFIAPlot.DefaultDWMDbFile;
            //oAdo.OpenConnection(oAdo.getMDBConnString(strDestFile, "", ""));
            //if (!oAdo.TableExist(oAdo.m_OleDbConnection, frmMain.g_oTables.m_oFIAPlot.DefaultMasterAuxGRMStandName))
            //{
            //    frmMain.g_oTables.m_oFIAPlot.CreateMasterAuxGRMStandTable(oAdo, oAdo.m_OleDbConnection,
            //        frmMain.g_oTables.m_oFIAPlot.DefaultMasterAuxGRMStandName);
            //    frmMain.g_oTables.m_oFIAPlot.CreateMasterAuxGRMTreeTable(oAdo, oAdo.m_OleDbConnection,
            //        frmMain.g_oTables.m_oFIAPlot.DefaultMasterAuxGRMTreeName);
            //}

            // Replace opcost_ref.accdb; In the future we want to back it up, but not used much yet
            frmMain.g_sbpInfo.Text = "Version Update: Updating OPCOST configuration database ...Stand by";
            strSourceFile = frmMain.g_oEnv.strAppDir + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile;
            strDestFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                            "\\" + Tables.Reference.DefaultOpCostReferenceDbFile;
            if (System.IO.File.Exists(strDestFile) == true)
            {
                System.IO.File.Delete(strDestFile);
            }
            System.IO.File.Copy(strSourceFile, strDestFile);

            //Rename existing harvest_methods table
            // Load project data sources table
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();

            int intHarvestMethodsTable = oDs.getValidTableNameRow(Datasource.TableTypes.HarvestMethods);
            string strDirectoryPath = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            string strFileStatus = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            string strTargetTable = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strTableStatus = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();

            if (strFileStatus == "F" && strTableStatus == "F")
            {
                oDao.RenameTable(strDirectoryPath + "\\" + strFileName, strTargetTable, strTargetTable + strTableSuffix, true, false);
            }

            // Copying the updated harvest_methods table into ref_master.accdb
            string strHarvestWorkTableName = "harvestmethod_worktable";
            string strSourceDbFile = frmMain.g_oEnv.strAppDir.Trim() + "\\" + Tables.Reference.DefaultHarvestMethodsTableDbFile;
            string strTargetDbFile = ReferenceProjectDirectory.Trim() + "\\" + Tables.Reference.DefaultHarvestMethodsTableDbFile;
            // Harvest Methods table
            oDao.CreateTableLink(strTargetDbFile, strHarvestWorkTableName, strSourceDbFile, strTargetTable);

            //copy contents of new harvest methods table into place
            oAdo.OpenConnection(oAdo.getMDBConnString(strTargetDbFile, "", ""));
            oAdo.m_strSQL = "SELECT * INTO " + strTargetTable + " FROM " + strHarvestWorkTableName;
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            //drop the harvest methods table link
            if (oAdo.TableExist(oAdo.m_OleDbConnection, strHarvestWorkTableName))
            {
                oAdo.m_strSQL = "DROP TABLE " + strHarvestWorkTableName;
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            }


            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
            if (oAdo != null)
            {
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }
        }

        private void UpdateDatasources_5_8_7()
        {
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();

            frmMain.g_sbpInfo.Text = "Version Update: Update variable source for Calculated Variables ...Stand by";

            string strRenameMdb = ReferenceProjectDirectory.Trim() + "\\optimizer\\db\\optimizer_definitions.accdb";
            string strRenameConn = m_oAdo.getMDBConnString(strRenameMdb, "", "");
            using (var oRenameConn = new OleDbConnection(strRenameConn))
            {
                oRenameConn.Open();
                oAdo.m_strSQL = "SELECT ID, VARIABLE_SOURCE FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                    " WHERE UCASE(variable_source) like 'PRODUCT_YIELDS%'";
                oAdo.SqlQueryReader(oRenameConn, oAdo.m_strSQL);
                if (oAdo.m_OleDbDataReader.HasRows)
                {
                    string[] arrOldSources = {"PRODUCT_YIELDS_NET_REV_COSTS_SUMMARY_BY_RXPACKAGE.chip_yield_cf",
                                              "PRODUCT_YIELDS_NET_REV_COSTS_SUMMARY_BY_RXPACKAGE.merch_yield_cf",
                                              "PRODUCT_YIELDS_NET_REV_COSTS_SUMMARY_BY_RXPACKAGE.MAX_NR_DPA",
                                              "PRODUCT_YIELDS_NET_REV_COSTS_SUMMARY_BY_RXPACKAGE.HARVEST_ONSITE_CPA"};
                    string[] arrUpdatedSources = {"ECON_BY_RX_SUM.chip_vol_cf",
                                                  "ECON_BY_RX_SUM.merch_vol_cf",
                                                  "ECON_BY_RX_SUM.MAX_NR_DPA",
                                                  "ECON_BY_RX_SUM.HARVEST_ONSITE_COST_DPA"};
                    while (oAdo.m_OleDbDataReader.Read())
                    {
                        string strVariableSource = Convert.ToString(oAdo.m_OleDbDataReader["variable_source"]).Trim();
                        int i = 0;
                        foreach (string strOldSource in arrOldSources)
                        {
                            if (strVariableSource.ToUpper().Equals(arrOldSources[i].ToUpper()))
                            {
                                int intId = Convert.ToInt16(oAdo.m_OleDbDataReader["id"]);
                                string strUpdate = "UPDATE " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                                    " SET VARIABLE_SOURCE = '" + arrUpdatedSources[i] + "'" +
                                    " WHERE ID = " + intId;
                                oAdo.SqlNonQuery(oRenameConn, strUpdate);
                                break;
                            }
                            i++;
                        }
                    }
                 }
              }

            frmMain.g_sbpInfo.Text = "Version Update: Updating travel times database and table ...Stand by";

            //Rename existing harvest_methods table
            // Load project data sources table
            FIA_Biosum_Manager.Datasource oProjectDs = new Datasource();
            oProjectDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oProjectDs.m_strDataSourceTableName = "datasource";
            oProjectDs.m_strScenarioId = "";
            oProjectDs.LoadTableColumnNamesAndDataTypes = false;
            oProjectDs.LoadTableRecordCount = false;
            oProjectDs.populate_datasource_array();

            int intTravelTimesTable = oProjectDs.getValidTableNameRow(Datasource.TableTypes.TravelTimes);
            string strDirectoryPath = "";
            string strFileName = "";

            if (intTravelTimesTable > -1)
            {
                strDirectoryPath = oProjectDs.m_strDataSource[intTravelTimesTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
                strFileName = oProjectDs.m_strDataSource[intTravelTimesTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
                //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
                string strTableName = oProjectDs.m_strDataSource[intTravelTimesTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
                string strTableStatus = oProjectDs.m_strDataSource[intTravelTimesTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();

                if (strTableStatus == "F")
                {
                    if (oDao.IndexExists(strDirectoryPath + "\\" + strFileName, strTableName, "travel_time_idx3"))
                    {
                        string strCommand = "DROP INDEX travel_time_idx3 ON " + strTableName;
                        oDao.OpenDb(strDirectoryPath + "\\" + strFileName);
                        oDao.m_DaoDatabase.Execute(strCommand, null);
                        oDao.RenameField(strDirectoryPath + "\\" + strFileName, strTableName, "PLOT_ID", "PLOT");
                        // Note: the RenameField method closes the database
                    }
                    oDao.OpenDb(strDirectoryPath + "\\" + strFileName);
                    if (!oDao.ColumnExist(oDao.m_DaoDatabase, strTableName, "STATECD"))
                    {
                        string strTravelConn = m_oAdo.getMDBConnString(strDirectoryPath + "\\" + strFileName, "", "");
                        using (var oTravelConn = new OleDbConnection(strTravelConn))
                        {
                            oTravelConn.Open();
                            oAdo.AddColumn(oTravelConn, strTableName, "STATECD", "INTEGER", "");
                        }

                    }
                    if (oDao.ColumnExist(oDao.m_DaoDatabase, strTableName, "TRAVEL_TIME"))
                    {
                        oDao.RenameField(strDirectoryPath + "\\" + strFileName, strTableName, "TRAVEL_TIME", "ONE_WAY_HOURS");
                        // Note: the RenameField method closes the database
                    }

                    if (oDao.m_DaoDatabase != null)
                        oDao.m_DaoDatabase.Close();
                }
            }

            // Check to see if gis_travel_times.accdb before trying to create it
            if (! System.IO.File.Exists(strDirectoryPath + "\\" + Tables.TravelTime.DefaultTravelTimeAccdbFile))
            {
                oDao.CreateMDB(strDirectoryPath + "\\" + Tables.TravelTime.DefaultTravelTimeAccdbFile);
                // create table links to copy tables
                string[] arrTableNames = new string[0];
                oDao.getTableNames(strDirectoryPath + "\\" + strFileName, ref arrTableNames, false);
                string strCopyConn = m_oAdo.getMDBConnString(strDirectoryPath + "\\" + Tables.TravelTime.DefaultTravelTimeAccdbFile, "", "");
                using (var oCopyConn = new OleDbConnection(strCopyConn))
                {
                    oCopyConn.Open();
                    foreach (string strTable in arrTableNames)
                    {
                        bool bNextTable = true;
                        if (!String.IsNullOrEmpty(strTable) && bNextTable == true)
                        {
                            oDao.CreateTableLink(strDirectoryPath + "\\" + Tables.TravelTime.DefaultTravelTimeAccdbFile, strTable + "_1",
                                strDirectoryPath + "\\" + strFileName, strTable);
                            int i = 0;
                            do
                            {
                                // break out of loop if it runs too long
                                if (i > 20)
                                {
                                    System.Windows.Forms.MessageBox.Show("An error occurred while trying to update gis_travel_times.accdb! " +
                                    "Validate the contents of this database before trying to run Treatment Optimizer.", "FIA Biosum");
                                    bNextTable = false;  //error flag: should we continue?
                                    break;
                                }
                                System.Threading.Thread.Sleep(1000);
                                i++;
                            }
                            while (!oAdo.TableExist(oCopyConn, strTable + "_1"));

                            if (bNextTable == true)
                            {
                                string strSql = "SELECT * INTO " + strTable + " FROM " + strTable + "_1";
                                oAdo.SqlNonQuery(oCopyConn, strSql);
                                strSql = "DROP TABLE " + strTable + "_1";
                                oAdo.SqlNonQuery(oCopyConn, strSql);
                            }
                        }
                    }
                }
            }

            int intOldAuditTable = oProjectDs.getValidTableNameRow("Plot And Condition Record Audit");
            if (intOldAuditTable >= 0)
            {
                string strOldAuditPath = oProjectDs.m_strDataSource[intOldAuditTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
                //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
                string strOldAuditStatus = oProjectDs.m_strDataSource[intOldAuditTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();

                frmMain.g_sbpInfo.Text = "Version Update: Delete old audit tables ...Stand by";
                if (strOldAuditStatus == "F")
                {
                    string[] arrDbPaths = System.IO.Directory.GetFiles(strOldAuditPath);
                    if (arrDbPaths.Length > 0)
                    {
                        foreach (string strDbPath in arrDbPaths)
                        {
                            string strDbFile = System.IO.Path.GetFileName(strDbPath);
                            if (strDbFile.IndexOf("audit") > -1)
                            {
                                System.IO.File.Delete(strDbPath);
                            }
                        }
                    }
                }
            }

            frmMain.g_sbpInfo.Text = "Version Update: Updating data source tables ...Stand by";
            
            // Main datasource table
            string strDataSourceMdb = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oAdo.OpenConnection(oAdo.getMDBConnString(strDataSourceMdb, "", ""));
            oAdo.m_strSQL = "UPDATE datasource " +
                            "SET file = '" + Tables.TravelTime.DefaultTravelTimeAccdbFile+ "' " +
                            "WHERE TABLE_TYPE IN ('" + Datasource.TableTypes.TravelTimes + "', '" +
                            Datasource.TableTypes.ProcessingSites + "')";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            oAdo.m_strSQL = "DELETE * FROM datasource " +
                            "WHERE UCASE(FILE) = 'AUDIT.MDB'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            oAdo.m_OleDbConnection.Close();

            // Processor datasource table
            strDataSourceMdb = ReferenceProjectDirectory.Trim() + "\\processor\\db\\scenario_processor_rule_definitions.mdb";
            oAdo.OpenConnection(oAdo.getMDBConnString(strDataSourceMdb, "", ""));
            oAdo.m_strSQL = "UPDATE scenario_datasource " +
                            "SET file = '" + Tables.TravelTime.DefaultTravelTimeAccdbFile + "' " +
                            "WHERE TABLE_TYPE IN ('" + Datasource.TableTypes.TravelTimes + "', '" +
                            Datasource.TableTypes.ProcessingSites + "')";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            oAdo.m_OleDbConnection.Close();

            // Optimizer datasource table
            strDataSourceMdb = ReferenceProjectDirectory.Trim() + "\\optimizer\\db\\scenario_optimizer_rule_definitions.mdb";
            oAdo.OpenConnection(oAdo.getMDBConnString(strDataSourceMdb, "", ""));
            oAdo.m_strSQL = "UPDATE scenario_datasource " +
                            "SET file = '" + Tables.TravelTime.DefaultTravelTimeAccdbFile + "' " +
                            "WHERE TABLE_TYPE IN ('" + Datasource.TableTypes.TravelTimes + "', '" +
                            Datasource.TableTypes.ProcessingSites + "')";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            oAdo.m_strSQL = "DELETE * FROM scenario_datasource " +
                "WHERE UCASE(FILE) = 'AUDIT.MDB'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            oAdo.m_strSQL = "ALTER TABLE SCENARIO_PSITES ALTER COLUMN NAME TEXT (100)"; 
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            oAdo.m_OleDbConnection.Close();

            frmMain.g_sbpInfo.Text = "Version Update: Delete obsolete fields from plot and cond tables ...Stand by";

            int intPlotTable = oProjectDs.getValidTableNameRow(Datasource.TableTypes.Plot);
            string strPlotDirectory = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strPlotMdbFile = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            string strPlotTable = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strPlotStatus = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
            if (strPlotStatus == "F")
            {
                oDao.OpenDb(strPlotDirectory + "\\" + strPlotMdbFile);
                if (oDao.IndexExists(strPlotDirectory + "\\" + strPlotMdbFile, strPlotTable, "plot_idx3"))
                {
                    string strCommand = "DROP INDEX plot_idx3 ON " + strPlotTable;
                    oDao.m_DaoDatabase.Execute(strCommand, null);
                }
                if (oDao.IndexExists(strPlotDirectory + "\\" + strPlotMdbFile, strPlotTable, "plot_idx4"))
                {
                    string strCommand = "DROP INDEX plot_idx4 ON " + strPlotTable;
                    oDao.m_DaoDatabase.Execute(strCommand, null);
                }
                string[] arrFieldsToDelete = {"MERCH_HAUL_COST_ID","MERCH_HAUL_COST_PSITE", "MERCH_HAUL_CPA_PT",
                                              "CHIP_HAUL_COST_ID","CHIP_HAUL_COST_PSITE", "CHIP_HAUL_CPA_PT",
                                              "gis_status_id","idb_plot_id", "gis_protected_area_yn",
                                              "gis_roadless_yn","PLOT_ACCESSIBLE_YN", "ALL_COND_NOT_ACCESSIBLE_YN"};
                oDao.DeleteField(strPlotDirectory + "\\" + strPlotMdbFile, strPlotTable, arrFieldsToDelete);
                oDao.OpenDb(strPlotDirectory + "\\" + strPlotMdbFile);
                if (oDao.ColumnExist(oDao.m_DaoDatabase, strPlotTable, "gis_yard_dist"))
                {
                    oDao.RenameField(strPlotDirectory + "\\" + strPlotMdbFile, strPlotTable, "gis_yard_dist", "gis_yard_dist_ft");
                }
            }

            intPlotTable = oProjectDs.getValidTableNameRow(Datasource.TableTypes.Condition);
            strPlotDirectory = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strPlotMdbFile = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            strPlotTable = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            strPlotStatus = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
            if (strPlotStatus == "F")
            {
                oDao.OpenDb(strPlotDirectory + "\\" + strPlotMdbFile);
                if (oDao.IndexExists(strPlotDirectory + "\\" + strPlotMdbFile, strPlotTable, "cond_idx4"))
                {
                    string strCommand = "DROP INDEX cond_idx4 ON " + strPlotTable;
                    oDao.m_DaoDatabase.Execute(strCommand, null);
                }
                if (oDao.IndexExists(strPlotDirectory + "\\" + strPlotMdbFile, strPlotTable, "cond_idx5"))
                {
                    string strCommand = "DROP INDEX cond_idx5 ON " + strPlotTable;
                    oDao.m_DaoDatabase.Execute(strCommand, null);
                }
                if (oDao.IndexExists(strPlotDirectory + "\\" + strPlotMdbFile, strPlotTable, "cond_idx3"))
                {
                    string strCommand = "DROP INDEX cond_idx3 ON " + strPlotTable;
                    oDao.m_DaoDatabase.Execute(strCommand, null);
                }
                string[] arrFieldsToDelete = {"COND_TOO_FAR_STEEP_YN","COND_ACCESSIBLE_YN", "harvest_technique",
                                              "idb_cond_id","idb_plot_id", "sdi",
                                              "ccf","topht", "fvs_filename"};
                oDao.DeleteField(strPlotDirectory + "\\" + strPlotMdbFile, strPlotTable, arrFieldsToDelete);
                strDataSourceMdb = strPlotDirectory + "\\" + strPlotMdbFile;
                oAdo.OpenConnection(oAdo.getMDBConnString(strDataSourceMdb, "", ""));
                if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, strPlotTable, "MODEL_YN"))
                {
                   oAdo.AddColumn(oAdo.m_OleDbConnection, strPlotTable, "MODEL_YN", "CHAR", "1", "Y");
                   oAdo.m_strSQL = "UPDATE " + strPlotTable + " SET MODEL_YN = 'Y'";
                   oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                }
                oAdo.m_OleDbConnection.Close();
            }

            frmMain.g_sbpInfo.Text = "Version Update: Update scenario_costs field names ...Stand by";
            strDataSourceMdb = ReferenceProjectDirectory.Trim() + "\\optimizer\\db\\scenario_optimizer_rule_definitions.mdb";
            oAdo.OpenConnection(oAdo.getMDBConnString(strDataSourceMdb, "", ""));
            if (oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName, "rail_chip_transfer_pgt_per_hour"))
            {
                oDao.RenameField(strDataSourceMdb, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName,
                    "rail_chip_transfer_pgt_per_hour", "rail_chip_transfer_pgt");
            }
            if (oAdo.ColumnExist(oAdo.m_OleDbConnection, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName, "rail_merch_transfer_pgt_per_hour"))
            {
                oDao.RenameField(strDataSourceMdb, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName,
                    "rail_merch_transfer_pgt_per_hour", "rail_merch_transfer_pgt");
            }
            oAdo.m_OleDbConnection.Close();


            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
            if (oAdo != null)
            {
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }
        }

        private void UpdateDatasources_5_8_8()
        {
            //v5.8.7 was never in wide release. This allows projects to be upgraded from v5.8.6 without having to install v5.8.7
            if (this.m_strProjectVersion.Equals("5.8.6"))
            {
                UpdateDatasources_5_8_7();
            }
            
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();

            // Resize column name
            string strMdb = ReferenceProjectDirectory.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableDbFile;
            oAdo.OpenConnection(oAdo.getMDBConnString(strMdb, "", ""));
            oAdo.m_strSQL = "ALTER TABLE " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableName + " ALTER COLUMN NAME TEXT (100)";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            // Add entries to Optimizer scenario_datasource table
            oAdo.m_strSQL = "SELECT SCENARIO_ID FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName;
            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                IList<string> lstScenarios = new List<string>();
                while (oAdo.m_OleDbDataReader.Read())
                {
                    if (oAdo.m_OleDbDataReader["scenario_id"] != System.DBNull.Value)
                    {
                        lstScenarios.Add(oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim());
                    }
                }
                oAdo.m_OleDbDataReader.Dispose();

                // Get table locations from project data sources
                // Load project data sources table
                FIA_Biosum_Manager.Datasource oProjectDs = new Datasource();
                oProjectDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
                oProjectDs.m_strDataSourceTableName = "datasource";
                oProjectDs.m_strScenarioId = "";
                oProjectDs.LoadTableColumnNamesAndDataTypes = false;
                oProjectDs.LoadTableRecordCount = false;
                oProjectDs.populate_datasource_array();

                int intHarvestMethodsTable = oProjectDs.getValidTableNameRow(Datasource.TableTypes.HarvestMethods);
                int intRxHarvestTable = oProjectDs.getValidTableNameRow("Treatment Prescriptions Harvest Cost Columns");
                string strHarvestMethodDirectoryPath = "";
                string strHarvestMethodFileName = "";
                string strHarvestMethodTableName = "";
                if (intHarvestMethodsTable > -1)
                {
                    strHarvestMethodDirectoryPath = oProjectDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
                    strHarvestMethodFileName = oProjectDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
                    //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
                    strHarvestMethodTableName = oProjectDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
                    
                    foreach (string strScenario in lstScenarios)
                    {
                        oAdo.m_strSQL = "select COUNT(*) from scenario_datasource where scenario_id = '" + strScenario + "'" +
                                        " and table_type = '" + Datasource.TableTypes.HarvestMethods + "'";
                        int intCount = oAdo.getRecordCount(oAdo.m_OleDbConnection, oAdo.m_strSQL, "scenario_datasource");
                        if (intCount < 1)
                        {
                            oAdo.m_strSQL = "insert into scenario_datasource" +
                                            " (scenario_id, table_type, path, file, table_name)" +
                                            "values ('" + strScenario + "', '" + Datasource.TableTypes.HarvestMethods + "','" +
                                            strHarvestMethodDirectoryPath + "', '" + strHarvestMethodFileName + "', '" + strHarvestMethodTableName + "')";
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                        }
                    }
                }
                if (intRxHarvestTable > -1)
                {
                    string strDirectoryPath = oProjectDs.m_strDataSource[intRxHarvestTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
                    string strFileName = oProjectDs.m_strDataSource[intRxHarvestTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
                    //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
                    string strTableName = oProjectDs.m_strDataSource[intRxHarvestTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
                    foreach (string strScenario in lstScenarios)
                    {
                        oAdo.m_strSQL = "select COUNT(*) from scenario_datasource where scenario_id = '" + strScenario + "'" +
                                        " and table_type = 'Treatment Prescriptions Harvest Cost Columns'";
                        int intCount = oAdo.getRecordCount(oAdo.m_OleDbConnection, oAdo.m_strSQL, "scenario_datasource");
                        if (intCount < 1)
                        {
                            oAdo.m_strSQL = "insert into scenario_datasource" +
                                            " (scenario_id, table_type, path, file, table_name)" +
                                            "values ('" + strScenario + "', 'Treatment Prescriptions Harvest Cost Columns','" +
                                            strDirectoryPath + "', '" + strFileName + "', '" + strTableName + "')";
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                        }
                    }
                }
                // Add new and populate column to harvest methods table
                if (intHarvestMethodsTable > -1)
                {
                    frmMain.g_sbpInfo.Text = "Version Update: Adding and populating new column in harvest_methods table ...Stand by";
                    string strHarvestMdb = strHarvestMethodDirectoryPath + "\\" + strHarvestMethodFileName;
                    string strWorktable = "harvest_methods_work";
                    oAdo.OpenConnection(oAdo.getMDBConnString(strHarvestMdb, "", ""));
                    if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, strHarvestMethodTableName, "top_limb_slope_status"))
                    {
                        oAdo.CloseConnection(oAdo.m_OleDbConnection);   // close/reopen connection to avoid missing table errors after creating link
                        string strSourceDbFile = frmMain.g_oEnv.strAppDir.Trim() + "\\db\\ref_master.mdb";
                        oDao.CreateTableLink(strHarvestMethodDirectoryPath + "\\" + strHarvestMethodFileName, strWorktable,
                            strSourceDbFile, "harvest_methods");
                        oAdo.OpenConnection(oAdo.getMDBConnString(strHarvestMdb, "", ""));
                        int i = 0;
                        do
                        {
                            // break out of loop if it runs too long
                            if (i > 20)
                            {
                                System.Windows.Forms.MessageBox.Show("An error occurred while trying to update the harvest_methods table! " +
                                "Validate the contents of this table before trying to run Treatment Optimizer.", "FIA Biosum");
                                break;
                            }
                            System.Threading.Thread.Sleep(1000);
                            i++;
                        }
                        while (!oAdo.TableExist(oAdo.m_OleDbConnection, strWorktable));
                        
                        oAdo.AddColumn(oAdo.m_OleDbConnection, strHarvestMethodTableName, "top_limb_slope_status", "CHAR", "100");
                        if (oAdo.m_intError == 0)
                        {
                            oAdo.m_strSQL = "UPDATE " + strHarvestMethodTableName +
                                            " INNER JOIN " + strWorktable + " ON TRIM(" + strWorktable + ".Method)" +
                                            " = TRIM(" + strHarvestMethodTableName + ".Method) AND " +
                                            strWorktable + ".STEEP_YN = " + strHarvestMethodTableName + ".STEEP_YN" +
                                            " SET " + strHarvestMethodTableName + ".top_limb_slope_status = " + strWorktable + ".top_limb_slope_status";
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            oAdo.m_strSQL = "DROP TABLE " + strWorktable;
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                        }
                    }
                }

                // Update variable_source values for calculated economic variables
                string strVariablesAccdb = ReferenceProjectDirectory.Trim() + "\\" + Tables.OptimizerDefinitions.DefaultDbFile;
                oAdo.OpenConnection(oAdo.getMDBConnString(strVariablesAccdb, "", ""));
                oAdo.m_strSQL = "UPDATE " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                                " SET variable_source = replace(variable_source, 'ECON_BY_RX_SUM','ECON_BY_RX_UTILIZED_SUM')";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                oAdo.m_strSQL = "UPDATE " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                " SET variable_source = replace(variable_source, 'chip_vol','chip_vol_utilized')";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
            }
            oAdo.m_OleDbConnection.Close();


            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
            if (oAdo != null)
            {
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }
        }

        // Major upgrade to support 64-bit and remove Oracle dependency for USFS users
        // No schema changes between 5.8.8 and 5.8.9 but we want to provide backwards compatibilty to v5.8.6
        // without having to install earlier 32-bit application versions
        private void UpdateDatasources_5_8_9()
        {
            if (this.m_strProjectVersion.Equals("5.8.6"))
            {
                UpdateDatasources_5_8_7();
                UpdateDatasources_5_8_8();
            }
            else if(this.m_strProjectVersion.Equals("5.8.7"))
            {
                UpdateDatasources_5_8_8();
            }
        }

        public void UpdateDatasources_5_8_10()
        {
            //Update Master Tree/Cond/Plot tables
            ado_data_access oAdo = new ado_data_access();
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();

            int intTreeTableType = oDs.getDataSourceTableNameRow("TREE");
            if (oDs.DataSourceTableExist(intTreeTableType))
            {
                string strTreeTable = oDs.m_strDataSource[intTreeTableType, Datasource.TABLE].Trim();
                string strTreeTableDb = oDs.m_strDataSource[intTreeTableType, Datasource.PATH].Trim() + "\\" +
                                        oDs.m_strDataSource[intTreeTableType, Datasource.MDBFILE].Trim();
                using (OleDbConnection conn = new OleDbConnection(oAdo.getMDBConnString(strTreeTableDb, "", "")))
                {
                    conn.Open();
                    new List<string[]>
                    {
                        new string[] {"bfsnd", "DECIMAL(3,0)"},
                        new string[] {"boleht", "DECIMAL(3,0)"},
                        new string[] {"centroid_dia", "DECIMAL(4,1)"},
                        new string[] {"centroid_dia_ht_actual", "DECIMAL(4,1)"},
                        new string[] {"cfsnd", "DECIMAL(3,0)"},
                        new string[] {"cull_fld", "DECIMAL(2,0)"},
                        new string[] {"cullcf", "INTEGER"},
                        new string[] {"culldead", "DECIMAL(3,0)"},
                        new string[] {"cullform", "DECIMAL(3,0)"},
                        new string[] {"cullmstop", "DECIMAL(3,0)"},
                        new string[] {"drybio_bole", "DECIMAL(13,6)"},
                        new string[] {"drybio_sapling", "DECIMAL(13,6)"},
                        new string[] {"drybio_top", "DECIMAL(13,6)"},
                        new string[] {"drybio_wdld_spp", "DECIMAL(13,6)"},
                        new string[] {"htdmp", "DECIMAL(3,1)"},
                        new string[] {"sawht", "DECIMAL(2,0)"},
                        new string[] {"sitree", "INTEGER"},
                        new string[] {"standing_dead_cd", "INTEGER"},
                        new string[] {"upper_dia", "DECIMAL(4,1)"},
                        new string[] {"upper_dia_ht", "DECIMAL(4,1)"},
                        new string[] {"volcfsnd", "DECIMAL(13,6)"},
                    }.ForEach(column =>
                    {
                        if (!oAdo.ColumnExist(conn, strTreeTable, column[0]))
                            oAdo.AddColumn(conn, strTreeTable, column[0], column[1], "");
                    });
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The Master Tree table was not found.",
                    "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }

            int intCondTableType = oDs.getDataSourceTableNameRow("CONDITION");
            if (oDs.DataSourceTableExist(intCondTableType))
            {
                string strCondTable = oDs.m_strDataSource[intCondTableType, Datasource.TABLE].Trim();
                string strCondTableDb = oDs.m_strDataSource[intCondTableType, Datasource.PATH].Trim() + "\\" +
                                        oDs.m_strDataSource[intCondTableType, Datasource.MDBFILE].Trim();
                using (OleDbConnection conn = new OleDbConnection(oAdo.getMDBConnString(strCondTableDb, "", "")))
                {
                    conn.Open();
                    if (!oAdo.ColumnExist(conn, strCondTable, "balive"))
                        oAdo.AddColumn(conn, strCondTable, "balive", "DOUBLE", "");
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The Master Cond table was not found.",
                    "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }

            int intPlotTableType = oDs.getDataSourceTableNameRow("PLOT");
            if (oDs.DataSourceTableExist(intPlotTableType))
            {
                string strPlotTable = oDs.m_strDataSource[intPlotTableType, Datasource.TABLE].Trim();
                string strPlotTableDb = oDs.m_strDataSource[intPlotTableType, Datasource.PATH].Trim() + "\\" +
                                        oDs.m_strDataSource[intPlotTableType, Datasource.MDBFILE].Trim();
                using (OleDbConnection conn = new OleDbConnection(oAdo.getMDBConnString(strPlotTableDb, "", "")))
                {
                    conn.Open();
                    if (!oAdo.ColumnExist(conn, strPlotTable, "precipitation"))
                        oAdo.AddColumn(conn, strPlotTable, "precipitation", "DOUBLE", "");
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The Master Plot table was not found.",
                    "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }

            //Update FVS_Tree tables 
            var strVariants = oDs.getVariants();
            var strFvsTreeTable = Tables.FVS.DefaultFVSTreeTableName;
            var strDataBaseNames = new List<string>(); 

            //Get full path of each cutlist \biosumcalc\VARIANT_PACKAGE_TREE_CUTLIST.MDB
            if (strVariants != null && strVariants.Length > 0)
            {
                string strFvsDataDir = ReferenceProjectDirectory.Trim() + "\\fvs\\data\\";
                foreach (string variant in strVariants)
                {
                    string strBioSumCalcPath = strFvsDataDir + variant + "\\BiosumCalc\\";
                    if (System.IO.Directory.Exists(strBioSumCalcPath))
                    {
                        strDataBaseNames.AddRange(System.IO.Directory.GetFiles(strBioSumCalcPath, "*.*", System.IO.SearchOption.AllDirectories)
                            .Where(s => (s.ToLower().EndsWith(".mdb") || s.ToLower().EndsWith(".accdb")) && s.ToLower().Contains("cutlist")) .ToArray());
                    }
                }
            }

            //Connect to each database looking for FVS_TREE table and add any missing columns
            foreach (string db in strDataBaseNames)
            {
                using (var conn = new OleDbConnection(oAdo.getMDBConnString(db, "", "")))
                {
                    conn.Open();
                    if (oAdo.TableExist(conn, strFvsTreeTable))
                    {
                        new List<string[]>
                        {
                            new string[] {"drybio_bole", "double"},
                            new string[] {"drybio_sapling", "double"},
                            new string[] {"drybio_top", "double"},
                            new string[] {"drybio_wdld_spp", "double"},
                            new string[] {"volcfsnd", "double" },
                            new string[] {"volcfgrs", "DOUBLE"},
                            new string[] {"volcfnet", "DOUBLE"},
                            new string[] {"volcsgrs", "DOUBLE"},
                        }.ForEach(column =>
                        {
                            if (!oAdo.ColumnExist(conn, strFvsTreeTable, column[0]))
                                oAdo.AddColumn(conn, strFvsTreeTable, column[0], column[1], "");
                        });
                    }
                }
            }

            //new master.db file: Migrating POP tables to SQLite
            DataMgr p_dataMgr = new DataMgr();
            string strDestFile = ReferenceProjectDirectory.Trim() + "\\" + frmMain.g_oTables.m_oFIAPlot.DefaultPopTableDbFile;
            if (!System.IO.File.Exists(strDestFile))
            {
                p_dataMgr.CreateDbFile(strDestFile);
                string strConn = p_dataMgr.GetConnectionString(strDestFile);
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    con.Open();
                    //pop estimation unit table
                    if (!p_dataMgr.TableExist(con, frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName))
                    {
                        frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopEstnUnitTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName);
                    }
                    //pop eval table
                    if (!p_dataMgr.TableExist(con, frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName))
                    {
                        frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopEvalTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName);
                    }
                    //pop plot stratum assignment table
                    if (!p_dataMgr.TableExist(con, frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName))
                    {
                        frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopPlotStratumAssgnTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName);
                    }
                    //pop stratum table
                    if (!p_dataMgr.TableExist(con, frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName))
                    {
                        frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopStratumTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName);
                    }
                }
            }

        }

        public void UpdateDatasources_5_9_0()
        {
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();
            DataMgr p_dataMgr = new DataMgr();
            string[] arrPopTables = { frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName, frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName,
                                      frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName, frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName};

            // Check to see if POP tables exist in master.db; If so make sure they have modified_date column
            string strDestFile = ReferenceProjectDirectory.Trim() + "\\" + frmMain.g_oTables.m_oFIAPlot.DefaultPopTableDbFile;
            string strDestConn = p_dataMgr.GetConnectionString(strDestFile);
            bool bMissingColumn = false;
            if (System.IO.File.Exists(strDestFile))
            {
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strDestConn))
                {
                    con.Open();
                    foreach (var pTable in arrPopTables)
                    {
                        if (p_dataMgr.TableExist(con, pTable))
                        {
                            if (!p_dataMgr.FieldExist(con, "select * from " + pTable, "MODIFIED_DATE"))
                            {
                                bMissingColumn = true;
                                break;
                            }
                        }
                    }
                }
                System.Threading.Thread.Sleep(4000);    //Needed to avoid IO error further down when renaming the file
            }

            string projectDbPath = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            string strConn = oAdo.getMDBConnString(projectDbPath, "", "");
            using (var conn = new OleDbConnection(strConn))
            {
                conn.Open();
                // Remove POP tables from datasource infrastructure
                oAdo.m_strSQL = "DELETE FROM datasource WHERE TRIM(UCASE(table_type)) IN " +
                    "('POPULATION EVALUATION','POPULATION ESTIMATION UNIT','POPULATION STRATUM','POPULATION PLOT STRATUM ASSIGNMENT')";
                oAdo.SqlNonQuery(conn, oAdo.m_strSQL);

                // Remove TREE REGIONAL BIOMASS table from datasource infrastructure
                oAdo.m_strSQL = "DELETE FROM datasource WHERE TRIM(UCASE(table_type)) = " +
                    "'TREE REGIONAL BIOMASS'";
                oAdo.SqlNonQuery(conn, oAdo.m_strSQL);

                // Delete POP tables from master.mdb
                conn.Close();
                conn.ConnectionString = oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\db\\master.mdb", "", "");
                conn.Open();
                foreach (var pTable in arrPopTables)
                {
                    bool bExists = oAdo.TableExist(conn, pTable);
                    if (bExists)
                    {
                        oAdo.m_strSQL = "DROP TABLE " + pTable;
                        oAdo.SqlNonQuery(conn, oAdo.m_strSQL);
                    }
                }

                // Delete TREE_REGIONAL_BIOMASS table from master.mdb
                if (oAdo.TableExist(conn, "TREE_REGIONAL_BIOMASS"))
                {
                    oAdo.m_strSQL = "DROP TABLE TREE_REGIONAL_BIOMASS";
                    oAdo.SqlNonQuery(conn, oAdo.m_strSQL);
                }
            }
            // If older version of .db POP tables, create backup copy of master.db and recreate master.db with current pop table schema
            if (bMissingColumn)
            {
                string strCopySuffix = "_ver_control_" + DateTime.Now.ToString("MMddyyyy");
                string strBackupFileName = System.IO.Path.GetFileNameWithoutExtension(strDestFile) + strCopySuffix + ".db";
                string strDestFolder = System.IO.Path.GetDirectoryName(strDestFile);
                if (!System.IO.File.Exists(strDestFolder + "\\" + strBackupFileName))
                {
                    System.IO.File.Move(strDestFile, strDestFolder + "\\" + strBackupFileName);
                }
                p_dataMgr.CreateDbFile(strDestFile);
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strDestConn))
                {
                    con.Open();
                    //pop estimation unit table
                    if (!p_dataMgr.TableExist(con, frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName))
                    {
                        frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopEstnUnitTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName);
                    }
                    //pop eval table
                    if (!p_dataMgr.TableExist(con, frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName))
                    {
                        frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopEvalTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName);
                    }
                    //pop plot stratum assignment table
                    if (!p_dataMgr.TableExist(con, frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName))
                    {
                        frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopPlotStratumAssgnTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName);
                    }
                    //pop stratum table
                    if (!p_dataMgr.TableExist(con, frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName))
                    {
                        frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopStratumTable(p_dataMgr, con, frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName);
                    }
                }
            }

            //Update Master Cond/Plot tables
            FIA_Biosum_Manager.Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();

            int intCondTableType = oDs.getDataSourceTableNameRow("CONDITION");
            if (oDs.DataSourceTableExist(intCondTableType))
            {
                string strCondTable = oDs.m_strDataSource[intCondTableType, Datasource.TABLE].Trim();
                string strCondTableDb = oDs.m_strDataSource[intCondTableType, Datasource.PATH].Trim() + "\\" +
                                        oDs.m_strDataSource[intCondTableType, Datasource.MDBFILE].Trim();
                using (OleDbConnection conn = new OleDbConnection(oAdo.getMDBConnString(strCondTableDb, "", "")))
                {
                    conn.Open();
                    if (!oAdo.ColumnExist(conn, strCondTable, "stdorgcd"))
                        oAdo.AddColumn(conn, strCondTable, "stdorgcd", "INTEGER", "");

                    if (!oAdo.ColumnExist(conn, strCondTable, "qmd_all_inch"))
                    {
                        oDao.RenameField(strCondTableDb, strCondTable, "qmd_tot_cm", "qmd_all_inch");
                    }
                    if (!oAdo.ColumnExist(conn, strCondTable, "qmd_hwd_inch"))
                    {
                        oDao.RenameField(strCondTableDb, strCondTable, "hwd_qmd_tot_cm", "qmd_hwd_inch");
                    }
                    if (!oAdo.ColumnExist(conn, strCondTable, "qmd_swd_inch"))
                    {
                        oDao.RenameField(strCondTableDb, strCondTable, "swd_qmd_tot_cm", "qmd_swd_inch");
                    }
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The Master Cond table was not found.",
                    "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }

            int intPlotTableType = oDs.getDataSourceTableNameRow("PLOT");
            if (oDs.DataSourceTableExist(intPlotTableType))
            {
                string strPlotTable = oDs.m_strDataSource[intPlotTableType, Datasource.TABLE].Trim();
                string strPlotTableDb = oDs.m_strDataSource[intPlotTableType, Datasource.PATH].Trim() + "\\" +
                                        oDs.m_strDataSource[intPlotTableType, Datasource.MDBFILE].Trim();
                using (OleDbConnection conn = new OleDbConnection(oAdo.getMDBConnString(strPlotTableDb, "", "")))
                {
                    conn.Open();
                    if (!oAdo.ColumnExist(conn, strPlotTable, "ecosubcd"))
                        oAdo.AddColumn(conn, strPlotTable, "ecosubcd", "CHAR(7)", "");
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The Master Plot table was not found.",
                    "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }

            if (oAdo != null)
            {
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }

            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
        }

        public void UpdateDatasources_5_10_1()
        {
            DataMgr oDataMgr = new DataMgr();
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();

            // issue #287 Delete obsolete FVS_ResidTree table
            string strConnection = oDataMgr.GetConnectionString($@"{ReferenceProjectDirectory.Trim()}\{Tables.FVS.DefaultFVSTreeListDbFile}");
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
            {
                con.Open();
                if (oDataMgr.TableExist(con, Tables.FVS.DefaultFVSResidTreeTableName))
                {
                    oDataMgr.SqlNonQuery(con, $@"DROP TABLE {Tables.FVS.DefaultFVSResidTreeTableName}");
                }
            }

            //issue #286 Rename landclcd to cond_status_cd
            Datasource oDs = new Datasource();
            oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oDs.m_strDataSourceTableName = "datasource";
            oDs.m_strScenarioId = "";
            oDs.LoadTableColumnNamesAndDataTypes = false;
            oDs.LoadTableRecordCount = false;
            oDs.populate_datasource_array();

            int intCondTableType = oDs.getDataSourceTableNameRow("CONDITION");
            string[] arrFieldNames = null;
            if (oDs.DataSourceTableExist(intCondTableType))
            {
                string strCondTable = oDs.m_strDataSource[intCondTableType, Datasource.TABLE].Trim();
                string strCondTableDb = oDs.m_strDataSource[intCondTableType, Datasource.PATH].Trim() + "\\" +
                                        oDs.m_strDataSource[intCondTableType, Datasource.MDBFILE].Trim();
                oDao.getFieldNames(strCondTableDb, strCondTable, ref arrFieldNames);
                if (!arrFieldNames.Contains("cond_status_cd"))
                {
                    oDao.RenameField(strCondTableDb, strCondTable, "landclcd", "cond_status_cd");
                }                
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The Master Cond table was not found.",
                    "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }

            // issue #295: updates to harvest_costs table
            string strScenarioDir = ReferenceProjectDirectory.Trim() + "\\processor\\db";
            //open the scenario_processor_rule_definitions.mdb file
            List<string> lstScenarioDb = new List<string>();
            using (var conn = new OleDbConnection(oAdo.getMDBConnString(strScenarioDir + "\\scenario_processor_rule_definitions.mdb", "", "")))
            {
                conn.Open();
                //retrieve paths for all scenarios in the project and put them in list
                oAdo.m_strSQL = "SELECT path from scenario";
                oAdo.SqlQueryReader(conn, oAdo.m_strSQL);
                if (oAdo.m_OleDbDataReader.HasRows)
                {
                    while (oAdo.m_OleDbDataReader.Read())
                    {
                        string strPath = "";
                        if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value)
                            strPath = oAdo.m_OleDbDataReader["path"].ToString().Trim();
                        if (!String.IsNullOrEmpty(strPath))
                        {
                            //Check to see if the .mdb exists before adding it to the list
                            string strPathToMdb = strPath + "\\db\\scenario_results.mdb";
                            //sample path: C:\\workspace\\BioSum\\biosum_data\\bluemountains\\processor\\scenario1\\db\\scenario_results.mdb
                            if (System.IO.File.Exists(strPathToMdb))
                                lstScenarioDb.Add(strPathToMdb);
                        }
                    }
                    oAdo.m_OleDbDataReader.Close();
                }
            }

            // Loop through the scenario_results.mdb looking for harvest_costs table
            string strTempTable = Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + "_NEW";
            string strInsertSql = $@"INSERT INTO {strTempTable} SELECT 
                biosum_cond_id,rxpackage, rx, rxcycle, complete_cpa, harvest_cpa, chip_cpa, assumed_movein_cpa,
                place_holder, override_YN, DateTimeCreated FROM {Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName}";
            arrFieldNames = null;
            foreach (string strPath in lstScenarioDb)
            {
                oDao.getFieldNames(strPath, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, ref arrFieldNames);
                if (arrFieldNames.Contains("ideal_complete_cpa"))
                {
                    // Add columns to harvest_costs table
                    using (var conn = new OleDbConnection(oAdo.getMDBConnString(strPath, "", "")))
                    {
                        conn.Open();
                        frmMain.g_oTables.m_oProcessor.CreateHarvestCostsTable(oAdo, conn, strTempTable);
                        oAdo.SqlNonQuery(conn, strInsertSql);
                        long lngSource = oAdo.getRecordCount(conn, "SELECT COUNT(*) FROM " + strTempTable, strTempTable);
                        long lngTarget = oAdo.getRecordCount(conn, "SELECT COUNT(*) FROM " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName,
                            Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName);
                        if (lngSource == lngTarget)
                        {
                            oAdo.SqlNonQuery(conn, "DROP TABLE " + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName);
                            oDao.RenameTable(strPath, strTempTable, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, false, false);
                        }
                        else
                        {
                            System.Windows.Forms.MessageBox.Show("An error occurred while updating the harvest_costs table in " + strPath + ". Please check the database!",
                                "FIA Biosum");
                        }
                    }
                }
                if (!oDao.TableExists(strPath, Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName))
                {
                    // Add new additional_kcp_cpa
                    using (var conn = new OleDbConnection(oAdo.getMDBConnString(strPath, "", "")))
                    {
                        conn.Open();
                        frmMain.g_oTables.m_oProcessorScenarioRun.CreateAdditionalKcpCpaTable(oAdo, conn, Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName, false);
                    }
                }
            }

            if (oAdo != null && oAdo.m_OleDbConnection != null)
            {
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }

            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
        }

        // Method to compare two versions. 
        // Returns 1 if v2 is smaller, -1 
        // if v1 is smaller, 0 if equal 
        public int VersionCompare(string v1, string v2)
        {
            // vnum stores each numeric 
            // part of version 
            int vnum1 = 0, vnum2 = 0;

            // loop until both string are 
            // processed 
            for (int i = 0, j = 0; (i < v1.Length
                                    || j < v2.Length);)
            {
                // storing numeric part of 
                // version 1 in vnum1 
                while (i < v1.Length && v1[i] != '.')
                {
                    vnum1 = vnum1 * 10 + (v1[i] - '0');
                    i++;
                }

                // storing numeric part of 
                // version 2 in vnum2 
                while (j < v2.Length && v2[j] != '.')
                {
                    vnum2 = vnum2 * 10 + (v2[j] - '0');
                    j++;
                }

                if (vnum1 > vnum2)
                    return 1;
                if (vnum2 > vnum1)
                    return -1;

                // if equal, reset variables and 
                // go for next numeric part 
                vnum1 = vnum2 = 0;
                i++;
                j++;
            }
            return 0;
        }

        public string ReferenceProjectDirectory
		{
			get {return _strProjDir;}
			set {_strProjDir=value;}
		}
		public FIA_Biosum_Manager.frmMain ReferenceMainForm
		{
			get {return _oFrmMain;}
			set {_oFrmMain = value;}
		}
		
	}
}

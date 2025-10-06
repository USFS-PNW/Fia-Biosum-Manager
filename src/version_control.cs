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
                    //if ((Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) < 5) ||
                    //   (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                    //   Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) <= 2 &&
                    //   Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) == 0))
                    //    UpgradeFVSOutTreeListFiles();

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
                            //if (result == System.Windows.Forms.DialogResult.Yes)
                            //{
                            //    UpgradeFVSOutPOTFireFiles();
                            //}
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
                                //Update_5_7_0();
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
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) == 10 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) == 1) &&
                        (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) >= 9))
                    {
                        UpdateDatasources_5_10_1();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    // Upgrade to 5.10.1 to 5.11.1 (Sequence numbers, Optimizer, Processor to SQLite, new field on Psites table)
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) == 11 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) == 1) &&
                        (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) == 10 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) == 1))
                    {
                        UpdateDatasources_5_11_0();
                        UpdateDatasources_5_11_1();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    // Upgrade from 5.11.0 to 5.11.1 (new field on Psites table)
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) == 11 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) == 1) &&
                        (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) > 10 ))
                    {
                        UpdateDatasources_5_11_1();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    // Upgrade from 5.11.1 to 5.11.2 (rx, rxPackage, harvest_methods, FVSIn move to SQLite)
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) == 11 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) == 2) &&
                        (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) == 11 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) == 1))
                    {
                        UpdateDatasources_5_11_2();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    // Upgraded from 5.11.2 to 5.12.0 (master.mdb to SQLite)
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) == 12 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) == 0) &&
                        (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) == 11 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) == 2))
                    {
                        UpdateDatasources_5_12_0();
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
                //frmMain.g_sbpInfo.Text = "Version Update: Checking Project Datasource Tables...Stand by";
                //CheckProjectDatasourceTables();
                //frmMain.g_sbpInfo.Text = "Version Update: Checking Pre-Populated Reference Tables...Stand by";
                //CheckProjectReferenceDatasourceTables();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Core Scenario Datasource Table Records...Stand by";
                CheckCoreScenarioDatasourceTableRecords();
                frmMain.g_sbpInfo.Text = "Version Update: Checking Processor Scenario Datasource Table Records...Stand by";
                CheckProcessorScenarioDatasourceTableRecords();


                //frmMain.g_sbpInfo.Text = "Version Update: Checking Rx Values...Stand By";
                //CheckRxValues();
                //frmMain.g_sbpInfo.Text = "Version Update: Checking Fvs Out Pre-Post Rx Values...Stand By";
                //CheckFVSOutPrePostValues();

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
                this.ReferenceMainForm.frmProject.uc_project1.CreateOptimizerScenarioRuleDefinitionSqliteDbAndTables(strFile1);
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
								frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioProcessorScenarioSelectTable_Access(oAdo,oConn,strTableName);

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
							frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioProcessorScenarioSelectTable_Access(oAdo,oConn,strTableName);
                            break;
					}
				}
			}

			oAdo.CloseConnection(oAdo.m_OleDbConnection);
			oAdo=null;
			oDao=null;
			
		}


		//private void CheckProjectReferenceDatasourceTables()
		//{

		//	int x,y;
		//	FIA_Biosum_Manager.Datasource oDs = new Datasource();
		//	oDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
		//	oDs.m_strDataSourceTableName = "datasource";
		//	oDs.m_strScenarioId="";
		//	oDs.LoadTableColumnNamesAndDataTypes=false;
		//	oDs.LoadTableRecordCount=false;
		//	oDs.populate_datasource_array();

		//	ado_data_access oAdo=new ado_data_access();
		//	ado_data_access oAdoCurrent=null;

		//	dao_data_access oDao= new dao_data_access();
		//	string strCurrDbFile="";
		//	string strDbFile="";
		//	string strTempDbFile="";
		//	string strTempTableName="";
		//	System.Data.OleDb.OleDbConnection oConn=null;


		//	//copy ref_master.mdb to project db directory if file does not exist
		//	if (System.IO.File.Exists(ReferenceProjectDirectory.Trim() + "\\db\\ref_master.mdb")==false)
		//		 System.IO.File.Copy(frmMain.g_oEnv.strAppDir + "\\db\\ref_master.mdb",this.ReferenceProjectDirectory + "\\db\\ref_master.mdb",true);

		//	//open the project db file
		//	oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" + 
		//		frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile,"",""));

		//	oAdo.m_strSQL = "SELECT * FROM datasource " + 
		//		            "WHERE table_type IS NOT NULL AND " + 
  //                                "UCASE(TRIM(table_type)) " + 
  //                                "IN ('OWNER GROUPS'," + 
  //                                    "'TREE SPECIES'," + 
  //                                    "'FVS TREE SPECIES'," + 
  //                                    "'FIADB FVS VARIANT'," + 
  //                                    "'INVENTORIES'," + 
  //                                    "'TREATMENT PRESCRIPTION CATEGORIES'," + 
  //                                    "'TREATMENT PRESCRIPTION SUBCATEGORIES'," + 
  //                                    "'HARVEST METHODS'," + 
  //                                    "'FVS WESTERN TREE SPECIES TRANSLATOR'," +
  //                                    "'FVS EASTERN TREE SPECIES TRANSLATOR'" + 
  //                                    ") ORDER BY path,file";
		//	oAdo.SqlQueryReader(oAdo.m_OleDbConnection,oAdo.m_strSQL);

		//	if (oAdo.m_OleDbDataReader.HasRows)
		//	{
		//		while (oAdo.m_OleDbDataReader.Read())
		//		{
		//			if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value &&
		//				oAdo.m_OleDbDataReader["file"] != System.DBNull.Value)
		//			{
		//				x=oDs.getDataSourceTableNameRow(oAdo.m_OleDbDataReader["table_type"].ToString().Trim());


						
		//				strDbFile = oAdo.m_OleDbDataReader["path"].ToString().Trim() + "\\" + 
		//					oAdo.m_OleDbDataReader["file"].ToString().Trim();

		//				if (strDbFile.Trim().ToUpper() != strCurrDbFile.Trim().ToUpper())
		//				{
		//					if (oAdoCurrent!=null) 
		//					{  
		//						if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"owner_groups_temp"))
		//							 oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE owner_groups_temp");
  //                              if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "harvest_methods_temp"))
  //                                  oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE harvest_methods_temp");
		//						oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
		//					}
		//					else 
		//					{
		//						strTempDbFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");
		//						oDao.CreateMDB(strTempDbFile);
		//						System.IO.File.Copy(frmMain.g_oEnv.strAppDir.Trim() + "\\db\\ref_master.mdb",strTempDbFile,true);

		//						oAdoCurrent = new ado_data_access();
		//					}

		//					//make sure the db file exists
		//					if (System.IO.File.Exists(strDbFile)==false)
		//					{
		//						oDao.CreateMDB(strDbFile);
		//					}
							
		//					//create a link to all the pre-populated reference tables
		//					oDao.CreateTableLink(strDbFile,"owner_groups_temp",strTempDbFile,"owner_groups",true);
		//                    oDao.CreateTableLink(strDbFile, "harvest_methods_temp", strTempDbFile, "harvest_methods", true);
		//					oAdoCurrent.OpenConnection(oAdo.getMDBConnString(strDbFile,"",""));
		//					strCurrDbFile=strDbFile;

		//				}
  //                      if (oDs.m_strDataSource[x, Datasource.TABLESTATUS] == "NF") //NF=table not found
  //                      {
  //                          //create the table
  //                          switch (oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim().ToUpper())
  //                          {

  //                              case "OWNER GROUPS":
  //                                  oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM owner_groups_temp";
  //                                  oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
  //                                  oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
  //                                      oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
  //                                      oDs.m_strDataSource[x, Datasource.TABLE].Trim());
  //                                  break;
  //                              //version 5 additions
  //                              case "TREATMENT PRESCRIPTION CATEGORIES":
  //                                  oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM fvs_rx_category_temp";
  //                                  oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
  //                                  oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
  //                                      oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
  //                                      oDs.m_strDataSource[x, Datasource.TABLE].Trim());
  //                                  break;
  //                              case "TREATMENT PRESCRIPTION SUBCATEGORIES":
  //                                  oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM fvs_rx_subcategory_temp";
  //                                  oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
  //                                  oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
  //                                      oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
  //                                      oDs.m_strDataSource[x, Datasource.TABLE].Trim());
  //                                  break;
  //                              case "HARVEST METHODS":
  //                                  oAdoCurrent.m_strSQL = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " FROM harvest_methods_temp";
  //                                  oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
  //                                  oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
  //                                      oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
  //                                      oDs.m_strDataSource[x, Datasource.TABLE].Trim());
  //                                  break;
  //                          }
  //                      }
  //                      else
  //                      {
  //                          //update columns and data
  //                          strTempTableName = "";
  //                          strInsertSql = "";
  //                          strUpdateSql = "";
  //                          string[] strColumnsArray = null;
  //                          string strColumnsList = "";
  //                          switch (oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim().ToUpper())
  //                          {
  //                              case "OWNER GROUPS":
  //                                  oAdoCurrent.m_strSQL = "DROP TABLE " + oDs.m_strDataSource[x, Datasource.TABLE];
  //                                  oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
  //                                  strInsertSql = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE] + " FROM owner_groups_temp";
  //                                  oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
  //                                  oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
  //                                          oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
  //                                          oDs.m_strDataSource[x, Datasource.TABLE].Trim());
  //                                  break;
  //                              case "INVENTORIES":
  //                                  oAdoCurrent.m_strSQL = "DROP TABLE " + oDs.m_strDataSource[x, Datasource.TABLE];
  //                                  oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
  //                                  strInsertSql = "SELECT * INTO " + oDs.m_strDataSource[x, Datasource.TABLE] + " FROM inventories_temp";
  //                                  oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
  //                                  oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
  //                                      oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
  //                                      oDs.m_strDataSource[x, Datasource.TABLE].Trim());
  //                                  break;
  //                              case "HARVEST METHODS":
  //                                      oAdoCurrent.m_strSQL = "DROP TABLE " + oDs.m_strDataSource[x, Datasource.TABLE];
  //                                      oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
  //                                      frmMain.g_oTables.m_oReference.CreateHarvestMethodsTable(oAdoCurrent, oAdoCurrent.m_OleDbConnection, oDs.m_strDataSource[x, Datasource.TABLE]);
  //                                      strTempTableName = "harvest_methods_temp";
  //                                      strColumnsList="";
  //                                      //insert any new tree species records
  //                                      strColumnsArray = oAdoCurrent.getFieldNamesArray(oAdoCurrent.m_OleDbConnection, "SELECT * FROM " + strTempTableName);
  //                                      for (y = 0; y <= strColumnsArray.Length - 1; y++)
  //                                      {
  //                                          strColumnsList = strColumnsList + "a." + strColumnsArray[y].Trim() + ",";
  //                                      }
  //                                      strColumnsList = strColumnsList.Substring(0, strColumnsList.Length - 1);
  //                                      strInsertSql = "INSERT INTO " +
  //                                          oDs.m_strDataSource[x, Datasource.TABLE] + " " +
  //                                          "SELECT " + strColumnsList + " " +
  //                                          "FROM " + strTempTableName + " a ";
                                           
  //                                      oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, strInsertSql);
  //                                  break;
  //                          }
  //                      }
		//			}
		//		}
		//	}
		//	if (oAdoCurrent!=null) 
		//	{
		//		if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection,"owner_groups_temp"))
		//			oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE owner_groups_temp");
  //              if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "harvest_methods_temp"))
  //                  oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE harvest_methods_temp");
  //              if (oAdoCurrent.TableExist(oAdoCurrent.m_OleDbConnection, "fvs_rx_category_temp"))
  //                  oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE fvs_rx_category_temp");
		//		oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
		//	}
		//	oAdo.CloseConnection(oAdo.m_OleDbConnection);
		//	if (oConn != null)
		//	{
		//		oAdo.CloseConnection(oConn);
		//	}
		//	if (oDao!=null)
		//	{
		//		oDao.m_DaoWorkspace.Close();
		//		oDao=null; 
		//	}
		//	oAdoCurrent=null;
		//	oAdo=null;
  //          strCurrDbFile = "";
  //          strDbFile = "";
  //          strTempDbFile = "";
  //          strTempTableName = "";
  //          oConn = null;
  //          //
  //          //fvs commands table
  //          //
            

            
  //          oAdo = new ado_data_access();
  //          oDao = new dao_data_access();
  //          string strNewDbFile = "";
  //          string strOldDbFile = "";

  //          //open the project db file
  //          oAdo.OpenConnection(oAdo.getMDBConnString(ReferenceProjectDirectory.Trim() + "\\" +
  //              frmMain.g_oTables.m_oProject.DefaultProjectTableDbFile, "", ""));

  //          oAdo.m_strSQL = "SELECT * FROM datasource " +
  //                          "WHERE table_type IS NOT NULL AND " +
  //                                "UCASE(TRIM(table_type)) " +
  //                                "IN ('FVS COMMANDS') ORDER BY path,file";
  //          oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);

  //          if (oAdo.m_OleDbDataReader.HasRows)
  //          {
  //              while (oAdo.m_OleDbDataReader.Read())
  //              {
  //                  if (oAdo.m_OleDbDataReader["path"] != System.DBNull.Value &&
  //                      oAdo.m_OleDbDataReader["file"] != System.DBNull.Value)
  //                  {
  //                      x = oDs.getDataSourceTableNameRow(oAdo.m_OleDbDataReader["table_type"].ToString().Trim());



  //                      strOldDbFile = oAdo.m_OleDbDataReader["path"].ToString().Trim() + "\\" +
  //                          oAdo.m_OleDbDataReader["file"].ToString().Trim();

  //                      if (System.IO.File.Exists(strOldDbFile))
  //                      {
  //                          oAdoCurrent = new ado_data_access();
  //                          //
  //                          //get all table names in the old file
  //                          //
  //                          string[] strTablesArray = null;
  //                          oDao.getTableNames(strOldDbFile, ref strTablesArray);
  //                          //
  //                          //get the new file name
  //                          //
  //                          strNewDbFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
  //                          //
  //                          //create the new file and open it
  //                          //
  //                          oDao.CreateMDB(strNewDbFile);
  //                          oDao.OpenDb(strNewDbFile);
  //                          //
  //                          //create a link to all tables from the old file to the new file
  //                          //
  //                          if (strTablesArray != null)
  //                          {
                               
  //                              for (y = 0; y <= strTablesArray.Length - 1; y++)
  //                              {
  //                                  if (strTablesArray[y] == null) break;
  //                                  if (strTablesArray[y].Trim().Length > 0)
  //                                  {
  //                                      //do not link to the table we are replacing
  //                                      if (strTablesArray[y].Trim().ToUpper() !=
  //                                          oDs.m_strDataSource[x, Datasource.TABLE].Trim().ToUpper())
  //                                      {
  //                                          oDao.CreateTableLink(oDao.m_DaoDatabase, strTablesArray[y].Trim() + "_temp", strOldDbFile, strTablesArray[y].Trim());
  //                                      }


  //                                  }
  //                              }
  //                              oDao.m_DaoWorkspace.Close();
  //                              oDao = null;
  //                              //open oledb connection to new file
  //                              oAdoCurrent.OpenConnection(oAdoCurrent.getMDBConnString(strNewDbFile,"",""));
  //                              for (y = 0; y <= strTablesArray.Length - 1; y++)
  //                              {
  //                                  if (strTablesArray[y] == null) break;
  //                                  //import linked tables
  //                                  if (strTablesArray[y].Trim().ToUpper() == oDs.m_strDataSource[x, Datasource.TABLE].Trim().ToUpper())
  //                                  {
                                        
  //                                  }
  //                                  else
  //                                  {
  //                                      oAdoCurrent.m_strSQL = "SELECT * INTO " + strTablesArray[y].Trim() + " " +
  //                                                             "FROM " + strTablesArray[y].Trim() + "_temp";
  //                                      oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
  //                                      //drop the linked table
  //                                      oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection,"DROP TABLE " + strTablesArray[y].Trim() + "_temp");
  //                                  }
  //                              }
                               
  //                          }
  //                          //
  //                          //import linked fvs_commands table
  //                          //
  //                          if (oAdoCurrent.m_OleDbConnection.State == System.Data.ConnectionState.Closed)
  //                              oAdoCurrent.OpenConnection(oAdoCurrent.getMDBConnString(strNewDbFile, "", ""));
  //                          oAdoCurrent.m_strSQL = "SELECT * " + 
  //                                                 "INTO " + oDs.m_strDataSource[x, Datasource.TABLE].Trim() + " " + 
  //                                                 "FROM fvs_commands_temp";
  //                          oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, oAdoCurrent.m_strSQL);
  //                          oDs.SetPrimaryIndexesAndAutoNumbers(oAdoCurrent, oAdoCurrent.m_OleDbConnection,
  //                              oDs.m_strDataSource[x, Datasource.TABLETYPE].Trim(),
  //                              oDs.m_strDataSource[x, Datasource.TABLE].Trim());
  //                          oAdoCurrent.SqlNonQuery(oAdoCurrent.m_OleDbConnection, "DROP TABLE fvs_commands_temp");
  //                          oAdoCurrent.CloseConnection(oAdoCurrent.m_OleDbConnection);
  //                          //
  //                          //delete old file and copy new file
  //                          //
  //                          System.IO.File.Delete(strOldDbFile);
  //                          System.IO.File.Copy(strNewDbFile,strOldDbFile,true);

  //                      }

                        
  //                  }
  //              }
  //          }
            
  //          oAdo.CloseConnection(oAdo.m_OleDbConnection);
  //          if (oConn != null)
  //          {
  //              oAdo.CloseConnection(oConn);
  //          }
  //          if (oDao != null)
  //          {
  //              oDao.m_DaoWorkspace.Close();
  //              oDao = null;
  //          }
  //          oAdoCurrent = null;
  //          oAdo = null;
		//}



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
                frmMain.g_oFrmMain.frmProject.uc_project1.CreateProcessorScenarioRuleDefinitionAccessDbAndTables(oDsScenario.m_strDataSourceMDBFile);

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
            //uc_fvs_output_prepost_seqnum.InitializePrePostSeqNumTablesAccess(oAdo, ReferenceProjectDirectory.Trim() + "\\" + Tables.FVS.DefaultFVSPrePostSeqNumTableMdbFile);

           

            oAdo.CloseConnection(oAdo.m_OleDbConnection);

           

            oAdo = null;
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


            oAdo.m_strSQL = "DELETE FROM datasource WHERE TRIM(UCASE(table_type))='PLOT, CONDITION AND TREATMENT RECORD AUDIT'";
            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

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
                    strPath = oDs.m_strDataSource[intTableType,Datasource.PATH].Trim() + "\\" + oDs.m_strDataSource[intTableType,Datasource.DBFILE].Trim();
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
            string strFileName = oDs.m_strDataSource[intHarvestMethodTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
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
                //oDao.CreateTableLink(strDestinationDbFile, strDestinationSpeciesTableName + "_worktable", strSourceDbFile, Tables.Reference.DefaultTreeSpeciesTableName);
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
            string strFileName = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
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
            strFileName = oDs.m_strDataSource[intFvsTreeSpeciesTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
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
            string strFileName = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            string strFileStatus = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            string strTargetTable = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strTableStatus = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();

            if (strFileStatus == "F" && strTableStatus == "F")
            {
                oDao.RenameTable(strDirectoryPath + "\\" + strFileName, strTargetTable, strTargetTable + strTableSuffix, true, false);
            }

            int intTreeSpeciesTable = oDs.getValidTableNameRow("Tree Species");
            strDirectoryPath = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strFileName = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            strFileStatus = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            strTargetTable = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            strTableStatus = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();

            if (strFileStatus == "F" && strTableStatus == "F")
            {
                oDao.RenameTable(strDirectoryPath + "\\" + strFileName, strTargetTable, strTargetTable + strTableSuffix, true, false);
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
            DataMgr oDataMgr = new DataMgr();

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
            string strFileName = oDs.m_strDataSource[intTreeSpeciesTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
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
            strFileName = oDs.m_strDataSource[intCondTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
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
            strFileName = oDs.m_strDataSource[intFvsVariantTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            strFileStatus = oDs.m_strDataSource[intFvsVariantTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            strTargetTable = oDs.m_strDataSource[intFvsVariantTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            strTableStatus = oDs.m_strDataSource[intFvsVariantTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
            if (strFileStatus == "F" && strTableStatus == "F")
            {
                oDao.RenameTable(strDirectoryPath + "\\" + strFileName, strTargetTable, strTargetTable + strTableSuffix, true, false);
            }

            frmMain.g_sbpInfo.Text = "Version Update: Adding fvsloccode to Plot table ...Stand by";
            int intPlotTable = oDs.getValidTableNameRow("Plot");
            strDirectoryPath = oDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strFileName = oDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
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
            DataMgr oDataMgr = new DataMgr();

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
            string strDestFileSqlite = ReferenceProjectDirectory.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile;
            if (!System.IO.File.Exists(strDestFileSqlite))
            {
                oDataMgr.CreateDbFile(strDestFileSqlite);
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
            string strFileName = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            string strFileStatus = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.FILESTATUS].Trim();
            string strTargetTable = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strTableStatus = oDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();

            if (strFileStatus == "F" && strTableStatus == "F")
            {
                oDao.RenameTable(strDirectoryPath + "\\" + strFileName, strTargetTable, strTargetTable + strTableSuffix, true, false);
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
                strFileName = oProjectDs.m_strDataSource[intTravelTimesTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
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
            string strPlotMdbFile = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
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
            strPlotMdbFile = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
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
                    strHarvestMethodFileName = oProjectDs.m_strDataSource[intHarvestMethodsTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
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
                    string strFileName = oProjectDs.m_strDataSource[intRxHarvestTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
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
                                        oDs.m_strDataSource[intTreeTableType, Datasource.DBFILE].Trim();
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
                                        oDs.m_strDataSource[intCondTableType, Datasource.DBFILE].Trim();
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
                                        oDs.m_strDataSource[intPlotTableType, Datasource.DBFILE].Trim();
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
                                        oDs.m_strDataSource[intCondTableType, Datasource.DBFILE].Trim();
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
                                        oDs.m_strDataSource[intPlotTableType, Datasource.DBFILE].Trim();
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
                                        oDs.m_strDataSource[intCondTableType, Datasource.DBFILE].Trim();
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

        public void UpdateDatasources_5_11_0()
        {
            DataMgr oDataMgr = new DataMgr();
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();
            ODBCMgr odbcmgr = new ODBCMgr();
            utils oUtils = new utils();
            env oEnv = new env();
            Datasource oProjectDs = new Datasource();

            // MIGRATING SEQUENCE NUMBER SETTINGS TO fvs_master.db
            string strPrePostSeqNumLink = $@"{Tables.FVS.DefaultFVSPrePostSeqNumTable}_1";
            string strRxPackageAssignLink = $@"{Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable}_1";

            oProjectDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oProjectDs.m_strDataSourceTableName = "datasource";
            oProjectDs.m_strScenarioId = "";
            oProjectDs.LoadTableColumnNamesAndDataTypes = false;
            oProjectDs.LoadTableRecordCount = false;
            oProjectDs.populate_datasource_array();
            int intSeqNumDefs = oProjectDs.getValidTableNameRow(Datasource.TableTypes.SeqNumDefinitions);
            int intSeqNumRxPkgAssign = oProjectDs.getValidTableNameRow(Datasource.TableTypes.SeqNumRxPackageAssign);
            if (intSeqNumDefs > -1 && intSeqNumRxPkgAssign > -1)
            {
                if (!System.IO.File.Exists($@"{ReferenceProjectDirectory.Trim()}\{Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile}"))
                {
                    oDataMgr.CreateDbFile($@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\{Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile}");
                }
                string dbConn = oDataMgr.GetConnectionString($@"{ReferenceProjectDirectory.Trim()}\{Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile}");
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dbConn))
                {
                    conn.Open();
                    if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultFVSPrePostSeqNumTable))
                    {
                        frmMain.g_oTables.m_oFvs.CreateFVSOutputSQLitePrePostSeqNumTable(oDataMgr, conn, Tables.FVS.DefaultFVSPrePostSeqNumTable);
                    }
                    if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable))
                    {
                        frmMain.g_oTables.m_oFvs.CreateFVSOutputPrePostSQLiteSeqNumRxPackageAssgnTable(oDataMgr, conn, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable);
                    }
                }
                // Create ODBC entry for the new SQLite fvs_master.db file
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.FvsMasterDbDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.FvsMasterDbDsnName);
                }
                odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.FvsMasterDbDsnName, $@"{ReferenceProjectDirectory.Trim()}\{Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile}");

                string strSeqNumDefsTable = oProjectDs.m_strDataSource[intSeqNumDefs, Datasource.TABLE].Trim();
                string strSeqNumRxPkgAssignTable = oProjectDs.m_strDataSource[intSeqNumRxPkgAssign, Datasource.TABLE].Trim();
                oDao.CreateSQLiteTableLink(oProjectDs.getFullPathAndFile(Datasource.TableTypes.SeqNumDefinitions), Tables.FVS.DefaultFVSPrePostSeqNumTable,
                    strPrePostSeqNumLink, ODBCMgr.DSN_KEYS.FvsMasterDbDsnName, ReferenceProjectDirectory.Trim() + "\\" + Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile, true);
                oDao.CreateSQLiteTableLink(oProjectDs.getFullPathAndFile(Datasource.TableTypes.SeqNumDefinitions), Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable,
                    strRxPackageAssignLink, ODBCMgr.DSN_KEYS.FvsMasterDbDsnName, ReferenceProjectDirectory.Trim() + "\\" + Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile, true);
                string strCopyConn = oAdo.getMDBConnString(oAdo.getMDBConnString(oProjectDs.getFullPathAndFile(Datasource.TableTypes.SeqNumDefinitions), "", ""), "", "");
                int i = 0;
                int intError = 0;
                using (var oCopyConn = new System.Data.OleDb.OleDbConnection(strCopyConn))
                {
                    oCopyConn.Open();
                    do
                    {
                        // break out of loop if it runs too long
                        if (i > 20)
                        {
                            System.Windows.Forms.MessageBox.Show("An error occurred while trying to migrate sequence number settings! ", "FIA Biosum");
                            break;
                        }
                        System.Threading.Thread.Sleep(1000);
                        i++;
                    }
                    while (!oAdo.TableExist(oCopyConn, strRxPackageAssignLink));

                    oAdo.m_strSQL = $@"INSERT INTO {strPrePostSeqNumLink} SELECT * FROM {strSeqNumDefsTable}";
                    oAdo.SqlNonQuery(oCopyConn, oAdo.m_strSQL);
                    if (oAdo.m_intError == 0)
                    {
                        oAdo.m_strSQL = $@"INSERT INTO {strRxPackageAssignLink} SELECT * FROM {strSeqNumRxPkgAssignTable}";
                        oAdo.SqlNonQuery(oCopyConn, oAdo.m_strSQL);
                        intError = oAdo.m_intError;
                    }
                    else
                    {
                        intError = oAdo.m_intError;
                    }

                    if (oAdo.TableExist(oCopyConn, strPrePostSeqNumLink))
                    {
                        oAdo.SqlNonQuery(oCopyConn, "DROP TABLE " + strPrePostSeqNumLink);
                    }
                    if (oAdo.TableExist(oCopyConn, strRxPackageAssignLink))
                    {
                        oAdo.SqlNonQuery(oCopyConn, "DROP TABLE " + strRxPackageAssignLink);
                    }
                }
                if (intError == 0)
                {
                    // Update entries in project data sources table
                    string strMasterPath = $@"{ReferenceProjectDirectory.Trim()}\{System.IO.Path.GetDirectoryName(Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile)}";
                    string strFvsMasterDb = System.IO.Path.GetFileName(Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile);
                    oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.SeqNumDefinitions, strMasterPath, strFvsMasterDb, Tables.FVS.DefaultFVSPrePostSeqNumTable);
                    oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.SeqNumRxPackageAssign, strMasterPath, strFvsMasterDb, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable);
                }
            }

            // MIGRATING SETTINGS FROM scenario_processor_rule_definitions.mdb TO scenario_processor_rule_definitions.db
            string targetDbFile = ReferenceProjectDirectory.Trim() +
                @"\processor\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            string sourceDbFile = ReferenceProjectDirectory.Trim() +
                @"\processor\" + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodDbFile;
            if (System.IO.File.Exists(targetDbFile) == false)
            {
                frmMain.g_oFrmMain.frmProject.uc_project1.CreateProcessorScenarioRuleDefinitionDbAndTables(targetDbFile);
            }

            try
            {
                string[] arrTargetTables = { };
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(targetDbFile)))
                {
                    conn.Open();
                    arrTargetTables = oDataMgr.getTableNames(conn);
                    if (arrTargetTables.Length < 1)
                    {
                        System.Windows.Forms.MessageBox.Show("Target SQLite tables could not be created. Migration stopped!!", "FIA Biosum");
                        return;
                    }

                    // custom processing for scenario_additional_harvest_costs
                    string[] strSourceColumnsArray = new string[0];
                    string strAddCpaTableName = Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName;
                    oDao.getFieldNames(sourceDbFile, strAddCpaTableName, ref strSourceColumnsArray);
                    foreach (string strColumn in strSourceColumnsArray)
                    {
                        if (!oDataMgr.ColumnExists(conn, strAddCpaTableName, strColumn))
                        {
                            oDataMgr.AddColumn(conn, strAddCpaTableName, strColumn, "DOUBLE", "");
                        }
                    }
                }

                // Check to see if the input SQLite DSN exists and if so, delete so we can add
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName);
                }
                odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName, targetDbFile);

                // Create temporary database
                string strTempAccdb = oUtils.getRandomFile(oEnv.strTempDir, "accdb");
                oDao.CreateMDB(strTempAccdb);

                // Link all the target tables to the database
                for (int i = 0; i < arrTargetTables.Length; i++)
                {
                    oDao.CreateSQLiteTableLink(strTempAccdb, arrTargetTables[i], arrTargetTables[i] + "_1",
                        ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName, targetDbFile);
                }
                oDao.CreateTableLinks(strTempAccdb, sourceDbFile);  // Link all the source tables to the database

                List<string> lstScenarioPaths = new List<string>();
                string strCopyConn = oAdo.getMDBConnString(strTempAccdb, "", "");
                using (var copyConn = new System.Data.OleDb.OleDbConnection(strCopyConn))
                {
                    copyConn.Open();
                    foreach (var strTable in arrTargetTables)
                    {
                        oAdo.m_strSQL = "INSERT INTO " + strTable + "_1" +
                        " SELECT * FROM " + strTable;
                        oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    }

                    if (oAdo.m_intError == 0)
                    {
                        // Set file (database) field to new Sqlite DB
                        string newDbFile = System.IO.Path.GetFileName(Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile);
                        oAdo.m_strSQL = "UPDATE scenario_1 set file = '" +
                            newDbFile + "'";
                        oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    }

                    //retrieve paths for all scenarios in the project and put them in list
                    oAdo.m_strSQL = "SELECT scenario_id from scenario";
                    oAdo.SqlQueryReader(copyConn, oAdo.m_strSQL);
                    if (oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (oAdo.m_OleDbDataReader.Read())
                        {
                            string strScenarioId = "";
                            if (oAdo.m_OleDbDataReader["scenario_id"] != System.DBNull.Value)
                                strScenarioId = oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim();
                            if (!String.IsNullOrEmpty(strScenarioId))
                            {
                                //Check to see if the .mdb exists before adding it to the list
                                string strPath = $@"{ReferenceProjectDirectory.Trim()}\processor\{strScenarioId}";
                                string strPathToMdb = strPath + "\\db\\scenario_results.mdb";
                                //sample path: C:\\workspace\\BioSum\\biosum_data\\bluemountains\\processor\\scenario1\\db\\scenario_results.mdb
                                if (System.IO.File.Exists(strPathToMdb))
                                    lstScenarioPaths.Add(strPath);
                            }
                        }
                        oAdo.m_OleDbDataReader.Close();
                    }
                }

                // Create tables in scenario_results.db if missing
                foreach (var sPath in lstScenarioPaths)
                {
                    string strScenarioDbPath = $@"{sPath}\{Tables.ProcessorScenarioRun.DefaultScenarioResultsTableDbFile}";
                    if (!System.IO.File.Exists(strScenarioDbPath))
                    {
                        oDataMgr.CreateDbFile(strScenarioDbPath);
                    }
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strScenarioDbPath)))
                    {
                        conn.Open();
                        if (!oDataMgr.TableExist(conn, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName))
                        {
                            frmMain.g_oTables.m_oProcessor.CreateSqliteHarvestCostsTable(oDataMgr,
                                conn, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName);
                        }
                        if (!oDataMgr.TableExist(conn, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName))
                        {
                            frmMain.g_oTables.m_oProcessor.CreateSqliteTreeVolValSpeciesDiamGroupsTable(oDataMgr,
                                conn, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, true);
                        }
                        if (!oDataMgr.TableExist(conn, Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName))
                        {
                            frmMain.g_oTables.m_oProcessorScenarioRun.CreateSqliteAdditionalKcpCpaTable(oDataMgr,
                                conn, Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName, false);
                        }
                    }
                }

                // Add SQLite OpCost config file to db directory
                if (System.IO.File.Exists(frmMain.g_oEnv.strAppDir + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile))
                {
                    if (!System.IO.File.Exists(ReferenceProjectDirectory.Trim() + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile))
                    {
                        System.IO.File.Copy(frmMain.g_oEnv.strAppDir + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile,
                            ReferenceProjectDirectory.Trim() + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile);
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show($@"The OpCost configuration file is missing from the AppData directory: {frmMain.g_oEnv.strAppDir + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile}");
                }

                // PUT OPTIMIZER AND GIS DATA MIGRATION CODE HERE. IF YOU NEED DATA THAT ISN'T IN THIS METHOD, WE CAN PASS IT
                // VIA THE METHOD CALL FOR NOW

                // MIGRATE CALCULATED VARIABLES
                string strCalculatedVariablesPathAndDbFile = this.ReferenceProjectDirectory + "\\" + Tables.OptimizerDefinitions.DefaultSqliteDbFile;
                if (!System.IO.File.Exists(strCalculatedVariablesPathAndDbFile))
                {
                    // Create SQLite copy of optimizer_definitions database
                    string variablesSourceFile = frmMain.g_oEnv.strAppDir + "\\db\\" +
                        System.IO.Path.GetFileName(Tables.OptimizerDefinitions.DefaultSqliteDbFile);
                    System.IO.File.Copy(variablesSourceFile, strCalculatedVariablesPathAndDbFile);

                    // Check to see if the input SQLite DSN exists for optimizer_definitions and if so, delete so we can add
                    if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName))
                    {
                        odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName);
                    }
                    odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName, strCalculatedVariablesPathAndDbFile);

                    // Create temporary database for optimizer_definitions
                    strTempAccdb = oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
                    oDao.CreateMDB(strTempAccdb);

                    // Create table links for transferring data for optimizer_definitions
                    string targetEcon = Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName + "_1";
                    string targetFvs = Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + "_1";
                    string targetVariables = Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + "_1";

                    string strCalculatedVariablesPathAndAccdbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerDefinitions.DefaultDbFile;
                    // Link to all tables in source database for optimizer_definitons
                    oDao.CreateTableLinks(strTempAccdb, strCalculatedVariablesPathAndAccdbFile);
                    oDao.CreateSQLiteTableLink(strTempAccdb, Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName, targetEcon,
                        ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName, strCalculatedVariablesPathAndDbFile);
                    oDao.CreateSQLiteTableLink(strTempAccdb, Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName, targetFvs,
                        ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName, strCalculatedVariablesPathAndDbFile);
                    oDao.CreateSQLiteTableLink(strTempAccdb, Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName, targetVariables,
                        ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName, strCalculatedVariablesPathAndDbFile);

                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strCalculatedVariablesPathAndDbFile)))
                    {
                        conn.Open();
                        // Delete any existing data from SQLite tables
                        string defaultEcon = Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName;
                        string defaultVariables = Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName;

                        oDataMgr.m_strSQL = "DELETE FROM " + defaultEcon;
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                        oDataMgr.m_strSQL = "DELETE FROM " + defaultVariables;
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    }

                    strCopyConn = oAdo.getMDBConnString(strTempAccdb, "", "");
                    using (var copyConn = new OleDbConnection(strCopyConn))
                    {
                        copyConn.Open();

                        oAdo.m_strSQL = "INSERT INTO " + targetEcon +
                                          " SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName;
                        oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);

                        oAdo.m_strSQL = "INSERT INTO " + targetFvs +
                          " SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName;
                        oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);

                        oAdo.m_strSQL = "INSERT INTO " + targetVariables +
                            " SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName;
                        oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    }
                }

                    // MIGRATE SCENARIO_OPTIMIZER_RULES_DEFINITIONS DATABASE
                    string scenarioAccessFile = this.ReferenceProjectDirectory + "\\" +
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;

                    // Create SQLite copy of scenario_optimizer_rule_definitions database
                    string scenarioSqliteFile = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory + "\\" +
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
                    //@ToDo: Don't have code
                    frmMain.g_oFrmMain.frmProject.uc_project1.CreateOptimizerScenarioRuleDefinitionSqliteDbAndTables(scenarioSqliteFile);

                    // Check to see if the input SQLite DSN exists and if so, delete so we can add
                    if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.OptimizerRuleDefinitionsDsnName))
                    {
                        odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.OptimizerRuleDefinitionsDsnName);
                    }
                    odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.OptimizerRuleDefinitionsDsnName, scenarioSqliteFile);

                // Set new temporary database
                strTempAccdb = oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
                    oDao.CreateMDB(strTempAccdb);

                    // Create table links for transferring data
                    string[] sourceTables = {Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName,
                        Tables.Scenario.DefaultScenarioDatasourceTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioHarvestCostColumnsTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLastTieBreakRankTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName};
                    string[] targetTables = new string[15];
                    foreach (string sourceTableName in sourceTables)
                    {
                        targetTables[Array.IndexOf(sourceTables, sourceTableName)] = sourceTableName + "_1";
                    }
                    // Link to all tables in source database
                    oDao.CreateTableLinks(strTempAccdb, scenarioAccessFile);
                    foreach (string targetTableName in targetTables)
                    {
                        oDao.CreateSQLiteTableLink(strTempAccdb, sourceTables[Array.IndexOf(targetTables, targetTableName)], targetTableName,
                            ODBCMgr.DSN_KEYS.OptimizerRuleDefinitionsDsnName, scenarioSqliteFile);
                    }

                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(scenarioSqliteFile)))
                    {
                        conn.Open();
                        // Delete any existing data from SQLite tables
                        foreach (string sourceTableName in sourceTables)
                        {
                            oDataMgr.m_strSQL = "DELETE FROM " + sourceTableName;
                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                        }
                    }

                    strCopyConn = oAdo.getMDBConnString(strTempAccdb, "", "");
                    using (var copyConn = new System.Data.OleDb.OleDbConnection(strCopyConn))
                    {
                        copyConn.Open();

                        foreach (string targetTableName in targetTables)
                        {
                            string sourceTableName = sourceTables[Array.IndexOf(targetTables, targetTableName)];
                            oAdo.m_strSQL = "INSERT INTO " + targetTableName +
                                " SELECT * FROM " + sourceTableName;
                            oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                        }
                    }

                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(scenarioSqliteFile)))
                    {
                        conn.Open();
                        oDataMgr.m_strSQL = "UPDATE scenario SET file = 'scenario_optimizer_rule_definitions.db'";
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    }

                // MIGRATE GIS DATA
                // Check if Processor parameters in SQLite
                string strTest = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile}";
                if (!System.IO.File.Exists(strTest))
                {
                    System.Windows.Forms.MessageBox.Show("Processor parameters have not been migrated to SQLite. SQLite GIS data cannot be loaded!", "FIA Biosum");
                    return;
                }
                // Check if Optimizer parameters in SQLite
                strTest = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\{Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile}";
                if (!System.IO.File.Exists(strTest))
                {
                    System.Windows.Forms.MessageBox.Show("Optimizer parameters have not been migrated to SQLite. SQLite GIS data cannot be loaded!", "FIA Biosum");
                    return;
                }

                string gisPathAndDbFile = this.ReferenceProjectDirectory + "\\" + Tables.TravelTime.DefaultTravelTimePathAndDbFile;
                if (!System.IO.File.Exists(gisPathAndDbFile))
                {
                    oDataMgr.CreateDbFile(gisPathAndDbFile);
                }
                // Create audit db
                string strAuditDBPath = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory}\{Tables.TravelTime.DefaultGisAuditPathAndDbFile}";
                if (!System.IO.File.Exists(strAuditDBPath))
                {
                    oDataMgr.CreateDbFile(strAuditDBPath);
                }

                // Create target tables in new database
                strCopyConn = oDataMgr.GetConnectionString(gisPathAndDbFile);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strCopyConn))
                {
                    conn.Open();
                    frmMain.g_oTables.m_oTravelTime.CreateSqliteProcessingSiteTable(oDataMgr, conn, Tables.TravelTime.DefaultProcessingSiteTableName);
                    frmMain.g_oTables.m_oTravelTime.CreateSqliteTravelTimeTable(oDataMgr, conn, Tables.TravelTime.DefaultTravelTimeTableName);
                }
                // Find path to existing tables
                oProjectDs = new Datasource();
                oProjectDs.m_strDataSourceMDBFile = this.ReferenceProjectDirectory + "\\db\\project.mdb";
                oProjectDs.m_strDataSourceTableName = "datasource";
                oProjectDs.m_strScenarioId = "";
                oProjectDs.LoadTableColumnNamesAndDataTypes = false;
                oProjectDs.LoadTableRecordCount = false;
                oProjectDs.populate_datasource_array();

                // Travel times
                int intTravelTable = oProjectDs.getTableNameRow(Datasource.TableTypes.TravelTimes);
                int intPSitesTable = oProjectDs.getTableNameRow(Datasource.TableTypes.ProcessingSites);
                string strDirectoryPath = oProjectDs.m_strDataSource[intTravelTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
                string strFileName = oProjectDs.m_strDataSource[intTravelTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
                //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
                string strTableName = oProjectDs.m_strDataSource[intTravelTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
                string strTableStatus = oProjectDs.m_strDataSource[intTravelTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strCopyConn))
                {
                    conn.Open();
                    string[] arrUpdateTableType = new string[2];
                    string[] arrUpdateTableName = new string[2];
                    if (oProjectDs.DataSourceTableExist(intTravelTable))
                    {
                        string strDbName = System.IO.Path.GetFileName(Tables.TravelTime.DefaultTravelTimePathAndDbFile);
                        string strNewDirectoryPath = System.IO.Path.GetDirectoryName(gisPathAndDbFile);
                        string strConn = oAdo.getMDBConnString(oProjectDs.getFullPathAndFile(Datasource.TableTypes.TravelTimes), "", "");
                        using (var pConn = new System.Data.OleDb.OleDbConnection(strConn))
                        {
                            pConn.Open();
                            oAdo.m_strSQL = $@"SELECT TRAVELTIME_ID, PSITE_ID, BIOSUM_PLOT_ID, COLLECTOR_ID,RAILHEAD_ID,
                                TRAVEL_MODE, ONE_WAY_HOURS, PLOT, STATECD FROM {strTableName}";
                            oAdo.CreateDataTable(pConn, oAdo.m_strSQL, strTableName, false);
                            using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(oAdo.m_strSQL, conn))
                            {
                                using (System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da))
                                {
                                    using (var transaction = conn.BeginTransaction())
                                    {
                                        da.InsertCommand = cb.GetInsertCommand();
                                        int rows = da.Update(oAdo.m_DataTable);
                                        transaction.Commit();
                                    }
                                }
                            }
                            oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.TravelTimes, strNewDirectoryPath, strDbName, strTableName);
                            arrUpdateTableType[0] = Datasource.TableTypes.TravelTimes;
                            arrUpdateTableName[0] = strTableName;
                            if (oProjectDs.DataSourceTableExist(intPSitesTable))
                            {
                                strTableName = oProjectDs.m_strDataSource[intPSitesTable, Datasource.TABLE].Trim();
                                oAdo.m_strSQL = $@"SELECT PSITE_ID,NAME,TRANCD,TRANCD_DEF,BIOCD,BIOCD_DEF,EXISTS_YN,LAT,LON,
                                            STATE,CITY,MILL_TYPE,COUNTY,STATUS,NOTES FROM {strTableName}";
                                oAdo.CreateDataTable(pConn, oAdo.m_strSQL, strTableName, false);
                                using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(oAdo.m_strSQL, conn))
                                {
                                    using (System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da))
                                    {
                                        using (var transaction = conn.BeginTransaction())
                                        {
                                            da.InsertCommand = cb.GetInsertCommand();
                                            int rows = da.Update(oAdo.m_DataTable);
                                            transaction.Commit();
                                        }
                                    }
                                }
                                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.ProcessingSites, strNewDirectoryPath, strDbName, strTableName);
                                arrUpdateTableType[1] = Datasource.TableTypes.ProcessingSites;
                                arrUpdateTableName[1] = strTableName;
                            }
                        }
                        strConn = oDataMgr.GetConnectionString($@"{this.ReferenceProjectDirectory}\{Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile}");
                        using (System.Data.SQLite.SQLiteConnection scenarioConn = new System.Data.SQLite.SQLiteConnection(strConn))
                        {
                            scenarioConn.Open();
                            for (int i = 0; i < arrUpdateTableType.Length; i++)
                            {
                                if (!string.IsNullOrEmpty(arrUpdateTableType[i]))
                                {
                                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                                    sb.Append($@"UPDATE {Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioDatasourceTableName} SET ");
                                    sb.Append($@"PATH = '{strNewDirectoryPath}', file='{strDbName}', table_name = '{arrUpdateTableName[i]}' ");
                                    sb.Append($@"WHERE TRIM(table_type) = '{arrUpdateTableType[i]}'");
                                    oDataMgr.SqlNonQuery(scenarioConn, sb.ToString());
                                }
                            }
                        }
                        strConn = oDataMgr.GetConnectionString($@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile}");
                        using (System.Data.SQLite.SQLiteConnection scenarioConn = new System.Data.SQLite.SQLiteConnection(strConn))
                        {
                            scenarioConn.Open();
                            for (int i = 0; i < arrUpdateTableType.Length; i++)
                            {
                                if (!string.IsNullOrEmpty(arrUpdateTableType[i]))
                                {
                                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                                    sb.Append($@"UPDATE {Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioDatasourceTableName} SET ");
                                    sb.Append($@"PATH = '{strNewDirectoryPath}', file='{strDbName}', table_name = '{arrUpdateTableName[i]}' ");
                                    sb.Append($@"WHERE TRIM(table_type) = '{arrUpdateTableType[i]}'");
                                    oDataMgr.SqlNonQuery(scenarioConn, sb.ToString());
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                // Clean-up
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName);
                }
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.FvsMasterDbDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.FvsMasterDbDsnName);
                }
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName);
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

        public void UpdateDatasources_5_11_1()
        {
            DataMgr oDataMgr = new DataMgr();
            Datasource oProjectDs = new Datasource();

            // Find path to existing tables
            oProjectDs.m_strDataSourceMDBFile = this.ReferenceProjectDirectory + "\\db\\project.mdb";
            oProjectDs.m_strDataSourceTableName = "datasource";
            oProjectDs.m_strScenarioId = "";
            oProjectDs.LoadTableColumnNamesAndDataTypes = false;
            oProjectDs.LoadTableRecordCount = false;
            oProjectDs.populate_datasource_array();

            // gis_travel_times.processing_site
            int intPSitesTable = oProjectDs.getTableNameRow(Datasource.TableTypes.ProcessingSites);
            string strDirectoryPath = oProjectDs.m_strDataSource[intPSitesTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = oProjectDs.m_strDataSource[intPSitesTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            string strTableName = oProjectDs.m_strDataSource[intPSitesTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            
            if (oProjectDs.DataSourceTableExist(intPSitesTable))
            {
                string strGisConn = oDataMgr.GetConnectionString(strDirectoryPath + "\\" + strFileName);
                using (System.Data.SQLite.SQLiteConnection gisConn = new System.Data.SQLite.SQLiteConnection(strGisConn))
                {
                    gisConn.Open();

                    if (!oDataMgr.FieldExist(gisConn, "SELECT * FROM " + strTableName, "PSITE_CN"))
                    {
                        string strSQL = "ALTER TABLE " + strTableName + " ADD COLUMN PSITE_CN CHAR(12)";
                        oDataMgr.SqlNonQuery(gisConn, strSQL);
                    }
                }
            }

            oDataMgr = null;
            oProjectDs = null;
        }

        public void UpdateDatasources_5_11_2()
        {
            DataMgr oDataMgr = new DataMgr();
            ODBCMgr odbcmgr = new ODBCMgr();
            dao_data_access oDao = new dao_data_access();
            ado_data_access oAdo = new ado_data_access();
            utils oUtils = new utils();
            //
            // Create master_aux.db and migrate values from master_aux.accdb
            //
            frmMain.g_sbpInfo.Text = "Version Update: Migrating DWM tables ...Stand by";

            string strSourceFile = ReferenceProjectDirectory.Trim() + "\\" + frmMain.g_oTables.m_oFIAPlot.DefaultDWMDbFile;
            string strDestFile = ReferenceProjectDirectory.Trim() + "\\" + frmMain.g_oTables.m_oFIAPlot.DefaultDWMSqliteDbFile;
            if (!System.IO.File.Exists(strDestFile))
            {
                oDataMgr.CreateDbFile(strDestFile);
            }
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
            {
                conn.Open();

                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteDWMCoarseWoodyDebrisTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName);
                }
                else
                {
                    oDataMgr.SqlNonQuery(conn, "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName);
                }
                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteDWMDuffLitterFuelTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName);
                }
                else
                {
                    oDataMgr.SqlNonQuery(conn, "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName);
                }
                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteDWMFineWoodyDebrisTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName);
                }
                else
                {
                    oDataMgr.SqlNonQuery(conn, "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName);
                }
                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteDWMTransectSegmentTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName);
                }
                else
                {
                    oDataMgr.SqlNonQuery(conn, "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName);
                }
            }

            // create DSN if needed
            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.MasterAuxDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.MasterAuxDsnName);
            }
            odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.MasterAuxDsnName, strSourceFile);

            // Set new temporary database
            string strTempAccdb = oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
            oDao.CreateMDB(strTempAccdb);

            //link access tables to temporary database
            oDao.CreateTableLinks(strTempAccdb, strSourceFile);

            //link sqlite tables to temporary database
            oDao.CreateSQLiteTableLink(strTempAccdb, frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName,
                frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName + "_1", ODBCMgr.DSN_KEYS.MasterAuxDsnName, strDestFile);
            oDao.CreateSQLiteTableLink(strTempAccdb, frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName,
                frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName + "_1", ODBCMgr.DSN_KEYS.MasterAuxDsnName, strDestFile);
            oDao.CreateSQLiteTableLink(strTempAccdb, frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName,
                frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName + "_1", ODBCMgr.DSN_KEYS.MasterAuxDsnName, strDestFile);
            oDao.CreateSQLiteTableLink(strTempAccdb, frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName,
                frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName + "_1", ODBCMgr.DSN_KEYS.MasterAuxDsnName, strDestFile);

            //insert data into sqlite tables
            string strCopyConn = oAdo.getMDBConnString(strTempAccdb, "", "");
            using (System.Data.OleDb.OleDbConnection copyConn = new System.Data.OleDb.OleDbConnection(strCopyConn))
            {
                copyConn.Open();

                oAdo.m_strSQL = "INSERT INTO " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName +
                    "_1 SELECT * FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName;
                oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);

                oAdo.m_strSQL = "INSERT INTO " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName +
                    "_1 SELECT * FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName;
                oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);

                oAdo.m_strSQL = "INSERT INTO " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName +
                    "_1 SELECT * FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName;
                oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);

                oAdo.m_strSQL = "INSERT INTO " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName +
                    "_1 SELECT * FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName;
                oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
            }

            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.MasterAuxDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.MasterAuxDsnName);
            }

            // Move sequence number tables from fvs_master.db to master.db
            frmMain.g_sbpInfo.Text = "Version Update: Move sequence number tables ...Stand by";
            strDestFile = ReferenceProjectDirectory.Trim() + "\\" + Tables.FVS.DefaultRxPackageDbFile;
            if (! System.IO.File.Exists(strDestFile))
            {
                oDataMgr.CreateDbFile(strDestFile);
            }
            // Create sequence number tables if they don't exist
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
            {
                conn.Open();
                if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultFVSPrePostSeqNumTable))
                {
                    frmMain.g_oTables.m_oFvs.CreateFVSOutputSQLitePrePostSeqNumTable(oDataMgr, conn, Tables.FVS.DefaultFVSPrePostSeqNumTable);
                }
                else
                {
                    oDataMgr.m_strSQL = $@"DELETE FROM {Tables.FVS.DefaultFVSPrePostSeqNumTable}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable))
                {
                    frmMain.g_oTables.m_oFvs.CreateFVSOutputPrePostSQLiteSeqNumRxPackageAssgnTable(oDataMgr, conn, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable);
                }
                else
                {
                    oDataMgr.m_strSQL = $@"DELETE FROM {Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultRxPackageTableName))
                {
                    frmMain.g_oTables.m_oFvs.CreateSQLiteRxPackageTable(oDataMgr, conn, Tables.FVS.DefaultRxPackageTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = $@"DELETE FROM {Tables.FVS.DefaultRxPackageTableName}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultRxTableName))
                {
                    frmMain.g_oTables.m_oFvs.CreateSQLiteRxTable(oDataMgr, conn, Tables.FVS.DefaultRxTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = $@"DELETE FROM {Tables.FVS.DefaultRxTableName}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultRxHarvestCostColumnsTableName))
                {
                    frmMain.g_oTables.m_oFvs.CreateSqliteRxHarvestCostColumnTable(oDataMgr, conn, Tables.FVS.DefaultRxHarvestCostColumnsTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = $@"DELETE FROM {Tables.FVS.DefaultRxHarvestCostColumnsTableName}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
            }

            Datasource oProjectDs = new Datasource();
            // Find path to existing tables
            oProjectDs.m_strDataSourceMDBFile = this.ReferenceProjectDirectory + "\\db\\project.mdb";
            oProjectDs.m_strDataSourceTableName = "datasource";
            oProjectDs.m_strScenarioId = "";
            oProjectDs.LoadTableColumnNamesAndDataTypes = false;
            oProjectDs.LoadTableRecordCount = false;
            oProjectDs.populate_datasource_array();
            // FVS PRE-POST SeqNum Definitions. Assuming that all the sequence number tables will be in the same db
            int intSeqNumTable = oProjectDs.getTableNameRow(Datasource.TableTypes.SeqNumDefinitions);
            // Again, assuming that the rx tables are all in the same database
            int intRxPackageTable = oProjectDs.getTableNameRow(Datasource.TableTypes.RxPackage);

            // create DSN if needed
            frmMain.g_sbpInfo.Text = "Version Update: Migrate tables from fvs_master.mdb ...Stand by";
            string fvsMasterDs = "MIGRATE_FVS_MASTER";
            if (odbcmgr.CurrentUserDSNKeyExist(fvsMasterDs))
            {
                odbcmgr.RemoveUserDSN(fvsMasterDs);
            }
            odbcmgr.CreateUserSQLiteDSN(fvsMasterDs, strDestFile);
            string[] arrTargetTables = { Tables.FVS.DefaultRxPackageTableName, Tables.FVS.DefaultRxTableName, Tables.FVS.DefaultRxHarvestCostColumnsTableName };
            strSourceFile = $@"{oProjectDs.m_strDataSource[intRxPackageTable, FIA_Biosum_Manager.Datasource.PATH].Trim()}\{oProjectDs.m_strDataSource[intRxPackageTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim()}";
            for (int i = 0; i < arrTargetTables.Length; i++)
            {
                oDao.CreateSQLiteTableLink(strSourceFile, arrTargetTables[i],
                    arrTargetTables[i] + "_1", fvsMasterDs, strDestFile);
            }
            System.Threading.Thread.Sleep(4000);

            string strDirectoryPath = oProjectDs.m_strDataSource[intSeqNumTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = oProjectDs.m_strDataSource[intSeqNumTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            string strTableName = oProjectDs.m_strDataSource[intSeqNumTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();            
            if (oProjectDs.DataSourceTableExist(intSeqNumTable))
            {
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
                {
                    conn.Open();
                    oDataMgr.m_strSQL = $@"ATTACH DATABASE '{strDirectoryPath}\{strFileName}' AS source";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    oDataMgr.m_strSQL = $@"INSERT INTO {Tables.FVS.DefaultFVSPrePostSeqNumTable} SELECT * FROM source.{strTableName}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    intSeqNumTable = oProjectDs.getTableNameRow(Datasource.TableTypes.SeqNumRxPackageAssign);
                    strTableName = oProjectDs.m_strDataSource[intSeqNumTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
                    oDataMgr.m_strSQL = $@"INSERT INTO {Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable} SELECT * FROM source.{strTableName}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }

                // Update entries in project data sources table
                string strMasterDb = System.IO.Path.GetFileName(Tables.FVS.DefaultRxPackageDbFile);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.SeqNumDefinitions, $@"{ReferenceProjectDirectory.Trim()}\db", strMasterDb, Tables.FVS.DefaultFVSPrePostSeqNumTable);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.SeqNumRxPackageAssign, $@"{ReferenceProjectDirectory.Trim()}\db", strMasterDb, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable);
            }

            strDirectoryPath = oProjectDs.m_strDataSource[intRxPackageTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strFileName = oProjectDs.m_strDataSource[intRxPackageTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            strTableName = oProjectDs.m_strDataSource[intRxPackageTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();

            if (oProjectDs.DataSourceTableExist(intRxPackageTable))
            {
                strCopyConn = oAdo.getMDBConnString(strSourceFile, "", "");
                using (OleDbConnection copyConn = new System.Data.OleDb.OleDbConnection(strCopyConn))
                {
                    copyConn.Open();
                    oAdo.m_strSQL = $@"INSERT INTO {arrTargetTables[0]}_1 SELECT RXPACKAGE, DESCRIPTION, rxcycle_length,
                        simyear1_rx, simyear1_fvscycle, simyear2_rx, simyear2_fvscycle, simyear3_rx, simyear3_fvscycle,
                        simyear4_rx, simyear4_fvscycle FROM {strTableName}";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = $@"DROP TABLE {arrTargetTables[0]}_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = $@"INSERT INTO {arrTargetTables[1]}_1 SELECT RX, DESCRIPTION, HarvestMethodLowSlope, HarvestMethodSteepSlope
                        FROM {arrTargetTables[1]}";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = $@"DROP TABLE {arrTargetTables[1]}_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "INSERT INTO " + arrTargetTables[2] + "_1" +
                        " SELECT * FROM " + arrTargetTables[2];
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = $@"DROP TABLE {arrTargetTables[2]}_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                }

                string strMasterDb = System.IO.Path.GetFileName(Tables.FVS.DefaultRxPackageDbFile);
                // Update project data sources
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.RxPackage, $@"{ReferenceProjectDirectory.Trim()}\db", strMasterDb, Tables.FVS.DefaultRxPackageTableName);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.Rx, $@"{ ReferenceProjectDirectory.Trim()}\db", strMasterDb, Tables.FVS.DefaultRxTableName);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.RxHarvestCostColumns, $@"{ ReferenceProjectDirectory.Trim()}\db", strMasterDb, Tables.FVS.DefaultRxHarvestCostColumnsTableName);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.HarvestMethods, "@@appdata@@\\fiabiosum", Tables.Reference.DefaultBiosumReferenceSqliteFile, Tables.Reference.DefaultHarvestMethodsTableName);

                // Update Optimizer data sources
                strDestFile = $@"{ReferenceProjectDirectory.Trim()}\{Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile}";
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
                {
                    conn.Open();
                    oDataMgr.m_strSQL = $@"UPDATE {Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioDatasourceTableName} SET file = '{strMasterDb}' 
                        where table_type in ('Treatment Prescriptions','Treatment Prescriptions Harvest Cost Columns','{Datasource.TableTypes.RxPackage}')";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    oDataMgr.m_strSQL = $@"UPDATE {Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioDatasourceTableName} SET path = '@@appdata@@\fiabiosum', file = '{Tables.Reference.DefaultBiosumReferenceSqliteFile}' 
                        where table_type = '{Datasource.TableTypes.HarvestMethods}' ";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                // Update Processor data sources
                strDestFile = $@"{ReferenceProjectDirectory.Trim()}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile}";
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
                {
                    conn.Open();
                    oDataMgr.m_strSQL = $@"UPDATE {Tables.Scenario.DefaultScenarioDatasourceTableName} SET file = '{strMasterDb}' 
                        where table_type in ('Treatment Prescriptions','Treatment Prescriptions Harvest Cost Columns','{Datasource.TableTypes.RxPackage}')";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    oDataMgr.m_strSQL = $@"UPDATE {Tables.Scenario.DefaultScenarioDatasourceTableName} SET path = '@@appdata@@\fiabiosum', file = '{Tables.Reference.DefaultBiosumReferenceSqliteFile}' 
                        where table_type = '{Datasource.TableTypes.HarvestMethods}' ";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                // Remove obsolete data source definitions
                using (OleDbConnection deleteConn = new System.Data.OleDb.OleDbConnection(oAdo.getMDBConnString(oProjectDs.m_strDataSourceMDBFile, "", "")))
                {
                    deleteConn.Open();
                    oAdo.m_strSQL = $@"DELETE FROM {oProjectDs.m_strDataSourceTableName} WHERE TABLE_TYPE IN 
                        ('Treatment Prescriptions Assigned FVS Commands', 'Treatment Prescription Categories', 'Treatment Prescription Subcategories',
                         'Treatment Package Assigned FVS Commands', 'Treatment Package FVS Commands Order', 'Treatment Package Members', 
                         'FVS Western Tree Species Translator', 'FVS Eastern Tree Species Translator')";
                    oAdo.SqlNonQuery(deleteConn, oAdo.m_strSQL);
                }
            }
            if (odbcmgr.CurrentUserDSNKeyExist(fvsMasterDs))
            {
                odbcmgr.RemoveUserDSN(fvsMasterDs);
            }
            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }

            // Migrate any existing variant specific FVSIn.db files to one combined FVSIn.db
            frmMain.g_sbpInfo.Text = "Version Update: Migrate FVSIn.db files from [variant]/FVSIn.db to combined FVSIn.db ...Stand by";
            string strFVSInFile = ReferenceProjectDirectory.Trim() + "\\fvs\\data\\" + Tables.FIA2FVS.DefaultFvsInputFile;

            // Get list of FVS Variants from Plot table
            List<String> lstVar = new List<String>();
            int intPlotTable = oProjectDs.getTableNameRow(Datasource.TableTypes.Plot);
            strDirectoryPath = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strFileName = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            strTableName = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strVariantCon = oAdo.getMDBConnString(strDirectoryPath + "\\" + strFileName, "", "");
            using (OleDbConnection variantCon = new OleDbConnection(strVariantCon))
            {
                variantCon.Open();

                oAdo.m_strSQL = Queries.FVS.GetFVSVariantSQL_access(strTableName);
                oAdo.SqlQueryReader(variantCon, oAdo.m_strSQL);

                if (oAdo.m_OleDbDataReader.HasRows)
                {
                    while (oAdo.m_OleDbDataReader.Read())
                    {
                        string strVariant = oAdo.m_OleDbDataReader["fvs_variant"].ToString().Trim();
                        lstVar.Add(strVariant);
                    }
                }
            }

            // Cycle through list. If variant FVSIn exists, add it to the combined FVSIn.db. Create the combined FVSIn.db if needed
            string strFVSConn = oDataMgr.GetConnectionString(strFVSInFile);
            foreach (string variant in lstVar)
            {
                string strVarFvsIn = ReferenceProjectDirectory.Trim() + "\\fvs\\data\\" + variant + "\\" + Tables.FIA2FVS.DefaultFvsInputFile;
                if (System.IO.File.Exists(strVarFvsIn))
                {
                    if (!System.IO.File.Exists(strFVSInFile))
                    {
                        System.IO.File.Copy(strVarFvsIn, strFVSInFile, true);

                        using (System.Data.SQLite.SQLiteConnection fvsConn = new System.Data.SQLite.SQLiteConnection(strFVSConn))
                        {
                            fvsConn.Open();

                            if (!oDataMgr.ColumnExist(fvsConn, Tables.FIA2FVS.DefaultFvsInputTreeTableName, "VARIANT"))
                            {
                                oDataMgr.AddColumn(fvsConn, Tables.FIA2FVS.DefaultFvsInputTreeTableName, "VARIANT", "CHAR", "2");
                            }

                            oDataMgr.m_strSQL = "UPDATE " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                                " SET VARIANT = '" + variant + "' WHERE VARIANT IS NULL";
                            oDataMgr.SqlNonQuery(fvsConn, oDataMgr.m_strSQL);
                        }
                    }
                    else
                    {
                        using (System.Data.SQLite.SQLiteConnection fvsConn = new System.Data.SQLite.SQLiteConnection(strFVSConn))
                        {
                            fvsConn.Open();

                            oDataMgr.m_strSQL = "ATTACH DATABASE '" + strVarFvsIn + "' AS variant";
                            oDataMgr.SqlNonQuery(fvsConn, oDataMgr.m_strSQL);

                            string strStandFields = oDataMgr.getFieldNames(fvsConn, "SELECT * FROM variant." + Tables.FIA2FVS.DefaultFvsInputStandTableName);
                            oDataMgr.m_strSQL = "INSERT INTO " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                                " (" + strStandFields + ") SELECT " + strStandFields +
                                " FROM variant." + Tables.FIA2FVS.DefaultFvsInputStandTableName;
                            oDataMgr.SqlNonQuery(fvsConn, oDataMgr.m_strSQL);

                            string strTreeFields = oDataMgr.getFieldNames(fvsConn, "SELECT * FROM variant." + Tables.FIA2FVS.DefaultFvsInputTreeTableName);
                            oDataMgr.m_strSQL = "INSERT INTO " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                                " (" + strTreeFields + ") SELECT " + strTreeFields +
                                " FROM variant." + Tables.FIA2FVS.DefaultFvsInputTreeTableName;
                            oDataMgr.SqlNonQuery(fvsConn, oDataMgr.m_strSQL);

                            if (!oDataMgr.ColumnExist(fvsConn, Tables.FIA2FVS.DefaultFvsInputTreeTableName, "VARIANT"))
                            {
                                oDataMgr.AddColumn(fvsConn, Tables.FIA2FVS.DefaultFvsInputTreeTableName, "VARIANT", "CHAR", "2");
                            }

                            oDataMgr.m_strSQL = "UPDATE " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                                " SET VARIANT = '" + variant + "' WHERE VARIANT IS NULL";
                            oDataMgr.SqlNonQuery(fvsConn, oDataMgr.m_strSQL);
                        }
                    }
                }
            }

            // update optimizer calculated variables db to add null threshold table
            // and use negative column to variables table
            frmMain.g_sbpInfo.Text = "Version Update: Updated Optimizer Calculated Variables ...Stand by";
            string strCalculatedVariablesDb = ReferenceProjectDirectory.Trim() + "\\" + Tables.OptimizerDefinitions.DefaultSqliteDbFile;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strCalculatedVariablesDb)))
            {
                conn.Open();
                if (!oDataMgr.ColumnExist(conn, Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName, "NEGATIVES_YN"))
                {
                    oDataMgr.AddColumn(conn, Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName, "NEGATIVES_YN", "CHAR", "1");

                    oDataMgr.m_strSQL = "UPDATE " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                        " SET NEGATIVES_YN = 'N'";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, Tables.OptimizerDefinitions.DefaultOptimizerProjectConfigTableName))
                {
                    oDataMgr.m_strSQL = "CREATE TABLE " + Tables.OptimizerDefinitions.DefaultOptimizerProjectConfigTableName + " (fvs_null_threshold INTEGER)";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                    oDataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerDefinitions.DefaultOptimizerProjectConfigTableName + " VALUES (4)";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
            }
        }

        public void UpdateDatasources_5_12_0()
        {
            DataMgr oDataMgr = new DataMgr();
            ODBCMgr odbcmgr = new ODBCMgr();
            dao_data_access oDao = new dao_data_access();
            ado_data_access oAdo = new ado_data_access();
            utils oUtils = new utils();

            // Migrate plot, cond, tree, and sitetree tables from master.mdb to master.db
            frmMain.g_sbpInfo.Text = "Version Update: Moving plot, cond, tree, and sitetree tables ...Stand by";

            string strDestFile = ReferenceProjectDirectory.Trim() + "\\" + frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableSqliteDbFile;
            if (!System.IO.File.Exists(strDestFile))
            {
                oDataMgr.CreateDbFile(strDestFile);
            }

            // Create tables if they don't exist
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
            {
                conn.Open();

                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqlitePlotTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName;
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteConditionTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName;
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteTreeTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName;
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteSiteTreeTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName;
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
            }

            Datasource oProjectDs = new Datasource();
            // Find path to existing tables
            oProjectDs.m_strDataSourceMDBFile = this.ReferenceProjectDirectory + "\\db\\project.mdb";
            oProjectDs.m_strDataSourceTableName = "datasource";
            oProjectDs.m_strScenarioId = "";
            oProjectDs.LoadTableColumnNamesAndDataTypes = false;
            oProjectDs.LoadTableRecordCount = false;
            oProjectDs.populate_datasource_array();
            // plot
            int intPlotTable = oProjectDs.getTableNameRow(Datasource.TableTypes.Plot);
            // cond
            int intCondTable = oProjectDs.getTableNameRow(Datasource.TableTypes.Condition);
            // tree
            int intTreeTable = oProjectDs.getTableNameRow(Datasource.TableTypes.Tree);
            // sitetree
            int intSiteTreeTable = oProjectDs.getTableNameRow(Datasource.TableTypes.SiteTree);

            // Create DSN if needed
            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.MasterDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.MasterDsnName);
            }
            odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.MasterDsnName, strDestFile);

            string[] arrTargetTables = { frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName, frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName, 
                frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName, frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName };
            string strSourceFile = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.PATH].Trim() + "\\" + oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            string[] strSourceTables = { oProjectDs.m_strDataSource[intPlotTable, Datasource.TABLE].Trim(), oProjectDs.m_strDataSource[intCondTable, Datasource.TABLE].Trim(),
                oProjectDs.m_strDataSource[intTreeTable, Datasource.TABLE].Trim(), oProjectDs.m_strDataSource[intSiteTreeTable, Datasource.TABLE].Trim() };
            string strMasterDb = System.IO.Path.GetFileName(frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableSqliteDbFile);

            // Set new temporary database
            string strTempAccdb = oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
            oDao.CreateMDB(strTempAccdb);

            //link access tables to temporary database
            foreach (string table in strSourceTables)
            {
                oDao.CreateTableLink(strTempAccdb, table, strSourceFile, table);
            }

            for (int i = 0; i < arrTargetTables.Length; i++)
            {
                oDao.CreateSQLiteTableLink(strTempAccdb, arrTargetTables[i],
                    arrTargetTables[i] + "_1", ODBCMgr.DSN_KEYS.MasterDsnName, strDestFile);
            }
            System.Threading.Thread.Sleep(4000);

            if (oProjectDs.DataSourceTableExist(intPlotTable))
            {
                string strCopyConn = oAdo.getMDBConnString(strTempAccdb, "", "");
                using (OleDbConnection copyConn = new System.Data.OleDb.OleDbConnection(strCopyConn))
                {
                    copyConn.Open();
                    oAdo.m_strSQL = "INSERT INTO " + arrTargetTables[0] + "_1 " +
                        "SELECT biosum_plot_id, statecd, invyr, unitcd, countycd, " + arrTargetTables[0] + ", " +
                        "measyear, elev, fvs_variant, fvsloccode AS fvs_loc_cd, half_state, gis_yard_dist_ft, " +
                        "num_cond, one_cond_yn, lat, lon, macro_breakpoint_dia, precipitation, " +
                        "ecosubcd, biosum_status_cd, cn FROM " + arrTargetTables[0];
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "INSERT INTO " + arrTargetTables[1] + "_1 " +
                        "SELECT biosum_cond_id, biosum_plot_id, invyr, condid, condprop, cond_status_cd, " +
                        "fortypcd, owncd, owngrpcd, reservcd, siteclcd, sibase, sicond, sisp, " +
                        "slope, aspect, stdage, stdszcd, habtypcd1, adforcd, qmd_all_inch, " +
                        "qmd_hwd_inch, qmd_swd_inch, acres, unitcd, vol_loc_grp, tpacurr, " +
                        "hwd_tpacurr, swd_tpacurr, ba_ft2_ac, hwd_ba_ft2_ac, swd_ba_ft2_ac, " +
                        "vol_ac_grs_ft3, hwd_vol_ac_grs_ft3, swd_vol_ac_grs_ft3, " +
                        "volcsgrs, hwd_volcsgrs, swd_volcsgrs, gsstkcd, alstkcd, " +
                        "condprop_unadj, micrprop_unadj, subpprop_unadj, macrprop_unadj, " +
                        "cn, biosum_status_cd, dwm_fuelbed_typcd, balive, stdorgcd FROM " + arrTargetTables[1];
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "INSERT INTO " + arrTargetTables[2] + "_1 " +
                        "SELECT biosum_cond_id, invyr, statecd, unitcd, countycd, subp, " +
                        "tree, condid, statuscd, spcd, spgrpcd, dia, diahtcd, ht, htcd, actualht, " +
                        "formcl, treeclcd, cr, cclcd, cull, roughcull, decaycd, stocking, tpacurr, wdldstem, " +
                        "volcfnet, volcfgrs, volcsnet, volcsgrs, volbfnet, volbfgrs, voltsgrs, " +
                        "drybiot, drybiom, bhage, cullbf, cullcf, totage, mist_cl_cd, agentcd, " +
                        "damtyp1, damsev1, damtyp2, damsev2, tpa_unadj, condprop_specific, sitree, " +
                        "upper_dia, upper_dia_ht, centroid_dia, centroid_dia_ht_actual, sawht, htdmp, " +
                        "boleht, cull_fld, culldead, cullform, cullmstop, cfsnd, bfsnd, standing_dead_cd, " +
                        "volcfsnd, drybio_bole, drybio_top, drybio_sapling, drybio_wdld_spp, " +
                        "fvs_tree_id, cn, biosum_status_cd FROM " + arrTargetTables[2];
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "INSERT INTO " + arrTargetTables[3] + "_1 " +
                        "SELECT biosum_plot_id, invyr, condid, tree, spcd, dia, ht, agedia, " +
                        "spgrpcd, sitree, sibase, subp, method, validcd, condlist, biosum_status_cd FROM " + arrTargetTables[3];
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "DROP TABLE " + arrTargetTables[0] + "_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "DROP TABLE " + arrTargetTables[1] + "_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "DROP TABLE " + arrTargetTables[2] + "_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "DROP TABLE " + arrTargetTables[3] + "_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                }

                // Update project datasources
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.Plot, ReferenceProjectDirectory + "\\db", strMasterDb, arrTargetTables[0]);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.Condition, ReferenceProjectDirectory + "\\db", strMasterDb, arrTargetTables[1]);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.Tree, ReferenceProjectDirectory + "\\db", strMasterDb, arrTargetTables[2]);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.SiteTree, ReferenceProjectDirectory + "\\db", strMasterDb, arrTargetTables[3]);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.PopStratumAdjFactors, ReferenceProjectDirectory + "\\db", strMasterDb, 
                    frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName);

                // Update processor datasources
                strDestFile = ReferenceProjectDirectory.Trim() + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
                {
                    conn.Open();

                    oDataMgr.m_strSQL = "UPDATE " + Tables.Scenario.DefaultScenarioDatasourceTableName +
                        " SET file = '" + strMasterDb + "' WHERE table_type IN ('" +
                        Datasource.TableTypes.Plot + "', '" +
                        Datasource.TableTypes.Condition + "', '" +
                        Datasource.TableTypes.Tree + "')";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }

                // Update optimizer datasources
                strDestFile = ReferenceProjectDirectory.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
                {
                    conn.Open();

                    oDataMgr.m_strSQL = "UPDATE " + Tables.Scenario.DefaultScenarioDatasourceTableName +
                        " SET file = '" + strMasterDb + "' WHERE table_type IN ('" +
                        Datasource.TableTypes.Plot + "', '" +
                        Datasource.TableTypes.Condition + "')";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
            }

            // Update project datasources; tree_species, fvs_tree_species, and fiadb_fvs_variant have been eliminated; fia_tree_species_ref has moved
            string strDsConn = oAdo.getMDBConnString(oProjectDs.m_strDataSourceMDBFile, "", "");
            using (OleDbConnection copyConn = new System.Data.OleDb.OleDbConnection(strDsConn))
            {
                copyConn.Open();
                oAdo.m_strSQL = $@"DELETE FROM {oProjectDs.m_strDataSourceTableName} WHERE TABLE_TYPE IN ('{Datasource.TableTypes.TreeSpecies}','{Datasource.TableTypes.FvsTreeSpecies}', '{Datasource.TableTypes.FVSVariant}',
                    '{Datasource.TableTypes.OwnerGroups}')";
                oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
            }

            // Update project datasources
            oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.FiaTreeMacroPlotBreakpointDia, "@@appdata@@\\fiabiosum",
                Tables.Reference.DefaultTreeMacroPlotBreakPointDiaTableDbFile, Tables.Reference.DefaultTreeMacroPlotBreakPointDiaTableName);
            oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.FiaTreeSpeciesReference, "@@appdata@@\\fiabiosum",
                Tables.Reference.DefaultBiosumReferenceSqliteFile, Tables.Reference.DefaultFIATreeSpeciesTableName);

            // Update processor datasource
            strDestFile = ReferenceProjectDirectory.Trim() + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
            {
                conn.Open();
                oDataMgr.m_strSQL = "DELETE FROM " + Tables.Scenario.DefaultScenarioDatasourceTableName +
                    " WHERE table_type in ('" + Datasource.TableTypes.TreeSpecies + "','" + Datasource.TableTypes.ProcessingSites + "')";
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                oDataMgr.m_strSQL = $@"UPDATE {Tables.Scenario.DefaultScenarioDatasourceTableName} SET FILE = '{Tables.Reference.DefaultBiosumReferenceSqliteFile}' 
                    WHERE table_type = '{Datasource.TableTypes.FiaTreeSpeciesReference}'";
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
            }

            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.MasterDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.MasterDsnName);
            }

            // Add BioSum generated site index flag column to FVS_STANDINIT_COND table in FVSIn.db
            frmMain.g_sbpInfo.Text = "Version Update: Updating FVSIn.db tables ...Stand by";
            string strFVSInFile = ReferenceProjectDirectory.Trim() + "\\fvs\\data\\" + Tables.FIA2FVS.DefaultFvsInputFile;
            if (System.IO.File.Exists(strFVSInFile))
            {
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strFVSInFile)))
                {
                    conn.Open();

                    if (!oDataMgr.ColumnExist(conn, Tables.FIA2FVS.DefaultFvsInputStandTableName, "SITE_INDEX_BSCALC_YN"))
                    {
                        oDataMgr.AddColumn(conn, Tables.FIA2FVS.DefaultFvsInputStandTableName, "SITE_INDEX_BSCALC_YN", "CHAR", "1");
                    }
                }
            }
            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }

            // Copy ref_master.db to project if it doesn't already exist
            if (!System.IO.File.Exists(ReferenceProjectDirectory.Trim() + "\\" + Tables.Reference.DefaultRefMasterDbFile))
            {
                System.IO.File.Copy($@"{frmMain.g_oEnv.strAppDir}\{Tables.Reference.DefaultRefMasterDbFile}", ReferenceProjectDirectory.Trim() + "\\" + Tables.Reference.DefaultRefMasterDbFile, true );
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

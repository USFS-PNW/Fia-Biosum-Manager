using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Threading;

namespace FIA_Biosum_Manager
{
  public class frmFCSTreeVolumeEdit : Form
  {
    const int COL_VOLCFGRS = 0;
    const int COL_VOLCSGRS = 1;
    const int COL_VOLCFNET = 2;
    const int COL_DRYBIOM = 3;
    const int COL_DRYBIOT = 4;
    const int COL_VOLTSGRS = 5;
    const int COL_ID = 6;
    const int COL_STATE = 7;
    const int COL_COUNTY = 8;
    const int COL_PLOT = 9;
    const int COL_FVSVARIANT =10;
    const int COL_INVYR = 11;
    const int COL_TREEID = 12;
    const int COL_VOLLOCGRP = 15;
    const int COL_SPCD = 12;
    const int COL_DBH = 13;
    const int COL_HT = 14;
    const int COL_ACTUALHT = 16;
    const int COL_CR = 19;
    const int COL_STATUSCD = 17;
    const int COL_TREECLCD = 18;
    const int COL_CULL = 20;
    const int COL_ROUGHCULL = 21;
    const int COL_DECAYCD = 22;
    const int COL_TOTAGE = 23;

    /*
    const int COL_SITREE = 24;
    const int COL_WDLDSTEM = 25;
    const int COL_UPPERDIA = 26;
    const int COL_UPPERDIAHT = 27;
    const int COL_CENTROIDDIA = 28;
     */

    const int COL_TRECN = 24;
    const int COL_PLTCN = 25;
    const int COL_CNDCN = 26;
    
    private System.Windows.Forms.Button btnTreeVolBatch;
    private System.Windows.Forms.Panel panel1;
    private System.Windows.Forms.GroupBox groupBox1;
    private System.Windows.Forms.GroupBox groupBox2;
    private System.Windows.Forms.Label label7;
    private System.Windows.Forms.TextBox txtDbh;
    private System.Windows.Forms.Label label6;
    private System.Windows.Forms.TextBox txtSpCd;
    private System.Windows.Forms.Label label5;
    private System.Windows.Forms.TextBox txtInvYr;
    private System.Windows.Forms.Label label4;
    private System.Windows.Forms.TextBox txtVolLocGrp;
    private System.Windows.Forms.Label label3;
    private System.Windows.Forms.TextBox txtStateCd;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.TextBox txtPlot;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.TextBox txtCountyCd;
    private System.Windows.Forms.Button btnTreeVolSingle;
    private System.Windows.Forms.Label label9;
    private System.Windows.Forms.TextBox txtActualHt;
    private System.Windows.Forms.Label label8;
    private System.Windows.Forms.TextBox txtHt;
    private System.Windows.Forms.Label label10;
    private System.Windows.Forms.TextBox txtCR;
    private System.Windows.Forms.Label label12;
    private System.Windows.Forms.TextBox txtTreeClCd;
    private System.Windows.Forms.Label label11;
    private System.Windows.Forms.TextBox txtStatusCd;
    private System.Windows.Forms.GroupBox groupBox3;
    private System.Windows.Forms.Label lblDRYBIOM;
    private System.Windows.Forms.Label label20;
    private System.Windows.Forms.Label lblDRYBIOT;
    private System.Windows.Forms.Label label19;
    private System.Windows.Forms.Label lblVOLCFNET;
    private System.Windows.Forms.Label label18;
    private System.Windows.Forms.Label lblVOLCSGRS;
    private System.Windows.Forms.Label label17;
    private System.Windows.Forms.Label lblVOLCFGRS;
    private System.Windows.Forms.Label label15;
    private System.Windows.Forms.Label label14;
    private System.Windows.Forms.TextBox txtRoughCull;
    private System.Windows.Forms.Label label13;
    private System.Windows.Forms.TextBox txtCull;
    private System.Windows.Forms.Button btnDefaultSingle;
    private System.Windows.Forms.Button btnLinkTableTest;
    //
    //GRID POP UP MENU
    //
    private const int MENU_FILTERBYVALUE = 0;
    private const int MENU_FILTERBYENTEREDVALUE = 1;
    private const int MENU_REMOVEFILTER = 2;
    private const int MENU_UNIQUEVALUES = 3;
    private const int MENU_MODIFY = 5;
    private const int MENU_DELETE = 6;
    private const int MENU_SELECTALL = 8;
    private const int MENU_IDXASC = 10;
    private const int MENU_IDXDESC = 11;
    private const int MENU_REMOVEIDX = 12;
    private const int MENU_MAX = 14;
    private const int MENU_MIN = 15;
    private const int MENU_AVG = 16;
    private const int MENU_SUM = 17;
    private const int MENU_COUNTBYVALUE = 18;
    private Button btnEdit;
    private Button btnLoad;
    private ComboBox cmbDatasource;
    private uc_gridview uc_gridview1 = new uc_gridview();
    private Button btnCancel;
    private Label lblPerc = new Label();
    private RxTools m_oRxTools = new RxTools();
    private Queries m_oQueries = new Queries();
    string[] m_strFVSTreeTableLinkNameArray = null;
    string m_strTempDBFile = "";
    string m_strGridTableSource = "";
    static string m_strOldPerc = "0";
    private ProgressBarBasic.ProgressBarBasic progressBarBasic1;
    private Label lblVOLTSGRS;
    private Label label21;
    private Label label16;
    private Label lblFSNetwork;

    

    private ado_data_access m_oAdo = new ado_data_access();
    private int m_intError = 0;
    private string m_strError;

   

    public frmFCSTreeVolumeEdit()
    {

        if (frmMain.Validate_OracleConnectivity() < 0)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Abort;
            this.Close();
            return;
        }

        
     
     this.SetStyle(ControlStyles.OptimizedDoubleBuffer |
               ControlStyles.AllPaintingInWmPaint,
               true);
      int x,y;
      string strFVSTreeTableLinkNameList="";
      string strTableName="";

      dao_data_access oDao = new dao_data_access();
      InitializeComponent();
      groupBox1.Controls.Add(uc_gridview1);
      uc_gridview1.Top = 15; uc_gridview1.Left = 5;
      uc_gridview1.Width = groupBox1.Width - 10;
      uc_gridview1.Height = cmbDatasource.Top - uc_gridview1.Top - 5;
     
           
      uc_gridview1.CloseButton_Visible = false;
      uc_gridview1.Show();
      ResizeForm();
      
      if (frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim().Length > 0)
      {

          m_oQueries.m_oFvs.LoadDatasource = true;
          m_oQueries.m_oFIAPlot.LoadDatasource = true;
          m_oQueries.LoadDatasources(true);
          m_strTempDBFile = m_oQueries.m_strTempDbFile;
          oDao.CreateTableLink(m_strTempDBFile, "treesample", frmMain.g_oEnv.strAppDir + "\\db\\treesample.accdb", "treesample");
          oDao.m_DaoWorkspace.Close();
          //
          //CREATE LINK IN TEMP MDB TO ALL VARIANT FVS_CUTLIST TABLES
          //
          m_oRxTools.CreateTableLinksToFVSOutTreeListTables(m_oQueries, m_oQueries.m_strTempDbFile);

          //
          //CREATE LINK IN TEMP MDB TO ALL VARIANT & PACKAGE FVS_SUMMARY TABLES
          //
          m_oRxTools.CreateTableLinksToFVSOutSummaryTables(m_oQueries, m_oQueries.m_strTempDbFile);

          //
          //OPEN CONNECTION TO TEMP DB FILE
          //
          m_oAdo = new ado_data_access();
          m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));

          if (m_oAdo.m_intError == 0)
          {
              RxPackageItem_Collection oRxPackageItemCollection = new RxPackageItem_Collection();
              m_oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(m_oAdo, m_oAdo.m_OleDbConnection, m_oQueries, oRxPackageItemCollection);
              //
              //GET LIST OF VARIANTS
              //
              string strVariantsList = m_oRxTools.GetListOfFVSVariantsInPlotTable(m_oAdo, m_oAdo.m_OleDbConnection, m_oQueries.m_oFIAPlot.m_strPlotTable);
              string[] strVariantsArray = frmMain.g_oUtils.ConvertListToArray(strVariantsList, ",");
              if (strVariantsArray != null && strVariantsArray[0] != null)
              {
                  //find the variants that have tree cut list tables
                  for (x = 0; x <= strVariantsArray.Length - 1; x++)
                  {
                      for (y = 0; y <= oRxPackageItemCollection.Count - 1; y++)
                      {
                          strTableName = "fvs_tree_IN_" + strVariantsArray[x].Trim() + "_P" + oRxPackageItemCollection.Item(y).RxPackageId + "_TREE_CUTLIST";
                          if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, strTableName))
                          {
                              strFVSTreeTableLinkNameList = strFVSTreeTableLinkNameList + strTableName + ",";
                          }
                      }
                  }
              }
              cmbDatasource.Items.Clear();
              cmbDatasource.Items.Add("Tree Sample");
              cmbDatasource.Items.Add("Tree Table");
              //load the list into an array
              if (strFVSTreeTableLinkNameList.Trim().Length > 0)
              {
                  strFVSTreeTableLinkNameList = strFVSTreeTableLinkNameList.Substring(0, strFVSTreeTableLinkNameList.Length - 1);
                  m_strFVSTreeTableLinkNameArray = frmMain.g_oUtils.ConvertListToArray(strFVSTreeTableLinkNameList, ",");
                  for (x = 0; x <= m_strFVSTreeTableLinkNameArray.Length - 1; x++)
                  {
                      cmbDatasource.Items.Add(m_strFVSTreeTableLinkNameArray[x].Trim());
                  }
              }

              
             
              oRxPackageItemCollection.Clear();
              oRxPackageItemCollection = null;
          }
      }
      else
      {
          m_strTempDBFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
          oDao.CreateMDB(m_strTempDBFile);
          oDao.CreateTableLink(m_strTempDBFile, "treesample", frmMain.g_oEnv.strAppDir + "\\db\\treesample.accdb", "treesample");
          oDao.m_DaoWorkspace.Close();
      }
      oDao = null;
      

      
   
      LoadDefaultSingleRecordValues();

      lblFSNetwork.Text = FIA_Biosum_Manager.utils.FS_NETWORK == utils.FS_NETWORK_STATUS.NotAvailable ? "ORACLE XE BIOSUM_VOLUME" : "FS NETWORK FCS_TREE.DB BIOSUM_CALC";
    }
     /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose( bool disposing )
    {
      if( disposing && ( components != null ) )
      {
        components.Dispose();
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
            this.btnTreeVolBatch = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.progressBarBasic1 = new ProgressBarBasic.ProgressBarBasic();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnDefaultSingle = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.lblVOLTSGRS = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.lblDRYBIOM = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.lblDRYBIOT = new System.Windows.Forms.Label();
            this.lblVOLCFNET = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.lblVOLCSGRS = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.lblVOLCFGRS = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.txtRoughCull = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txtCull = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.txtTreeClCd = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.txtStatusCd = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.txtCR = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtActualHt = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtHt = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtDbh = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtSpCd = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtInvYr = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtVolLocGrp = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtStateCd = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPlot = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtCountyCd = new System.Windows.Forms.TextBox();
            this.btnTreeVolSingle = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblFSNetwork = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnLoad = new System.Windows.Forms.Button();
            this.cmbDatasource = new System.Windows.Forms.ComboBox();
            this.btnLinkTableTest = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnTreeVolBatch
            // 
            this.btnTreeVolBatch.Location = new System.Drawing.Point(387, 293);
            this.btnTreeVolBatch.Name = "btnTreeVolBatch";
            this.btnTreeVolBatch.Size = new System.Drawing.Size(208, 21);
            this.btnTreeVolBatch.TabIndex = 2;
            this.btnTreeVolBatch.Text = "Batch Calculate Volume And Biomass";
            this.btnTreeVolBatch.UseVisualStyleBackColor = true;
            this.btnTreeVolBatch.Click += new System.EventHandler(this.btnTreeVolBatch_Click);
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.progressBarBasic1);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(856, 467);
            this.panel1.TabIndex = 3;
            // 
            // progressBarBasic1
            // 
            this.progressBarBasic1.BackColor = System.Drawing.Color.White;
            this.progressBarBasic1.Font = new System.Drawing.Font("Arial", 10.25F);
            this.progressBarBasic1.ForeColor = System.Drawing.Color.ForestGreen;
            this.progressBarBasic1.Location = new System.Drawing.Point(12, 555);
            this.progressBarBasic1.Name = "progressBarBasic1";
            this.progressBarBasic1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            this.progressBarBasic1.Size = new System.Drawing.Size(803, 23);
            this.progressBarBasic1.TabIndex = 33;
            this.progressBarBasic1.Text = "progressBarBasic1";
            this.progressBarBasic1.TextStyle = ProgressBarBasic.ProgressBarBasic.TextStyleType.Percentage;
            this.progressBarBasic1.Visible = false;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(820, 556);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(69, 23);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "Cancel\r\n";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnDefaultSingle);
            this.groupBox2.Controls.Add(this.groupBox3);
            this.groupBox2.Controls.Add(this.label14);
            this.groupBox2.Controls.Add(this.txtRoughCull);
            this.groupBox2.Controls.Add(this.label13);
            this.groupBox2.Controls.Add(this.txtCull);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.txtTreeClCd);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.txtStatusCd);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.txtCR);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.txtActualHt);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.txtHt);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.txtDbh);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.txtSpCd);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.txtInvYr);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.txtVolLocGrp);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.txtStateCd);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.txtPlot);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.txtCountyCd);
            this.groupBox2.Controls.Add(this.btnTreeVolSingle);
            this.groupBox2.Location = new System.Drawing.Point(12, 330);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(877, 219);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Tree Volumes And Biomass One Record Test";
            // 
            // btnDefaultSingle
            // 
            this.btnDefaultSingle.Location = new System.Drawing.Point(18, 145);
            this.btnDefaultSingle.Name = "btnDefaultSingle";
            this.btnDefaultSingle.Size = new System.Drawing.Size(126, 40);
            this.btnDefaultSingle.TabIndex = 32;
            this.btnDefaultSingle.Text = "Default Values";
            this.btnDefaultSingle.UseVisualStyleBackColor = true;
            this.btnDefaultSingle.Click += new System.EventHandler(this.btnDefaultSingle_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.lblVOLTSGRS);
            this.groupBox3.Controls.Add(this.label21);
            this.groupBox3.Controls.Add(this.lblDRYBIOM);
            this.groupBox3.Controls.Add(this.label20);
            this.groupBox3.Controls.Add(this.lblDRYBIOT);
            this.groupBox3.Controls.Add(this.lblVOLCFNET);
            this.groupBox3.Controls.Add(this.label18);
            this.groupBox3.Controls.Add(this.label19);
            this.groupBox3.Controls.Add(this.lblVOLCSGRS);
            this.groupBox3.Controls.Add(this.label17);
            this.groupBox3.Controls.Add(this.lblVOLCFGRS);
            this.groupBox3.Controls.Add(this.label15);
            this.groupBox3.Location = new System.Drawing.Point(500, 15);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(303, 197);
            this.groupBox3.TabIndex = 31;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Results";
            // 
            // lblVOLTSGRS
            // 
            this.lblVOLTSGRS.BackColor = System.Drawing.Color.Beige;
            this.lblVOLTSGRS.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVOLTSGRS.Location = new System.Drawing.Point(152, 156);
            this.lblVOLTSGRS.Name = "lblVOLTSGRS";
            this.lblVOLTSGRS.Size = new System.Drawing.Size(142, 30);
            this.lblVOLTSGRS.TabIndex = 22;
            this.lblVOLTSGRS.Text = "0";
            this.lblVOLTSGRS.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(192, 135);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(65, 13);
            this.label21.TabIndex = 21;
            this.label21.Text = "VOLTSGRS";
            // 
            // lblDRYBIOM
            // 
            this.lblDRYBIOM.BackColor = System.Drawing.Color.Beige;
            this.lblDRYBIOM.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDRYBIOM.Location = new System.Drawing.Point(154, 93);
            this.lblDRYBIOM.Name = "lblDRYBIOM";
            this.lblDRYBIOM.Size = new System.Drawing.Size(142, 30);
            this.lblDRYBIOM.TabIndex = 20;
            this.lblDRYBIOM.Text = "0";
            this.lblDRYBIOM.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(185, 75);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(57, 13);
            this.label20.TabIndex = 19;
            this.label20.Text = "DRYBIOM";
            // 
            // lblDRYBIOT
            // 
            this.lblDRYBIOT.BackColor = System.Drawing.Color.Beige;
            this.lblDRYBIOT.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDRYBIOT.Location = new System.Drawing.Point(4, 156);
            this.lblDRYBIOT.Name = "lblDRYBIOT";
            this.lblDRYBIOT.Size = new System.Drawing.Size(142, 30);
            this.lblDRYBIOT.TabIndex = 18;
            this.lblDRYBIOT.Text = "0";
            this.lblDRYBIOT.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblVOLCFNET
            // 
            this.lblVOLCFNET.BackColor = System.Drawing.Color.Beige;
            this.lblVOLCFNET.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVOLCFNET.Location = new System.Drawing.Point(4, 93);
            this.lblVOLCFNET.Name = "lblVOLCFNET";
            this.lblVOLCFNET.Size = new System.Drawing.Size(142, 30);
            this.lblVOLCFNET.TabIndex = 16;
            this.lblVOLCFNET.Text = "0";
            this.lblVOLCFNET.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(41, 75);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(63, 13);
            this.label18.TabIndex = 15;
            this.label18.Text = "VOLCFNET";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(49, 133);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(55, 13);
            this.label19.TabIndex = 17;
            this.label19.Text = "DRYBIOT";
            // 
            // lblVOLCSGRS
            // 
            this.lblVOLCSGRS.BackColor = System.Drawing.Color.Beige;
            this.lblVOLCSGRS.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVOLCSGRS.Location = new System.Drawing.Point(154, 39);
            this.lblVOLCSGRS.Name = "lblVOLCSGRS";
            this.lblVOLCSGRS.Size = new System.Drawing.Size(142, 30);
            this.lblVOLCSGRS.TabIndex = 12;
            this.lblVOLCSGRS.Text = "0";
            this.lblVOLCSGRS.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(185, 17);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(65, 13);
            this.label17.TabIndex = 11;
            this.label17.Text = "VOLCSGRS";
            this.label17.Click += new System.EventHandler(this.label17_Click);
            // 
            // lblVOLCFGRS
            // 
            this.lblVOLCFGRS.BackColor = System.Drawing.Color.Beige;
            this.lblVOLCFGRS.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVOLCFGRS.Location = new System.Drawing.Point(4, 39);
            this.lblVOLCFGRS.Name = "lblVOLCFGRS";
            this.lblVOLCFGRS.Size = new System.Drawing.Size(142, 30);
            this.lblVOLCFGRS.TabIndex = 10;
            this.lblVOLCFGRS.Text = "0";
            this.lblVOLCFGRS.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(41, 17);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(64, 13);
            this.label15.TabIndex = 9;
            this.label15.Text = "VOLCFGRS";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(363, 79);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(56, 13);
            this.label14.TabIndex = 30;
            this.label14.Text = "RoughCull";
            // 
            // txtRoughCull
            // 
            this.txtRoughCull.Location = new System.Drawing.Point(425, 76);
            this.txtRoughCull.Name = "txtRoughCull";
            this.txtRoughCull.Size = new System.Drawing.Size(62, 20);
            this.txtRoughCull.TabIndex = 29;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(255, 79);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(24, 13);
            this.label13.TabIndex = 28;
            this.label13.Text = "Cull";
            // 
            // txtCull
            // 
            this.txtCull.Location = new System.Drawing.Point(286, 79);
            this.txtCull.Name = "txtCull";
            this.txtCull.Size = new System.Drawing.Size(71, 20);
            this.txtCull.TabIndex = 27;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(131, 82);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(51, 13);
            this.label12.TabIndex = 26;
            this.label12.Text = "TreeClCd";
            // 
            // txtTreeClCd
            // 
            this.txtTreeClCd.Location = new System.Drawing.Point(190, 79);
            this.txtTreeClCd.Name = "txtTreeClCd";
            this.txtTreeClCd.Size = new System.Drawing.Size(59, 20);
            this.txtTreeClCd.TabIndex = 25;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(15, 82);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(50, 13);
            this.label11.TabIndex = 24;
            this.label11.Text = "StatusCd";
            // 
            // txtStatusCd
            // 
            this.txtStatusCd.Location = new System.Drawing.Point(66, 79);
            this.txtStatusCd.Name = "txtStatusCd";
            this.txtStatusCd.Size = new System.Drawing.Size(59, 20);
            this.txtStatusCd.TabIndex = 23;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(15, 108);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(89, 13);
            this.label10.TabIndex = 22;
            this.label10.Text = "Crown Ratio (CR)";
            // 
            // txtCR
            // 
            this.txtCR.Location = new System.Drawing.Point(104, 105);
            this.txtCR.Name = "txtCR";
            this.txtCR.Size = new System.Drawing.Size(80, 20);
            this.txtCR.TabIndex = 21;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(363, 59);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(48, 13);
            this.label9.TabIndex = 20;
            this.label9.Text = "ActualHt";
            // 
            // txtActualHt
            // 
            this.txtActualHt.Location = new System.Drawing.Point(425, 53);
            this.txtActualHt.Name = "txtActualHt";
            this.txtActualHt.Size = new System.Drawing.Size(62, 20);
            this.txtActualHt.TabIndex = 19;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(255, 59);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(18, 13);
            this.label8.TabIndex = 18;
            this.label8.Text = "Ht";
            // 
            // txtHt
            // 
            this.txtHt.Location = new System.Drawing.Point(286, 55);
            this.txtHt.Name = "txtHt";
            this.txtHt.Size = new System.Drawing.Size(71, 20);
            this.txtHt.TabIndex = 17;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(131, 59);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(30, 13);
            this.label7.TabIndex = 16;
            this.label7.Text = "DBH";
            // 
            // txtDbh
            // 
            this.txtDbh.Location = new System.Drawing.Point(190, 55);
            this.txtDbh.Name = "txtDbh";
            this.txtDbh.Size = new System.Drawing.Size(59, 20);
            this.txtDbh.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(15, 59);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(33, 13);
            this.label6.TabIndex = 14;
            this.label6.Text = "SpCd";
            // 
            // txtSpCd
            // 
            this.txtSpCd.Location = new System.Drawing.Point(66, 55);
            this.txtSpCd.Name = "txtSpCd";
            this.txtSpCd.Size = new System.Drawing.Size(59, 20);
            this.txtSpCd.TabIndex = 13;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(363, 37);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(32, 13);
            this.label5.TabIndex = 12;
            this.label5.Text = "InvYr";
            // 
            // txtInvYr
            // 
            this.txtInvYr.Location = new System.Drawing.Point(425, 34);
            this.txtInvYr.Name = "txtInvYr";
            this.txtInvYr.Size = new System.Drawing.Size(62, 20);
            this.txtInvYr.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(190, 108);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(69, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Vol_Loc_Grp";
            // 
            // txtVolLocGrp
            // 
            this.txtVolLocGrp.Location = new System.Drawing.Point(265, 105);
            this.txtVolLocGrp.Name = "txtVolLocGrp";
            this.txtVolLocGrp.Size = new System.Drawing.Size(80, 20);
            this.txtVolLocGrp.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 37);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(45, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "StateCd";
            // 
            // txtStateCd
            // 
            this.txtStateCd.Location = new System.Drawing.Point(66, 34);
            this.txtStateCd.Name = "txtStateCd";
            this.txtStateCd.Size = new System.Drawing.Size(59, 20);
            this.txtStateCd.TabIndex = 7;
            this.txtStateCd.TextChanged += new System.EventHandler(this.txtStateCd_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(255, 37);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(25, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Plot";
            // 
            // txtPlot
            // 
            this.txtPlot.Location = new System.Drawing.Point(286, 34);
            this.txtPlot.Name = "txtPlot";
            this.txtPlot.Size = new System.Drawing.Size(71, 20);
            this.txtPlot.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(131, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "CountyCd";
            // 
            // txtCountyCd
            // 
            this.txtCountyCd.Location = new System.Drawing.Point(190, 34);
            this.txtCountyCd.Name = "txtCountyCd";
            this.txtCountyCd.Size = new System.Drawing.Size(59, 20);
            this.txtCountyCd.TabIndex = 3;
            // 
            // btnTreeVolSingle
            // 
            this.btnTreeVolSingle.Location = new System.Drawing.Point(193, 141);
            this.btnTreeVolSingle.Name = "btnTreeVolSingle";
            this.btnTreeVolSingle.Size = new System.Drawing.Size(98, 45);
            this.btnTreeVolSingle.TabIndex = 2;
            this.btnTreeVolSingle.Text = "Calculate Volume And Biomass";
            this.btnTreeVolSingle.UseVisualStyleBackColor = true;
            this.btnTreeVolSingle.Click += new System.EventHandler(this.btnTreeVolSingle_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblFSNetwork);
            this.groupBox1.Controls.Add(this.label16);
            this.groupBox1.Controls.Add(this.btnEdit);
            this.groupBox1.Controls.Add(this.btnLoad);
            this.groupBox1.Controls.Add(this.cmbDatasource);
            this.groupBox1.Controls.Add(this.btnLinkTableTest);
            this.groupBox1.Controls.Add(this.btnTreeVolBatch);
            this.groupBox1.Location = new System.Drawing.Point(3, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(886, 321);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Tree Volumes And Biomass Batch Test";
            // 
            // lblFSNetwork
            // 
            this.lblFSNetwork.AutoSize = true;
            this.lblFSNetwork.Location = new System.Drawing.Point(180, 16);
            this.lblFSNetwork.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblFSNetwork.Name = "lblFSNetwork";
            this.lblFSNetwork.Size = new System.Drawing.Size(22, 13);
            this.lblFSNetwork.TabIndex = 9;
            this.lblFSNetwork.Text = "NA";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(5, 16);
            this.label16.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(170, 13);
            this.label16.TabIndex = 8;
            this.label16.Text = "FCS Volume and Biomass Method:";
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(267, 293);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(114, 21);
            this.btnEdit.TabIndex = 7;
            this.btnEdit.Text = "Edit Selected Row";
            this.btnEdit.UseVisualStyleBackColor = true;
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(184, 292);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(77, 21);
            this.btnLoad.TabIndex = 6;
            this.btnLoad.Text = "Load";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // cmbDatasource
            // 
            this.cmbDatasource.FormattingEnabled = true;
            this.cmbDatasource.Items.AddRange(new object[] {
            "Tree Sample"});
            this.cmbDatasource.Location = new System.Drawing.Point(12, 293);
            this.cmbDatasource.Name = "cmbDatasource";
            this.cmbDatasource.Size = new System.Drawing.Size(166, 21);
            this.cmbDatasource.TabIndex = 5;
            this.cmbDatasource.Text = "Tree Sample";
            // 
            // btnLinkTableTest
            // 
            this.btnLinkTableTest.Location = new System.Drawing.Point(601, 293);
            this.btnLinkTableTest.Name = "btnLinkTableTest";
            this.btnLinkTableTest.Size = new System.Drawing.Size(249, 21);
            this.btnLinkTableTest.TabIndex = 3;
            this.btnLinkTableTest.Text = utils.FS_NETWORK_IS_AVAILABLE ? "Test FCS_TREE.BIOSUM_CALC Workflow" : "Test Link To Oracle BIOSUM_VOLUME table";
            this.btnLinkTableTest.UseVisualStyleBackColor = true;
            this.btnLinkTableTest.Click += new System.EventHandler(this.btnLinkTableTest_Click);
            // 
            // frmFCSTreeVolumeEdit
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(856, 467);
            this.Controls.Add(this.panel1);
            this.Name = "frmFCSTreeVolumeEdit";
            this.Text = "Tree Volume and Biomass Calculator Troubleshooter";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmFCSTreeVolumeEdit_FormClosing);
            this.Resize += new System.EventHandler(this.frmFCSTreeVolumeEdit_Resize);
            this.panel1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

    }

    #endregion

       
  
    private void LoadDefaultSingleRecordValues()
    {
        this.txtActualHt.Text = "80";
        this.txtCountyCd.Text = "87";
        this.txtCR.Text = "0";
        this.txtCull.Text = "0";
        this.txtDbh.Text = "20.0";
        this.txtHt.Text = "80";
        this.txtInvYr.Text = "2012";
        this.txtPlot.Text = "3633";
        this.txtRoughCull.Text = "0";
        this.txtStateCd.Text = "6";
        this.txtSpCd.Text = "202";
        this.txtStatusCd.Text = "1";
        this.txtTreeClCd.Text = "2";
        this.txtVolLocGrp.Text = "S26LCA";

     
    }

    private void btnTreeVolSingle_Click( object sender, EventArgs e )
    {
        lblDRYBIOM.Text = "0"; lblDRYBIOT.Text = "0"; lblVOLCFGRS.Text = "0";
        lblVOLCFNET.Text = "0"; lblVOLCSGRS.Text = "0"; lblVOLTSGRS.Text = "0";

        FIADBOracle.Services.FS_NETWORK_AVAILABLE = utils.FS_NETWORK_IS_AVAILABLE;
     
      FIADBOracle.Services m_oOracleServices = new FIADBOracle.Services();
      
      m_oOracleServices.Start();
      m_oOracleServices.m_oTree.GetVolumesMode = FIADBOracle.Services.Tree.GetVolumesModeValues.InsertRowTrigger;
      if( m_oOracleServices.m_intError == 0 )
      {
          try
          {

              m_oOracleServices.m_oTree.InstantiateNewBiosumTreeInputRecord();
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.ActualHt = Convert.ToInt32(txtActualHt.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CND_CN = "1";
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CountyCd = Convert.ToInt32(txtCountyCd.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CR = Convert.ToInt32(txtCR.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.Cull = Convert.ToInt32(txtCull.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.DBH = Convert.ToDouble(txtDbh.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.Ht = Convert.ToInt32(txtHt.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.InvYr = Convert.ToInt32(txtInvYr.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.Plot = Convert.ToInt32(txtPlot.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.PLT_CN = "1";
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.RecordId = 1;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.RoughCull = Convert.ToInt32(txtRoughCull.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.SpCd = Convert.ToInt32(txtSpCd.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.StateCd = Convert.ToInt32(txtStateCd.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.StatusCd = Convert.ToInt32(txtStatusCd.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.TRE_CN = "1";
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.TreeClCd = Convert.ToInt32(txtTreeClCd.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.Tree = 123456;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.Vol_Loc_Grp = txtVolLocGrp.Text.Trim();
               //START: ADDED BIOSUM_VOLUME COLUMNS
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.SiTree = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.UpperDia = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.UpperDiaHt = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.WoodlandStem = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CentroidDia = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CentroidDiaHtActual = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.SawHt = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.HtDmp = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.BoleHt = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CullCf = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CullFld = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CullDead = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CullForm = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CullMStop = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CfSnd = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.BfSnd = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.Precipitation = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.BaLive = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.DiaHtCd = null; 
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.StandingDeadCd = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.VolcfsndCalc = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.DrybioBoleCalc = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.DrybioTopCalc = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.DrybioSaplingCalc = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.DrybioWdldSppCalc = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CullBf = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.Subp = null;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.FormCl = null;
              //END: ADDED BIOSUM_VOLUME COLUMNS

              m_oOracleServices.m_oTree.AddBiosumRecord(m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord);
              if (m_oOracleServices.m_intError == 0)
              {
                  m_oOracleServices.m_oTree.GetBiosumVolumes();
                  if (m_oOracleServices.m_intError == 0) //TODO: m_intError == -5 after GetBiosumVolumes()
                  {
                      lblDRYBIOM.Text = m_oOracleServices.m_oTree.BiosumTreeInputRecordCollection.Item(0).DRYBIOM.ToString().Trim();
                      lblDRYBIOT.Text = m_oOracleServices.m_oTree.BiosumTreeInputRecordCollection.Item(0).DRYBIOT.ToString().Trim();
                      lblVOLCFGRS.Text = m_oOracleServices.m_oTree.BiosumTreeInputRecordCollection.Item(0).VOLCFGRS.ToString().Trim();
                      lblVOLCSGRS.Text = m_oOracleServices.m_oTree.BiosumTreeInputRecordCollection.Item(0).VOLCSGRS.ToString().Trim();
                      lblVOLCFNET.Text = m_oOracleServices.m_oTree.BiosumTreeInputRecordCollection.Item(0).VOLCFNET.ToString().Trim();
                      lblVOLTSGRS.Text = m_oOracleServices.m_oTree.BiosumTreeInputRecordCollection.Item(0).VOLTSGRS.ToString().Trim();

                  }
                  else
                  {
                      lblDRYBIOM.Text = "ERROR";
                      lblDRYBIOT.Text = "ERROR";
                      lblVOLCFGRS.Text = "ERROR";
                      lblVOLCSGRS.Text = "ERROR";
                      lblVOLCFNET.Text = "ERROR";
                      lblVOLTSGRS.Text = "ERROR";
                      MessageBox.Show(m_oOracleServices.m_strError, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Error);
                  }
              }
              else
              {
                  lblDRYBIOM.Text = "ERROR";
                  lblDRYBIOT.Text = "ERROR";
                  lblVOLCFGRS.Text = "ERROR";
                  lblVOLCSGRS.Text = "ERROR";
                  lblVOLCFNET.Text = "ERROR";
                  lblVOLTSGRS.Text = "ERROR";
                  MessageBox.Show(m_oOracleServices.m_strError, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Error);
              }
          }
          catch (Exception err)
          {
              MessageBox.Show(err.Message, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Error);
          }
      }
      else
        MessageBox.Show( m_oOracleServices.m_strError, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Error );

     

    }

    private void btnDefaultSingle_Click(object sender, EventArgs e)
    {
        LoadDefaultSingleRecordValues();
    }

    private void btnTreeVolBatch_Click(object sender, EventArgs e)
    {
        if (m_strGridTableSource.Trim().Length == 0) return;
        RunBatch_Start();
       
    }
    private void RunBatch_Start()
    {
        btnCancel.Visible = true;
        btnCancel.Invalidate();
        btnCancel.Refresh();
        groupBox1.Enabled = false;
        groupBox2.Enabled = false;
        groupBox3.Enabled = false;
        frmMain.g_oDelegate.InitializeThreadEvents();
        frmMain.g_oDelegate.m_oEventStopThread.Reset();
        frmMain.g_oDelegate.m_oEventThreadStopped.Reset();

        FIA_Biosum_Manager.utils.FS_NETWORK_CHECK();

        if (FIA_Biosum_Manager.utils.FS_NETWORK == utils.FS_NETWORK_STATUS.NotAvailable)
             frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(this.RunBatch_Main_XE));
        else
             frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(this.RunBatch_Main_FS_NETWORK));
        frmMain.g_oDelegate.m_oThread.IsBackground = true;
        frmMain.g_oDelegate.m_oThread.Start();
    }
    
    private void RunBatch_Main_XE()
    {
        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
        {
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//frmFCSTreeVolumeEdit.RunBatch_Main_XE \r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
        }
        string strWorkDbFile = "";
        string strConn = "";
        string strColumns = "";
        string strValues = "";
        int x = 0;
        int y = 0;
        System.Windows.Forms.CurrencyManager oCm;
        System.Data.DataView oDv;
        int intCurrRow = 0;
        int intRecordCount = 0;
        int intThermValue = 0;
        string strTable = "";

        FIADBOracle.Services.FS_NETWORK_AVAILABLE = false;
        dao_data_access oDao = new dao_data_access();
        System.Data.DataRow p_rowFound;

        frmMain.g_oDelegate.CurrentThreadProcessName = "main";
        frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
        
        strTable = m_strGridTableSource;
        
            intRecordCount = Convert.ToInt32(m_oAdo.getRecordCount(
                m_oAdo.getMDBConnString(m_strTempDBFile, "", ""),
                "SELECT COUNT(*) FROM " + strTable,
                strTable));

            frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Minimum", 0);
            frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Value", 0);
            frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Maximum", 100);
            frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Visible", true);
          
            

            //step 1 - delete old existing link
            if (oDao.TableExists(m_strTempDBFile, "fcs_biosum_volume"))
            {

                oDao.DeleteTableFromMDB(m_strTempDBFile, "fcs_biosum_volume");



            }
            intThermValue++;
           
            UpdateThermPercent(0,intRecordCount+8,intThermValue);
            //step 2 - create new link
            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Create Access Table Link To BIOSUM_VOLUME Oracle Table...Stand By");
            strWorkDbFile = m_strTempDBFile;
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);

            
            oDao.CreateOracleXETableLink("FIA Biosum Oracle Services", "fcs_biosum", "fcs", "FCS_BIOSUM", "BIOSUM_VOLUME", strWorkDbFile, "fcs_biosum_volume");

            oDao.m_DaoWorkspace.Close();
            oDao = null;
            
            intThermValue++;
            UpdateThermPercent(0, intRecordCount + 8, intThermValue);
            //step 3 - open ado connection
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDBFile, "", ""));
            intThermValue++;
            UpdateThermPercent(0, intRecordCount + 8, intThermValue);
            //step 4 - start oracle XE services
            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent,1,"Starting Oracle Services...Stand By");
            FIADBOracle.Services m_oOracleServices = new FIADBOracle.Services();
            m_oOracleServices.Start();
            intThermValue++;
            UpdateThermPercent(0, intRecordCount + 8, intThermValue);


            if (m_oOracleServices.m_oTree == null) MessageBox.Show("m_oTree==null");
            m_oOracleServices.m_oTree.GetVolumesMode = FIADBOracle.Services.Tree.GetVolumesModeValues.SQLUpdate;
            

            if (m_strGridTableSource.Trim() != Tables.VolumeAndBiomass.BiosumVolumesInputTable)
            {
                //step 5 - delete and create work tables
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, Tables.VolumeAndBiomass.BiosumVolumesInputTable))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                frmMain.g_oTables.m_oFvs.CreateOracleInputBiosumVolumesTable(m_oAdo, m_oAdo.m_OleDbConnection, Tables.VolumeAndBiomass.BiosumVolumesInputTable);

                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);
                frmMain.g_oTables.m_oFvs.CreateOracleInputFCSBiosumVolumesTable(m_oAdo, m_oAdo.m_OleDbConnection, Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);

                intThermValue++;
                UpdateThermPercent(0, intRecordCount + 8, intThermValue);

                //step 6 - insert records
                frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Prepare Tree Data For Oracle...Stand By");
                strColumns = "STATECD,COUNTYCD,PLOT,INVYR,VOL_LOC_GRP,TREE,SPCD,DIA,HT," +
                            "ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL,DECAYCD,TOTAGE," +
                            //START: ADDED BIOSUM_VOLUME COLUMNS
                            "SITREE,WDLDSTEM,UPPER_DIA,UPPER_DIA_HT," +
                            "CENTROID_DIA,CENTROID_DIA_HT_ACTUAL,SAWHT," +
                            "HTDMP,BOLEHT,CULLCF,CULL_FLD,CULLDEAD," +
                            "CULLFORM,CULLMSTOP,CFSND,BFSND,PRECIPITATION,BALIVE," +
                            "DIAHTCD,STANDING_DEAD_CD," +
                            //END: ADDED BIOSUM_VOLUME COLUMNS
                            "TRE_CN,CND_CN,PLT_CN";


                strValues = "CINT(MID(BIOSUM_COND_ID,6,2)) AS STATECD," +
                            "CINT(MID(BIOSUM_COND_ID,12,3)) AS COUNTYCD," +
                            "CINT(MID(BIOSUM_COND_ID,16,5)) AS PLOT," +
                            "INVYR,VOL_LOC_GRP,ID AS TREE,SPCD,DBH AS DIA,HT,ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL,DECAYCD,TOTAGE," +
                            //START: ADDED BIOSUM_VOLUME COLUMNS
                            "SITREE,WDLDSTEM,UPPER_DIA,UPPER_DIA_HT," +
                            "CENTROID_DIA,CENTROID_DIA_HT_ACTUAL,SAWHT," +
                            "HTDMP,BOLEHT,CULLCF,CULL_FLD,CULLDEAD," +
                            "CULLFORM,CULLMSTOP,CFSND,BFSND,PRECIPITATION,BALIVE," +
                            "DIAHTCD,STANDING_DEAD_CD," +
                            //END: ADDED BIOSUM_VOLUME COLUMNS
                            "CSTR(ID) AS TRE_CN," +
                            "BIOSUM_COND_ID AS CND_CN," +
                            "MID(BIOSUM_COND_ID,1,LEN(BIOSUM_COND_ID)-1) AS PLT_CN";

                m_oAdo.m_strSQL = "INSERT INTO " + Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable + " " +
                                 "(" + strColumns + ") SELECT " + strValues + " FROM " + strTable;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            else
            {
                strColumns = "STATECD,COUNTYCD,PLOT,INVYR,VOL_LOC_GRP,TREE,SPCD,DIA,HT," +
                            "ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL,DECAYCD,TOTAGE," +
                            //START: ADDED BIOSUM_VOLUME COLUMNS
                            "SITREE,WDLDSTEM,UPPER_DIA,UPPER_DIA_HT," +
                            "CENTROID_DIA,CENTROID_DIA_HT_ACTUAL,SAWHT," +
                            "HTDMP,BOLEHT,CULLCF,CULL_FLD,CULLDEAD," +
                            "CULLFORM,CULLMSTOP,CFSND,BFSND,PRECIPITATION,BALIVE," +
                            "DIAHTCD,STANDING_DEAD_CD," +
                            //END: ADDED BIOSUM_VOLUME COLUMNS
                            "TRE_CN,CND_CN,PLT_CN";
                intThermValue++;
                UpdateThermPercent(0, intRecordCount + 8, intThermValue);
            }

            m_oAdo.m_strSQL = "INSERT INTO fcs_biosum_volume (" + strColumns + ") SELECT " + strColumns + " FROM " + Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            intThermValue++;
            UpdateThermPercent(0, intRecordCount + 8, intThermValue);

            //step 7 - Get returned results from FCS Oracle 
            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Wait For Oracle Volume Compilation To Complete...Stand By");
            m_oOracleServices.m_oTree.GetBiosumVolumes();
            intThermValue++;
            UpdateThermPercent(0, intRecordCount + 8, intThermValue);
            //step 8 - Update grid with returned results
            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Update Grid With Volume Values...Stand By");
            if (m_oOracleServices.m_intError == 0)
            {

                strConn = m_oAdo.m_OleDbConnection.ConnectionString;
                m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);

                m_oAdo.OpenConnection(strConn);

                oCm = (CurrencyManager)this.BindingContext[uc_gridview1.m_dg.DataSource, uc_gridview1.m_dg.DataMember];
                oDv = (DataView)oCm.List;
                intCurrRow = uc_gridview1.m_intCurrRow - 1;
                y = oDv.Count;



                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, "SELECT * FROM fcs_biosum_volume");
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    DataColumn[] colPk = new DataColumn[1];
                    colPk[0] = uc_gridview1.m_ds.Tables[0].Columns["id"];
                    uc_gridview1.m_ds.Tables[0].PrimaryKey = colPk;
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        if (intThermValue < intRecordCount+8)
                        {
                            intThermValue++;
                            UpdateThermPercent(0, intRecordCount + 8, intThermValue);
                        }
                        System.Object[] p_searchID = new Object[1];
                        p_searchID[0] = Convert.ToInt32(m_oAdo.m_OleDbDataReader["tre_cn"]);
                        p_rowFound = uc_gridview1.m_ds.Tables[0].Rows.Find(p_searchID[0]);
                        if (p_rowFound != null)
                        {
                            if (m_oAdo.m_OleDbDataReader["VOLCSGRS_CALC"] != DBNull.Value)
                                p_rowFound["volcsgrs"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["VOLCSGRS_CALC"]);
                            else
                                p_rowFound["volcsgrs"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["VOLCFGRS_CALC"] != DBNull.Value)
                                p_rowFound["VOLCFGRS"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["VOLCFGRS_CALC"]);
                            else
                                p_rowFound["VOLCFGRS"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["VOLCFNET_CALC"] != DBNull.Value)
                                p_rowFound["VOLCFNET"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["VOLCFNET_CALC"]);
                            else
                                p_rowFound["VOLCFNET"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["DRYBIOM_CALC"] != DBNull.Value)
                                p_rowFound["DRYBIOM"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["DRYBIOM_CALC"]);
                            else
                                p_rowFound["DRYBIOM"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["DRYBIOT_CALC"] != DBNull.Value)
                                p_rowFound["DRYBIOT"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["DRYBIOT_CALC"]);
                            else
                                p_rowFound["DRYBIOT"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["VOLTSGRS_CALC"] != DBNull.Value)
                                p_rowFound["VOLTSGRS"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["VOLTSGRS_CALC"]);
                            else
                                p_rowFound["VOLTSGRS"] = DBNull.Value;
                        }
                    }
                }
                m_oAdo.m_OleDbDataReader.Close();
                UpdateThermPercent(0, intRecordCount + 8, intRecordCount + 8);
                System.Threading.Thread.Sleep(2000);
            }
            else
            {
                MessageBox.Show(m_oOracleServices.m_strError, "FIA Biosum", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Visible", false);
            frmMain.g_oDelegate.SetControlPropertyValue(btnCancel, "Visible", false);
        

        RunBatch_Finished();

       
        frmMain.g_oDelegate.CurrentThreadProcessDone = true;
        frmMain.g_oDelegate.m_oEventThreadStopped.Set();
        this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "---Leaving: frmFCSTreeVolumeEdit.RunBatch_Main \r\n");

    }

    private void RunBatch_Main_FS_NETWORK()
    {
        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
        {
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                "//frmFCSTreeVolumeEdit.RunBatch_Main_FS_NETWORK \r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
        }

        string strTable = m_strGridTableSource.Trim();
        var columnsAndDataTypes = Tables.VolumeAndBiomass.ColumnsAndDataTypes;
        string strColumns = "";
        string strValues = "";
        var COUNT = 0;
        int x = 0;
        int y = 0;
        System.Windows.Forms.CurrencyManager oCm;
        System.Data.DataView oDv;
        int intCurrRow = 0;
        int intRecordCount = 0;
        int intThermValue = 0;
        int intTotalRecs = 0;

        System.Data.DataRow p_rowFound;

        frmMain.g_oDelegate.CurrentThreadProcessName = "RunBatch_Main_FS_NETWORK";
        frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
        frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Minimum", 0);
        frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Value", 0);
        frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Maximum", 100);
        frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Visible", true);

        using (var conn = new System.Data.OleDb.OleDbConnection(m_oAdo.getMDBConnString(m_strTempDBFile, "", "")))
        {
            conn.Open();
            intRecordCount = Convert.ToInt32(m_oAdo.getRecordCount(conn, "SELECT COUNT(*) FROM " + strTable, strTable));

            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1,
                "Building inputs for BioSumComps Volume and Biomass calculations...Stand By");
            intThermValue++;
            UpdateThermPercent(0, intRecordCount *3 + 8, intThermValue);

            //Prepare FcsBiosumVolumesInputTable for SQLite fcs_tree.db & BioSumComps.jar if calculating for TreeSample or Tree_Work_Table built using master.tree, cond, plot
            if (strTable != Tables.VolumeAndBiomass.BiosumVolumesInputTable)
            {
                //delete and create work tables
                if (m_oAdo.TableExist(conn, Tables.VolumeAndBiomass.BiosumVolumesInputTable))
                    m_oAdo.SqlNonQuery(conn, "DROP TABLE " + Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                frmMain.g_oTables.m_oFvs.CreateOracleInputBiosumVolumesTable(m_oAdo, conn,
                    Tables.VolumeAndBiomass.BiosumVolumesInputTable);

                if (m_oAdo.TableExist(conn, Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable))
                    m_oAdo.SqlNonQuery(conn, "DROP TABLE " + Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);
                frmMain.g_oTables.m_oFvs.CreateOracleInputFCSBiosumVolumesTable(m_oAdo, conn,
                    Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);

                intThermValue++;
                UpdateThermPercent(0, intRecordCount *3 + 8, intThermValue);

                //step 6 - insert records
                frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1,
                    "Prepare Tree Data For BioSumComps Volume and Biomass calculations...Stand By");

                var treeToFcsBiosumVolumesInputTable = new List<Tuple<string, string>>
                {
                    Tuple.Create("STATECD", "CINT(MID(BIOSUM_COND_ID,6,2)) AS STATECD"),
                    Tuple.Create("COUNTYCD", "CINT(MID(BIOSUM_COND_ID,12,3)) AS COUNTYCD"),
                    Tuple.Create("PLOT", "CINT(MID(BIOSUM_COND_ID,16,5)) AS PLOT"),
                    Tuple.Create("INVYR", "INVYR"),
                    Tuple.Create("VOL_LOC_GRP", "VOL_LOC_GRP"),
                    Tuple.Create("TREE", "ID AS TREE"),
                    Tuple.Create("SPCD", "SPCD"),
                    Tuple.Create("DIA", "DBH AS DIA"),
                    Tuple.Create("HT", "HT"),
                    Tuple.Create("ACTUALHT", "ACTUALHT"),
                    Tuple.Create("CR", "CR"),
                    Tuple.Create("STATUSCD", "STATUSCD"),
                    Tuple.Create("TREECLCD", "TREECLCD"),
                    Tuple.Create("ROUGHCULL", "ROUGHCULL"),
                    Tuple.Create("CULL", "CULL"),
                    Tuple.Create("DECAYCD", "DECAYCD"),
                    Tuple.Create("TOTAGE", "TOTAGE"),
                    Tuple.Create("SITREE", "SITREE"),
                    Tuple.Create("WDLDSTEM", "WDLDSTEM"),
                    Tuple.Create("UPPER_DIA", "UPPER_DIA"),
                    Tuple.Create("UPPER_DIA_HT", "UPPER_DIA_HT"),
                    Tuple.Create("CENTROID_DIA", "CENTROID_DIA"),
                    Tuple.Create("CENTROID_DIA_HT_ACTUAL", "CENTROID_DIA_HT_ACTUAL"),
                    Tuple.Create("SAWHT", "SAWHT"),
                    Tuple.Create("HTDMP", "HTDMP"),
                    Tuple.Create("BOLEHT", "BOLEHT"),
                    Tuple.Create("CULLCF", "CULLCF"),
                    Tuple.Create("CULL_FLD", "CULL_FLD"),
                    Tuple.Create("CULLDEAD", "CULLDEAD"),
                    Tuple.Create("CULLFORM", "CULLFORM"),
                    Tuple.Create("CULLMSTOP", "CULLMSTOP"),
                    Tuple.Create("CFSND", "CFSND"),
                    Tuple.Create("BFSND", "BFSND"),
                    Tuple.Create("PRECIPITATION", "PRECIPITATION"),
                    Tuple.Create("BALIVE", "BALIVE"),
                    Tuple.Create("DIAHTCD", "DIAHTCD"),
                    Tuple.Create("STANDING_DEAD_CD", "STANDING_DEAD_CD"),
                    Tuple.Create("TRE_CN", "CSTR(ID) AS TRE_CN"),
                    Tuple.Create("CND_CN", "BIOSUM_COND_ID AS CND_CN"),
                    Tuple.Create("PLT_CN", "MID(BIOSUM_COND_ID,1,LEN(BIOSUM_COND_ID)-1) AS PLT_CN"),
                };
                strColumns = string.Join(",", treeToFcsBiosumVolumesInputTable.Select(e => e.Item1));
                strValues = string.Join(",", treeToFcsBiosumVolumesInputTable.Select(e => e.Item2));

                m_oAdo.m_strSQL = $"INSERT INTO {Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable} ({strColumns}) SELECT {strValues} FROM {strTable}";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                        m_oAdo.m_strSQL + "\r\n\r\n");
                m_oAdo.SqlNonQuery(conn, m_oAdo.m_strSQL);
            }

            intThermValue++;
            UpdateThermPercent(0, intRecordCount *3 + 8, intThermValue);

            strColumns = string.Join(",", columnsAndDataTypes.Select(item => item.Item1));

            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1,
                $"Clearing data from FCS_TREE.DB BIOSUM_CALC...Stand By");
            SQLite.ADO.DataMgr oSQLite = new SQLite.ADO.DataMgr();
            oSQLite.OpenConnection(false, 1,
                frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" +
                Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase, "BIOSUM_CALC");
            oSQLite.SqlNonQuery(oSQLite.m_Connection, $"DELETE FROM {Tables.VolumeAndBiomass.BiosumVolumeCalcTable}");
            System.Threading.Thread.Sleep(2000);


            //Parse MSAccess fcs_biosum_volumes_input and insert into SQLite FCS_TREE.Biosum_Calc table
            intTotalRecs = (int) m_oAdo.getRecordCount(conn,
                $"SELECT COUNT(*) AS ROWCOUNT FROM {Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable}",
                Tables.VolumeAndBiomass.SqliteWorkTable);
            m_oAdo.SqlQueryReader(conn, $"SELECT * FROM {Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable}");
            if (m_oAdo.m_OleDbDataReader.HasRows)
            {
                frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1,
                    $"Filling FCS_TREE.DB BIOSUM_CALC with data from {strTable}...Stand By");
                System.Data.SQLite.SQLiteTransaction transaction =
                    oSQLite.m_Connection.BeginTransaction(IsolationLevel.ReadCommitted);
                System.Data.SQLite.SQLiteCommand command = oSQLite.m_Connection.CreateCommand();
                command.Transaction = transaction;
                try
                {
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        intThermValue++;
                        UpdateThermPercent(0, intRecordCount *3 + 8, intThermValue);
                        strValues = utils.GetParsedInsertValues(m_oAdo.m_OleDbDataReader, columnsAndDataTypes);
                        command.CommandText = $"INSERT INTO {Tables.VolumeAndBiomass.BiosumVolumeCalcTable} ({strColumns}) VALUES ({strValues})";
                        command.ExecuteNonQuery();
//                        frmMain.g_oDelegate.SetControlPropertyValue(
//                            (System.Windows.Forms.Control) m_frmTherm.lblMsg, "Text",
//                            "Prepare Tree Data for Volume and Biomass Calculations...Stand By [" +
//                            COUNT.ToString() + "/" + intTotalRecs.ToString() + "]");
//                        frmMain.g_oDelegate.ExecuteControlMethod(
//                            (System.Windows.Forms.Control) this.m_frmTherm.lblMsg, "Refresh");
//                        if (COUNT <= intTotalRecs)
//                            UpdateThermPercent(0, intTotalRecs, COUNT);
                    }

                    transaction.Commit();
                }
                catch (Exception err)
                {
                    m_intError = -1;
                    MessageBox.Show(err.Message);
                    transaction.Rollback();
                }

                transaction.Dispose();
                transaction = null;
            }

            m_oAdo.m_OleDbDataReader.Close();
            m_oAdo.m_OleDbDataReader.Dispose();
            oSQLite.CloseAndDisposeConnection(oSQLite.m_Connection, true);
        }

        //RUN JAVA APP TO CALCULATE VOLUME/BIOMASS
        if (System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase) == false)
        {
            m_intError = -1;
            m_strError = frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase + " not found";
        }

        if (m_intError == 0 && System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\BioSumComps.JAR") == false)
        {
            m_intError = -1;
            m_strError = frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\BioSumComps.JAR not found";
        }

        if (m_intError == 0 && System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_tree_calc.bat") == false)
        {
            m_intError = -1;
            m_strError = frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_tree_calc.bat not found";
        }

        if (m_intError == 0)
        {
            frmMain.g_oUtils.RunProcess(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum", "fcs_tree_calc.bat", "BAT");
            if (System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_error_msg.txt"))
            {
                m_strError = System.IO.File.ReadAllText(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_error_msg.txt");
                if (m_strError.IndexOf("JAVA.EXE", 0) > 0)
                    m_strError = "Problem detected running JAVA.EXE";
                m_intError = -2;
            }
        }

        intThermValue++;
        UpdateThermPercent(0, intRecordCount *3 + 8, intThermValue);

        //step 7 - Get returned results from SQLite
        frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1,
            "Wait For BioSumComps.jar Volume and Biomass Calculations To Complete...Stand By");

        //Parse SQLite output and insert into Biosum_Calc_Output access table
        using (var conn = new System.Data.OleDb.OleDbConnection(m_oAdo.getMDBConnString(m_strTempDBFile, "", "")))
        {
            conn.Open();
            if (m_intError == 0)
            {
                SQLite.ADO.DataMgr oSQLite = new SQLite.ADO.DataMgr();
                oSQLite.OpenConnection(false, 1,
                    frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" +
                    Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase, "BIOSUM_CALC");

                if (m_oAdo.TableExist(conn, Tables.VolumeAndBiomass.BiosumCalcOutputTable)) 
                    m_oAdo.SqlNonQuery(conn, $"DROP TABLE {Tables.VolumeAndBiomass.BiosumCalcOutputTable}");

                System.Threading.Thread.Sleep(3000);

                m_oAdo.SqlNonQuery(conn, $"SELECT * INTO {Tables.VolumeAndBiomass.BiosumCalcOutputTable} FROM {Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable} WHERE 1=2");

                intTotalRecs = Convert.ToInt32(oSQLite.getSingleDoubleValueFromSQLQuery(oSQLite.m_Connection,
                    $"SELECT COUNT(*) AS ROWCOUNT FROM {Tables.VolumeAndBiomass.BiosumVolumeCalcTable} WHERE VOLTSGRS_CALC IS NOT NULL",
                    Tables.VolumeAndBiomass.BiosumVolumeCalcTable));

                UpdateThermPercent(0, intRecordCount *3 + 8, intThermValue);

                oSQLite.SqlQueryReader(oSQLite.m_Connection, $"SELECT * FROM {Tables.VolumeAndBiomass.BiosumVolumeCalcTable} WHERE VOLTSGRS_CALC IS NOT NULL");
                if (oSQLite.m_DataReader.HasRows)
                {
                    System.Data.OleDb.OleDbTransaction transaction;
                    System.Data.OleDb.OleDbCommand command = conn.CreateCommand();
                    // Start a local transaction
                    transaction = conn.BeginTransaction(IsolationLevel.ReadCommitted);
                    // Assign transaction object for a pending local transaction
                    command.Transaction = transaction;
                    try
                    {
                        while (oSQLite.m_DataReader.Read())
                        {
                            intThermValue++;
                            UpdateThermPercent(0, intRecordCount *3 + 8, intThermValue);
                            if (oSQLite.m_DataReader["TRE_CN"] != DBNull.Value && Convert.ToString(oSQLite.m_DataReader["TRE_CN"]).Trim().Length > 0)
                            {
                                strValues = utils.GetParsedInsertValues(oSQLite.m_DataReader, columnsAndDataTypes);
                                command.CommandText = $"INSERT INTO {Tables.VolumeAndBiomass.BiosumCalcOutputTable} ({strColumns}) VALUES ({strValues})";
                                command.ExecuteNonQuery();
                            }

//                                frmMain.g_oDelegate.SetControlPropertyValue(
//                                    (System.Windows.Forms.Control) m_frmTherm.lblMsg, "Text",
//                                    "Update MSAccess with SQLite Results..Stand By [" +
//                                    COUNT.ToString() + "/" + intTotalRecs.ToString() + "]");
//                                frmMain.g_oDelegate.ExecuteControlMethod(
//                                    (System.Windows.Forms.Control) this.m_frmTherm.lblMsg, "Refresh");
                        }

                        transaction.Commit();
                    }
                    catch (Exception err)
                    {
                        m_intError = -1;
                        MessageBox.Show(err.Message);
                        transaction.Rollback();
                    }
                    finally
                    {
                    }

                    transaction.Dispose();
                    transaction = null;
                }

                oSQLite.CloseAndDisposeConnection(oSQLite.m_Connection, true);
                oSQLite = null;
            }

            UpdateThermPercent(0, intRecordCount *3 + 8, intThermValue);

            //step 8 - Update grid with returned results
            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1,
                "Update Grid With Volume Values...Stand By");
            if (m_intError == 0)
            {
                //LPOTTS all data will be lost if connection is closed
                //strConn = m_oAdo.m_OleDbConnection.ConnectionString;
                //m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
                //m_oAdo.OpenConnection(strConn);

                oCm = (CurrencyManager) this.BindingContext[uc_gridview1.m_dg.DataSource,
                    uc_gridview1.m_dg.DataMember];
                oDv = (DataView) oCm.List;
                intCurrRow = uc_gridview1.m_intCurrRow - 1;
                y = oDv.Count;


                m_oAdo.SqlQueryReader(conn, $"SELECT * FROM {Tables.VolumeAndBiomass.BiosumCalcOutputTable}");
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    DataColumn[] colPk = new DataColumn[1];
                    colPk[0] = uc_gridview1.m_ds.Tables[0].Columns["id"];
                    uc_gridview1.m_ds.Tables[0].PrimaryKey = colPk;
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        if (intThermValue < intRecordCount *3 + 8)
                        {
                            intThermValue++;
                            UpdateThermPercent(0, intRecordCount *3 + 8, intThermValue);
                        }

                        System.Object[] p_searchID = new Object[1];
                        p_searchID[0] = Convert.ToInt32(m_oAdo.m_OleDbDataReader["tre_cn"]);
                        p_rowFound = uc_gridview1.m_ds.Tables[0].Rows.Find(p_searchID[0]);
                        if (p_rowFound != null)
                        {
                            if (m_oAdo.m_OleDbDataReader["DRYBIOM_CALC"] != DBNull.Value)
                                p_rowFound["DRYBIOM"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["DRYBIOM_CALC"]);
                            else p_rowFound["DRYBIOM"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["DRYBIOT_CALC"] != DBNull.Value)
                                p_rowFound["DRYBIOT"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["DRYBIOT_CALC"]);
                            else p_rowFound["DRYBIOT"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["DRYBIO_BOLE_CALC"] != DBNull.Value)
                                p_rowFound["DRYBIO_BOLE"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["DRYBIO_BOLE_CALC"]);
                            else p_rowFound["DRYBIO_BOLE"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["DRYBIO_SAPLING_CALC"] != DBNull.Value)
                                p_rowFound["DRYBIO_SAPLING"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["DRYBIO_SAPLING_CALC"]);
                            else p_rowFound["DRYBIO_SAPLING"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["DRYBIO_TOP_CALC"] != DBNull.Value)
                                p_rowFound["DRYBIO_TOP"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["DRYBIO_TOP_CALC"]);
                            else p_rowFound["DRYBIO_TOP"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["DRYBIO_WDLD_SPP_CALC"] != DBNull.Value)
                                p_rowFound["DRYBIO_WDLD_SPP"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["DRYBIO_WDLD_SPP_CALC"]);
                            else p_rowFound["DRYBIO_WDLD_SPP"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["VOLCFGRS_CALC"] != DBNull.Value)
                                p_rowFound["VOLCFGRS"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["VOLCFGRS_CALC"]);
                            else p_rowFound["VOLCFGRS"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["VOLCFNET_CALC"] != DBNull.Value)
                                p_rowFound["VOLCFNET"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["VOLCFNET_CALC"]);
                            else p_rowFound["VOLCFNET"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["VOLCFSND_CALC"] != DBNull.Value)
                                p_rowFound["VOLCFSND"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["VOLCFSND_CALC"]);
                            else p_rowFound["VOLCFSND"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["VOLCSGRS_CALC"] != DBNull.Value)
                                p_rowFound["VOLCSGRS"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["VOLCSGRS_CALC"]);
                            else p_rowFound["VOLCSGRS"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["VOLTSGRS_CALC"] != DBNull.Value)
                                p_rowFound["VOLTSGRS"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["VOLTSGRS_CALC"]);
                            else p_rowFound["VOLTSGRS"] = DBNull.Value;
                        }
                    }
                }

                m_oAdo.m_OleDbDataReader.Close();
                m_oAdo.m_OleDbDataReader.Dispose();
                UpdateThermPercent(0, intRecordCount *3 + 8, intRecordCount *3 + 8);
                System.Threading.Thread.Sleep(2000);
            }
            else
            {
                MessageBox.Show(m_strError, "FIA Biosum", MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Visible", false);
        frmMain.g_oDelegate.SetControlPropertyValue(btnCancel, "Visible", false);


        RunBatch_Finished();


        frmMain.g_oDelegate.CurrentThreadProcessDone = true;
        frmMain.g_oDelegate.m_oEventThreadStopped.Set();
        this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "---Leaving: frmFCSTreeVolumeEdit.RunBatch_Main \r\n");
    }

    private void RunBatch_Finished()
    {
        frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Visible", false);
        frmMain.g_oDelegate.SetControlPropertyValue(btnCancel, "Visible", false);
        frmMain.g_oDelegate.SetControlPropertyValue(groupBox1, "Enabled", true);
        frmMain.g_oDelegate.SetControlPropertyValue(groupBox2, "Enabled", true);
        frmMain.g_oDelegate.SetControlPropertyValue(groupBox3, "Enabled", true);
        frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Ready");
    }
    private void UpdateThermPercent(int p_intMin, int p_intMax, int p_intValue)
    {
        int intPercent = (int)(((double)(p_intValue - p_intMin) /
            (double)(p_intMax - p_intMin)) * 100);

        frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Value", intPercent);

    }
    private void button1_Click(object sender, EventArgs e)
    {
    }
       

    private void btnLoad_Click(object sender, EventArgs e)
    {
        if (cmbDatasource.Text.Trim().ToUpper() == "TREE SAMPLE") LoadTreeSample();
        else if (cmbDatasource.Text.Trim().ToUpper() == "TREE TABLE")
        {
            
           
            if (m_oAdo.m_OleDbConnection == null || m_oAdo.m_OleDbConnection.State == ConnectionState.Closed)
                m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDBFile, "", ""));

                m_oAdo.m_strSQL = "SELECT count(fvs_variant) FROM " + m_oQueries.m_oFIAPlot.m_strPlotTable + " " +
                                  "WHERE fvs_variant IS NOT NULL AND LEN(TRIM(fvs_variant)) > 0";

                if ((int)m_oAdo.getRecordCount(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL, "fvs_variant") > 0)

                {
                    FIA_Biosum_Manager.frmDialog oDlg = new frmDialog();

                oDlg.uc_select_list_item1.lblTitle.Text = "FVS Variant";
                oDlg.uc_select_list_item1.listBox1.Sorted = true;
                oDlg.uc_select_list_item1.lblMsg.Hide();


                oDlg.uc_select_list_item1.loadvalues(m_oAdo, m_oAdo.m_OleDbConnection,
                                                "SELECT DISTINCT fvs_variant FROM " + m_oQueries.m_oFIAPlot.m_strPlotTable, "fvs_variant");
                oDlg.uc_select_list_item1.Show();

                DialogResult result = oDlg.ShowDialog();
                if (result == DialogResult.OK)
                {
                    if (oDlg.uc_select_list_item1.listBox1.SelectedItems.Count > 0)
                    {
                        LoadTreeTable(oDlg.uc_select_list_item1.listBox1.SelectedItems[0].ToString().Trim());
                    }
                }
                oDlg.Dispose();
                oDlg = null;
            }
            else
            {
                LoadTreeTable("");
            }
        }
        else
        {
           
            LoadFVSOutTrees(cmbDatasource.Text.Trim().Substring(16,3));
            
        }
    }
    private void LoadFVSOutTrees(string p_strRxPackage)
    {
        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
        {
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//frmFCSTreeVolumeEdit.LoadFVSOutTrees \r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
        }
        frmMain.g_oFrmMain.ActivateStandByAnimation(
            this.WindowState,
            this.Left,
            this.Height,
            this.Width,
            this.Top);
        frmMain.g_sbpInfo.Text = "Loading FVS Out Tree Table Data...Stand By";

        if (m_oAdo.m_OleDbConnection == null || m_oAdo.m_OleDbConnection.State == ConnectionState.Closed)
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDBFile, "", ""));


        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, Tables.VolumeAndBiomass.BiosumVolumesInputTable))
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + Tables.VolumeAndBiomass.BiosumVolumesInputTable);

        frmMain.g_oTables.m_oFvs.CreateOracleInputBiosumVolumesTable(m_oAdo, m_oAdo.m_OleDbConnection, Tables.VolumeAndBiomass.BiosumVolumesInputTable);

        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable))
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);

        frmMain.g_oTables.m_oFvs.CreateOracleInputFCSBiosumVolumesTable(m_oAdo, m_oAdo.m_OleDbConnection, Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);

        var strFvsTreeTable = this.cmbDatasource.Text; //e.g., fvs_tree_IN_BM_P009_TREE_CUTLIST
        var regex = new Regex(@"fvs_tree_IN_([A-Z]{2})_(P\d{3}).*");
        var strVariant = regex.Match(strFvsTreeTable).Groups[1];
        var strPackage = regex.Match(strFvsTreeTable).Groups[2];
        var strFvsSummaryTable = $"fvs_summary_IN_{strVariant}_{strPackage}"; //e.g., fvs_tree_IN_BM_P009_TREE_CUTLIST
        var strFiaTreeSpeciesRefTableLink = Tables.ProcessorScenarioRun.DefaultFiaTreeSpeciesRefTableName;

        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "cull_work_table"))
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE cull_work_table");

        m_oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step1(
            Tables.VolumeAndBiomass.BiosumVolumesInputTable, cmbDatasource.Text.Trim(),p_strRxPackage);
        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        m_oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step2(
            Tables.VolumeAndBiomass.BiosumVolumesInputTable,
            m_oQueries.m_oFIAPlot.m_strTreeTable,
            m_oQueries.m_oFIAPlot.m_strPlotTable,
            m_oQueries.m_oFIAPlot.m_strCondTable);
        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        //Set DIAHTCD for FIADB Cycle<>1 trees to their Cycle=1 DIAHTCD values
        m_oAdo.m_strSQL = $@"UPDATE {Tables.VolumeAndBiomass.BiosumVolumesInputTable} b 
                             INNER JOIN {m_oQueries.m_oFIAPlot.m_strTreeTable} t 
                             ON t.biosum_cond_id=b.biosum_cond_id AND t.fvs_tree_id=b.fvs_tree_id
                             SET b.diahtcd=t.diahtcd WHERE b.fvscreatedtree_yn='N'";

        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        //Set DIAHTCD for FVS-Created trees using FIA_TREE_SPECIES_REF.WOODLAND_YN
        m_oAdo.m_strSQL = $@"UPDATE {Tables.VolumeAndBiomass.BiosumVolumesInputTable} b 
                             INNER JOIN {strFiaTreeSpeciesRefTableLink} ref ON cint(b.spcd)=ref.spcd
                             SET b.diahtcd=IIF(ref.woodland_yn='N', 1, 2) WHERE b.fvscreatedtree_yn='Y'";
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        //Update FVSCreatedTrees balive=fvs_summary.BA
        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, strFvsSummaryTable))
        {
            m_oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step2a(
                Tables.VolumeAndBiomass.BiosumVolumesInputTable, m_oQueries.m_oFIAPlot.m_strPlotTable,
                m_oQueries.m_oFIAPlot.m_strCondTable, strFvsSummaryTable);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
        }

        m_oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step3(
            Tables.VolumeAndBiomass.BiosumVolumesInputTable,
            m_oQueries.m_oFIAPlot.m_strCondTable);
        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        m_oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step4(
            "cull_work_table",
            Tables.VolumeAndBiomass.BiosumVolumesInputTable);
        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        m_oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.PNWRS.BuildInputTableForVolumeCalculation_Step5(
            "cull_work_table",
            Tables.VolumeAndBiomass.BiosumVolumesInputTable);
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        m_oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.PNWRS.BuildInputTableForVolumeCalculation_Step6(
            "cull_work_table",
            Tables.VolumeAndBiomass.BiosumVolumesInputTable);
        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        m_oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step7(
            Tables.VolumeAndBiomass.BiosumVolumesInputTable,
            Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);
        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        uc_gridview1.LoadGridView(
            m_oAdo.getMDBConnString(m_strTempDBFile, "", ""),
            "SELECT DRYBIOM," +
                   "DRYBIOT," +
                   "DRYBIO_BOLE_CALC AS DRYBIO_BOLE," +
                   "DRYBIO_SAPLING_CALC AS DRYBIO_SAPLING," +
                   "DRYBIO_TOP_CALC AS DRYBIO_TOP," +
                   "DRYBIO_WDLD_SPP_CALC AS DRYBIO_WDLD_SPP," +
                   "VOLCFGRS," +
                   "VOLCFNET," +
                   "VOLCFSND_CALC AS VOLCFSND," +
                   "VOLCSGRS," +
                   "VOLTSGRS," +
                   "id," +
                   "biosum_cond_id, " +
                   "fvs_tree_id," +
                   "MID(biosum_cond_id, 6, 2 ) AS state," +
                   "MID(biosum_cond_id,12,3) AS county," +
                   "MID(biosum_cond_id,15,7) AS plot," + 
                   "fvs_variant," +
                   "InvYr," +
                   "SpCd," +
                   "Dbh," +
                   "ROUND(Ht,0) AS Ht," +
                   "vol_loc_grp," +
                   "IIF(actualht IS NULL,Ht,ROUND(actualht,0)) AS actualht," +
                   "statuscd," +
                   "treeclcd," +
                   "cr," +
                   "cull," +
                   "IIF(roughcull IS NULL,0,roughcull) AS roughcull," +
                   "decaycd," +
                   "totage," +
                   //START: ADDED BIOSUM_VOLUME COLUMNS
                   "sitree," + 
                   "wdldstem," + 
                   "upper_dia," + 
                   "upper_dia_ht," + 
                   "centroid_dia," + 
                   "centroid_dia_ht_actual," + 
                   "sawht," + 
                   "htdmp," + 
                   "boleht," + 
                   "cullcf," + 
                   "cull_fld," + 
                   "culldead," + 
                   "cullform," + 
                   "cullmstop," + 
                   "cfsnd," +
                   "bfsnd," + 
                   "precipitation," + 
                   "balive," +
                   "diahtcd," +
                   "standing_dead_cd " +
                   //END: ADDED BIOSUM_VOLUME COLUMNS
             "FROM " + Tables.VolumeAndBiomass.BiosumVolumesInputTable, this.cmbDatasource.Text.Trim());

        m_strGridTableSource = Tables.VolumeAndBiomass.BiosumVolumesInputTable;
        frmMain.g_oFrmMain.DeactivateStandByAnimation();
        frmMain.g_sbpInfo.Text = "Ready";

        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,"---frmFCSTreeVolumeEdit.LoadFVSOutTrees \r\n");



    }
    private void LoadTreeSample()
    {
        if (m_oAdo.m_OleDbConnection != null)
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
        uc_gridview1.LoadGridView(
            m_oAdo.getMDBConnString(m_strTempDBFile, "", ""),
            "SELECT DRYBIOM," +
                   "DRYBIOT," +
                   "DRYBIO_BOLE," +
                   "DRYBIO_SAPLING," +
                   "DRYBIO_TOP," +
                   "DRYBIO_WDLD_SPP," +
                   "VOLCFGRS," +
                   "VOLCFNET," +
                   "VOLCFSND," +
                   "VOLCSGRS," +
                   "VOLTSGRS," + 
                   "id," +
                   "biosum_cond_id, " +
                   "fvs_tree_id," +
                   "MID(biosum_cond_id, 6, 2 ) AS state," + 
                   "MID(biosum_cond_id,12,3) AS county," + 
                   "MID(biosum_cond_id,15,7) AS plot," + 
                   "fvs_variant," + 
                   "InvYr," +
                   "SpCd," +
                   "Dbh," +
                   "Ht," +
                   "vol_loc_grp," +
                   "IIF(actualht IS NULL,Ht,actualht) AS actualht," +
                   "statuscd," +
                   "treeclcd," +
                   "cr," +
                   "cull," +
                   "IIF(roughcull IS NULL,0,roughcull) AS roughcull," +
                   "decaycd," +
                   "totage," +
                   //START: ADDED BIOSUM_VOLUME COLUMNS
                   "sitree," +
                   "wdldstem," +
                   "upper_dia," +
                   "upper_dia_ht," +
                   "centroid_dia," +
                   "centroid_dia_ht_actual," +
                   "sawht," +
                   "htdmp," +
                   "boleht," +
                   "cullcf," +
                   "cull_fld," +
                   "culldead," +
                   "cullform," +
                   "cullmstop," +
                   "cfsnd," +
                   "bfsnd," +
                   "precipitation," +
                   "balive," +
                   "diahtcd," +
                   "standing_dead_cd " +
                   //END: ADDED BIOSUM_VOLUME COLUMNS
             "FROM TreeSample", "TreeSample");

        m_strGridTableSource = "TreeSample";
    }

    private void LoadTreeTable(string p_strFVSVariant)
    {
        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
        {
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//frmFCSTreeVolumeEdit.LoadTreeTable \r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
        }
        frmMain.g_oFrmMain.ActivateStandByAnimation(
          this.WindowState,
          this.Left,
          this.Height,
          this.Width,
          this.Top);

        if (m_oAdo.m_OleDbConnection == null)
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDBFile, "", ""));

        //
        //CREATE TREE TABLE WORK TABLE
        //
        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "tree_work_table"))
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection,"DROP TABLE tree_work_table");

            frmMain.g_sbpInfo.Text = "Loading Tree Table Data...Stand By";
            m_oAdo.m_strSQL = frmMain.g_oTables.m_oFvs.CreateOracleInputBiosumVolumesTableSQL("tree_work_table");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            m_oAdo.AddAutoNumber(m_oAdo.m_OleDbConnection, "tree_work_table", "id");
            //
            //POPULATE TREE WORK TABLE
            //
            m_oAdo.m_strSQL = "INSERT INTO tree_work_table " + 
                              "SELECT t.biosum_cond_id," + 
                                     "IIF(p.InvYr IS NULL AND p.MeasYear IS NOT NULL,p.MeasYear,IIF(p.InvYr IS NOT NULL,p.InvYr,null)) AS InvYr," + 
                                     "p.fvs_variant," + 
                                     "t.spcd, t.dia AS dbh," + 
                                     "t.ht, c.vol_loc_grp," + 
                                     "IIF(t.actualht IS NULL,t.Ht,t.actualht) AS actualht," + 
                                     "t.statuscd, t.treeclcd," + 
                                     "IIF(t.cr IS NULL,0,t.cr) AS cr," + 
                                     "IIF(t.cull IS NULL,0,ROUND(t.cull,0)) AS cull," + 
                                     "IIF(t.roughcull IS NULL,0,ROUND(t.roughcull,0)) AS roughcull," + 
                                     "IIF(t.decaycd IS NULL,0,t.decaycd) AS decaycd," + 
                                     "t.totage, t.diahtcd, t.standing_dead_cd, c.balive, p.precipitation, " + 
                                     "t.fvs_tree_id " + 
                             "FROM " + m_oQueries.m_oFIAPlot.m_strTreeTable + " t," + 
                                       m_oQueries.m_oFIAPlot.m_strCondTable + " c," + 
                                       m_oQueries.m_oFIAPlot.m_strPlotTable + " p " + 
                             "WHERE c.biosum_cond_id = t.biosum_cond_id AND " + 
                                   "p.biosum_plot_id = c.biosum_plot_id AND " + 
                                   "(p.fvs_variant IS NULL OR p.fvs_variant='" + p_strFVSVariant + "') AND " + 
                                   "(p.InvYr IS NOT NULL OR p.MeasYear IS NOT NULL)";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            //
            //update columns
            //
            //total cull
            //populate treeclcd column
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "cull_total_work_table") == true)
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE cull_total_work_table");

            m_oAdo.m_strSQL = "SELECT id, IIF(cull IS NOT NULL AND roughcull IS NOT NULL, cull + roughcull," +
                                           "IIF(cull IS NOT NULL,cull," +
                                           "IIF(roughcull IS NOT NULL, roughcull,0))) AS totalcull " +
                              "INTO cull_total_work_table " +
                              "FROM Tree_Work_Table";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            m_oAdo.m_strSQL = "UPDATE Tree_Work_Table a " +
                            "INNER JOIN cull_total_work_table b " +
                            "ON a.id=b.id " +
                            "SET a.treeclcd=" +
                            "IIF(a.SpCd IN (62,65,66,106,133,138,304,321,322,475,756,758,990),3," +
                            "IIF(a.StatusCd=2,3," +
                            "IIF(b.totalcull < 75,2," +
                            "IIF(roughcull > 37.5,3,4))))";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            m_oAdo.m_strSQL = "UPDATE tree_work_table  a " +
                                     "INNER JOIN cull_total_work_table b " +
                                     "ON a.id=b.id " +
                                     "SET a.treeclcd=" +
                                     "IIF(a.DecayCd > 1,4,IIF(a.dbh < 9 AND a.SpCd < 300,4,a.treeclcd)) " +
                                     "WHERE a.treeclcd=3 AND a.statuscd=2 AND a.SpCd NOT IN (62,65,66,106,133,138,304,321,322,475,756,758,990)";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_oAdo.m_strSQL + "\r\n\r\n");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDBFile, "", ""));

        uc_gridview1.LoadGridView(
            m_oAdo.getMDBConnString(m_strTempDBFile, "", ""),
            "SELECT DRYBIOM," +
                   "DRYBIOT," +
                   "DRYBIO_BOLE_CALC AS DRYBIO_BOLE," +
                   "DRYBIO_SAPLING_CALC AS DRYBIO_SAPLING," +
                   "DRYBIO_TOP_CALC AS DRYBIO_TOP," +
                   "DRYBIO_WDLD_SPP_CALC AS DRYBIO_WDLD_SPP," +
                   "VOLCFGRS," +
                   "VOLCFNET," +
                   "VOLCFSND_CALC AS VOLCFSND," +
                   "VOLCSGRS," +
                   "VOLTSGRS," +
                   "id," +
                   "biosum_cond_id, " +
                   "fvs_tree_id," +
                   "MID(biosum_cond_id, 6, 2 ) AS state," +
                   "MID(biosum_cond_id,12,3) AS county," +
                   "MID(biosum_cond_id,15,7) AS plot," +
                   "fvs_variant," + 
                   "InvYr," +
                   "SpCd," +
                   "Dbh," +
                   "Ht," +
                   "vol_loc_grp," +
                   "IIF(actualht IS NULL,Ht,actualht) AS actualht," +
                   "statuscd," +
                   "treeclcd," +
                   "cr," +
                   "cull," +
                   "IIF(roughcull IS NULL,0,roughcull) AS roughcull," +
                   "decaycd," + 
                   "totage," +
                   //START: ADDED BIOSUM_VOLUME COLUMNS
                   "sitree," + 
                   "wdldstem," + 
                   "upper_dia," + 
                   "upper_dia_ht," + 
                   "centroid_dia," + 
                   "centroid_dia_ht_actual," + 
                   "sawht," + 
                   "htdmp," + 
                   "boleht," + 
                   "cullcf," + 
                   "cull_fld," + 
                   "culldead," + 
                   "cullform," + 
                   "cullmstop," + 
                   "cfsnd," + 
                   "bfsnd," + 
                   "precipitation," + 
                   "balive," +
                   "diahtcd," +
                   "standing_dead_cd " +
                   //END: ADDED BIOSUM_VOLUME COLUMNS
             "FROM Tree_Work_Table", m_oQueries.m_oFIAPlot.m_strTreeTable);

        m_strGridTableSource = "Tree_Work_Table";
        frmMain.g_oFrmMain.DeactivateStandByAnimation();
        frmMain.g_sbpInfo.Text = "Ready";
        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "---Leaving: frmFCSTreeVolumeEdit.LoadTreeTable \r\n");
    }


    private void btnEdit_Click(object sender, EventArgs e)
    {
        this.txtActualHt.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1,COL_ACTUALHT].ToString().Trim() ;
        this.txtCountyCd.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_COUNTY].ToString().Trim();
        this.txtCR.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_CR].ToString().Trim();
        this.txtCull.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_CULL].ToString().Trim();
        this.txtDbh.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_DBH].ToString().Trim();
        this.txtHt.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_HT].ToString().Trim();
        this.txtInvYr.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_INVYR].ToString().Trim();
        this.txtPlot.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_PLOT].ToString().Trim();
        this.txtRoughCull.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1,COL_ROUGHCULL].ToString().Trim();
        this.txtStateCd.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_STATE].ToString().Trim();
        this.txtSpCd.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_SPCD].ToString().Trim();
        this.txtStatusCd.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_STATUSCD].ToString().Trim();
        this.txtTreeClCd.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_TREECLCD].ToString().Trim();
        this.txtVolLocGrp.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_VOLLOCGRP].ToString().Trim();
        
    }

    private void frmFCSTreeVolumeEdit_Resize(object sender, EventArgs e)
    {
        ResizeForm();
    }
    private void ResizeForm()
    {
        //width
        groupBox2.Width = this.ClientSize.Width - (int)(groupBox2.Left * 2);
        groupBox1.Width = groupBox2.Width;
        btnCancel.Left = this.ClientSize.Width - btnCancel.Width - 5;

        progressBarBasic1.Width = btnCancel.Left - progressBarBasic1.Left - 3;
        lblPerc.Left = (int)(progressBarBasic1.Width / 2) - (int)(lblPerc.Width * .5);

        //top
        progressBarBasic1.Top = this.ClientSize.Height - progressBarBasic1.Height - 5;
        btnCancel.Top = progressBarBasic1.Top;
        groupBox2.Top = progressBarBasic1.Top - groupBox2.Height - 5;

        //height
        groupBox1.Height = groupBox2.Top - groupBox1.Top - 5;
        btnEdit.Top = groupBox1.Height - btnEdit.Height - 5;
        cmbDatasource.Top = btnEdit.Top;
        btnLoad.Top = btnEdit.Top;
        btnTreeVolBatch.Top = btnEdit.Top;
        btnLinkTableTest.Top = btnEdit.Top;
        uc_gridview1.Height = cmbDatasource.Top - uc_gridview1.Top - 5;
        uc_gridview1.Width = groupBox1.Width - (uc_gridview1.Left * 2) - 5;

    }

    private void btnCancel_Click(object sender, EventArgs e)
    {
        frmMain.g_oDelegate.CurrentThreadProcessSuspended = true;
        frmMain.g_oDelegate.m_oThread.Suspend();
        bool bAbort = frmMain.g_oDelegate.AbortProcessing("FIA Biosum", "Cancel Running The Volume and Biosum Calculator (Y/N)?");
        if (bAbort)
        {
            
            if (frmMain.g_oDelegate.m_oThread.IsAlive)
            {
                frmMain.g_oDelegate.m_oThread.Join();
            }
            frmMain.g_oDelegate.StopThread();
            RunBatch_Finished();

            frmMain.g_oDelegate.m_oThread = null;

        }
        else frmMain.g_oDelegate.m_oThread.Resume();
    }

    private void frmFCSTreeVolumeEdit_FormClosing(object sender, FormClosingEventArgs e)
    {
        if (frmMain.g_oDelegate.CurrentThreadProcessStarted == true &&
            frmMain.g_oDelegate.CurrentThreadProcessDone == false)
        {
            btnCancel_Click(null, null);
            if (frmMain.g_oDelegate.CurrentThreadProcessAborted == true)
            {
            }
            else
            {
                e.Cancel = true; return;
            }

            

        }
        this.Dispose();
    }

    private void btnLinkTableTest_Click(object sender, EventArgs e)
    {

        frmMain.g_oDelegate.InitializeThreadEvents();
        frmMain.g_oDelegate.m_oEventStopThread.Reset();
        frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
        frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(this.TestExternalDatabaseConnection));
        frmMain.g_oDelegate.m_oThread.IsBackground = true;
        frmMain.g_oDelegate.m_oThread.Start();

    
    }
    private void TestExternalDatabaseConnection()
    {
        frmMain.g_oDelegate.CurrentThreadProcessDone = false;
        string strFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
        string str="";
        if (utils.FS_NETWORK_IS_NOT_AVAILABLE)
        {
            frmMain.g_oFrmMain.ActivateStandByAnimation(this.WindowState, this.Left, this.Height, this.Width, this.Top);
            FIA_Biosum_Manager.dao_data_access oDao = new FIA_Biosum_Manager.dao_data_access();
            oDao.CreateMDB(strFile);

            try
            {
                str = "ODBC\r\n-----------------\r\nData Source Name:FIA Biosum Oracle Services\r\nTNS Service Name:XE\r\nUser Id:fcs\r\nOracle Table Name:BIOSUM_VOLUME\r\nMS Access Table Link Name:fcs_biosum_volume\r\nConnection Status:";
                oDao.CreateOracleXETableLink("FIA Biosum Oracle Services", "fcs_biosum", "fcs", "FCS_BIOSUM", "BIOSUM_VOLUME", strFile, "fcs_biosum_volume");
                if (oDao.TableExists(strFile, "fcs_biosum_volume"))
                {
                    MessageBox.Show(str + "Success", "FIA Biosum");
                }
                else
                {
                    MessageBox.Show( str + "Error creating MS Access table link to FIADB01P Oracle table 'ANL_PNW_FIA_FCS.BIOSUM_VOLUME'", "FIA Biosum");
                }

                oDao.m_DaoDbEngine.Idle(1);
                oDao.m_DaoDbEngine.Idle(8);
                oDao.m_DaoWorkspace.Close();
                oDao.m_DaoDbEngine = null;
                oDao = null;
            }
            catch (Exception err)
            {
                frmMain.g_oFrmMain.DeactivateStandByAnimation();
                MessageBox.Show( str + "Error creating MS Access table link to Oracle XE table 'BIOSUM_VOLUME'\r\n Error Message:" + err.Message, "FIA Biosum");
            }
        }
        else
        {
            bool fcsTreeDbExists = System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase);
            bool bioSumCompsJarExists = System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\BioSumComps.JAR");
            bool fcsTreeCalcBatExists = System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_tree_calc.bat");
            bool bioSumCompsJarExecuted = false;

            //RUN here
            if (fcsTreeDbExists && bioSumCompsJarExists && fcsTreeCalcBatExists)
            {
                frmMain.g_oUtils.RunProcess(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum", "fcs_tree_calc.bat", "BAT");
                bioSumCompsJarExecuted = true;
            }
            else
            {
                m_intError = -1;
                m_strError = "Missing required files";
            }

            bool fcsErrorMsgExists = System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_error_msg.txt");
            if (fcsErrorMsgExists)
            {
                    m_strError = System.IO.File.ReadAllText(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_error_msg.txt");
                    if (m_strError.IndexOf("JAVA.EXE", 0) > 0)
                        m_strError = "Problem detected running JAVA.EXE";
                    m_intError = -2;
            }

            str = $@"FCS_TREE.DB & BioSumComps.jar test report
                {Environment.NewLine}-----------------------------------------
                {Environment.NewLine}User Id: Windows Authentication
                {Environment.NewLine}Table Name: BIOSUM_CALC
                {Environment.NewLine}FCS_TREE.DB FOUND: {(fcsTreeDbExists ? "Yes" : "No")}
                {Environment.NewLine}FCS_TREE_CALC.BAT FOUND: {(fcsTreeCalcBatExists ? "Yes" : "No")}
                {Environment.NewLine}BioSumComps.jar FOUND: {(bioSumCompsJarExists ? "Yes" : "No")}
                {Environment.NewLine}BioSumComps EXECUTED:{(bioSumCompsJarExecuted ? "Yes" : "No")}
                {Environment.NewLine}FCS_Errors FOUND: {(fcsErrorMsgExists ? "Yes" : "No")}";

            if (m_intError != 0)
                str += $"{Environment.NewLine}Error Message: {m_strError}";

            MessageBox.Show(str, "FIA Biosum");

        }

        frmMain.g_oFrmMain.DeactivateStandByAnimation();
        frmMain.g_oDelegate.CurrentThreadProcessDone = true;
        frmMain.g_oDelegate.m_oEventThreadStopped.Set();
        this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
    }

    private void txtStateCd_TextChanged(object sender, EventArgs e)
    {

    }

    private void label17_Click(object sender, EventArgs e)
    {

    }
  }
}

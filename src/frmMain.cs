using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;


namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class frmMain : System.Windows.Forms.Form
    {
        public MainMenu mnuMain;
		private System.Windows.Forms.MenuItem menuItem5;
		private System.Windows.Forms.MenuItem menuItem9;
		private System.Windows.Forms.MenuItem menuItem23;
		private System.Windows.Forms.MenuItem menuItem25;
		private System.Windows.Forms.MenuItem mnuFile;
		private System.Windows.Forms.MenuItem mnuFileOpenProject;
		private System.Windows.Forms.MenuItem mnuFileSaveProject;
		private System.Windows.Forms.MenuItem mnuFileExit;
		private System.Windows.Forms.MenuItem mnuFileNewProject;
		private System.Windows.Forms.MenuItem mnuFileRecentProjects;
		private System.Windows.Forms.MenuItem mnuFileRecentProjects1;
		private System.Windows.Forms.MenuItem mnuFileRecentProjects2;
		private System.Windows.Forms.MenuItem mnuFileRecentProjects3;
		private System.Windows.Forms.MenuItem mnuFileRecentProjects4;
		private System.Windows.Forms.MenuItem mnuView;
		private System.Windows.Forms.MenuItem mnuViewProject;
		private System.Windows.Forms.MenuItem mnuViewNotes;
		private System.Windows.Forms.MenuItem mnuViewLinks;
		private System.Windows.Forms.MenuItem mnuViewContacts;
		private System.Windows.Forms.MenuItem mnuHelp;
		private System.Windows.Forms.MenuItem mnuHelpBiosummatic;
		private System.Windows.Forms.MenuItem mnuHelpTechnicalSupport;
		private System.Windows.Forms.MenuItem mnuHelpAbout;
		private System.Windows.Forms.MenuItem mnuFileRecentProjects5;
		public System.Windows.Forms.ToolBar tlbMain;
		private System.Windows.Forms.ContextMenu ctxMenu1;
		private System.Windows.Forms.ImageList imgList1;
		private System.Windows.Forms.ToolBarButton btnOpen;
		private System.Windows.Forms.ToolBarButton btnProject;
		private System.Windows.Forms.ToolBarButton btnNotes;
		private System.Windows.Forms.ToolBarButton btnLinks;
		private System.Windows.Forms.ToolBarButton btnContacts;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;

		public FIA_Biosum_Manager.frmDialog frmProject;
		public System.Windows.Forms.GroupBox grpboxLeft;
		private System.Windows.Forms.Button btnDB;
		private System.Windows.Forms.Button btnProcessor;
		private System.Windows.Forms.Button btnFVS;
		private System.Windows.Forms.Button btnOptimizer;
		private System.ComponentModel.IContainer components;
		static int intGrpBoxLeftTopPosition = 0;
		static int intListHtPosition = 0;
		public System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        
		
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ImageList imgList2;
		private System.Windows.Forms.ToolTip toolTip1;
		public System.Windows.Forms.TextBox txtDropDown;
		
        public bool m_ProjectOpen=false ;
		public string[] m_LocalHardDrive;
		private System.Windows.Forms.ToolBarButton btnSave;
		public System.Windows.Forms.Panel panel1;
		
		//define the main form panels and buttons
		//DATABASE panel and buttons
		public System.Windows.Forms.Panel m_pnlDb;
		public FIA_Biosum_Manager.btnMainForm m_btnDbPlotData;
		public FIA_Biosum_Manager.btnMainForm m_btnDbTreeDiam;
		public FIA_Biosum_Manager.btnMainForm m_btnDbTreeSpGps;
		public FIA_Biosum_Manager.btnMainForm m_btnDbTableMgmt;
        public FIA_Biosum_Manager.btnMainForm m_btnSQLiteTableMgmt;
		public FIA_Biosum_Manager.btnMainForm m_btnDbRandomTravelTimes;
		public FIA_Biosum_Manager.btnMainForm m_btnDbDataSource;
		public FIA_Biosum_Manager.btnMainForm m_btnDbPSite;



		//CORE panel and buttons
        public System.Windows.Forms.Panel m_pnlOptimizer;
		public FIA_Biosum_Manager.btnMainForm m_btnOptimizerScenario;
		public FIA_Biosum_Manager.btnMainForm m_btnCoreMerge;
        public FIA_Biosum_Manager.btnMainForm m_btnOptimizerUserVariables;
        public FIA_Biosum_Manager.btnMainForm m_btnOptimizerLoadGisData;
        public FIA_Biosum_Manager.btnMainForm m_btnOptimizerSqlite;


		public System.Windows.Forms.Panel m_pnlFrcs;

		//FVS panel and buttons
		public System.Windows.Forms.Panel m_pnlFvs;
		public FIA_Biosum_Manager.btnMainForm m_btnFvsVariant;
		public FIA_Biosum_Manager.btnMainForm m_btnFvsRx;
		public FIA_Biosum_Manager.btnMainForm m_btnFvsRxPackage;
		public FIA_Biosum_Manager.btnMainForm m_btnFvsInput;
        //public FIA_Biosum_Manager.btnMainForm m_btnCreateFvsOutputMdbs;
        public FIA_Biosum_Manager.btnMainForm m_btnFvsOutput;
		public FIA_Biosum_Manager.btnMainForm m_btnFvsTreeSpcCvt;
		public FIA_Biosum_Manager.btnMainForm m_btnFvsTreeSpc;

		//PROCESSOR panel and buttons
		public System.Windows.Forms.Panel m_pnlProcessor;
        public FIA_Biosum_Manager.btnMainForm m_btnProcessorStart;
		//public FIA_Biosum_Manager.btnMainForm m_btnProcessorTreeDiam;
        //public FIA_Biosum_Manager.btnMainForm m_btnProcessorTreeSpcGrps;
        public FIA_Biosum_Manager.btnMainForm m_btnProcessorTreeSpc;
		//public FIA_Biosum_Manager.btnMainForm m_btnProcessorFrcs;

		

		public System.Windows.Forms.Panel m_pnlCurrent;
		private System.Windows.Forms.Button btnMain1;
		private env m_oEnv;
		
        

		private FIA_Biosum_Manager.frmDialog m_frmCoreMerge;      //core analysis merge scenarios form
        private FIA_Biosum_Manager.frmDialog m_frmCoreUserVariables;      //user pre/post variables form
		private FIA_Biosum_Manager.frmDialog m_frmPlotData;       //plot data form
		private FIA_Biosum_Manager.frmOptimizerScenario m_frmScenario;     //core analysis scenario form
		private FIA_Biosum_Manager.frmProcessorScenario m_frmProcessorScenario; //processor scenario form
		private FIA_Biosum_Manager.frmDialog m_frmTreeDiam;       //processor tree diameter form
		private FIA_Biosum_Manager.frmDialog m_frmSpcGrp;         //processor species group form
		private FIA_Biosum_Manager.frmDialog m_frmRx;             //treatment form 
		private FIA_Biosum_Manager.frmDialog m_frmRxPackage;      //treatment package form
		private FIA_Biosum_Manager.frmDialog m_frmDataSource;     //datasource form
		private FIA_Biosum_Manager.frmDialog m_frmFvsInput;       //fvs input form
		private FIA_Biosum_Manager.frmDialog m_frmFvsTreeSpcCvt;  //fvs tree spc conversion form
        private FIA_Biosum_Manager.frmDialog m_frmCreateFvsOutputMdbs; // popup form for fvs access DB creation
        private FIA_Biosum_Manager.frmDialog m_frmFvsOutput;      //fvs output
		private FIA_Biosum_Manager.frmDialog m_frmProcessorSpc;   //processor spc audit form
		private FIA_Biosum_Manager.frmDialog m_frmPSite;          //wood processing site form
		private FIA_Biosum_Manager.frmDialog m_frmFvsVariant;     //plot fvs variant
		private FIA_Biosum_Manager.frmDialog m_frmDb;                //database utilities
        private FIA_Biosum_Manager.frmDialog m_frmLoadGis;            //load gis data
        private FIA_Biosum_Manager.frmDialog m_frmDbSQLite;                //SQLite database utilities


        public const int TABLETYPE = 0;
		public const int PATH = 1;
		public const int MDBFILE = 2;
		public const int FILESTATUS = 3;
		public const int TABLE = 4;
		public const int TABLESTATUS = 5;
		public const int RECORDCOUNT = 6;

		public static string m_strProjectDriveLetter="c:";

		//status bar
		private System.Windows.Forms.StatusBarPanel sbpInfo;
		private System.Windows.Forms.StatusBarPanel sbpProgress;
		ProgressStatus ProgressStatus1;
		public static System.Windows.Forms.StatusBarPanel g_sbpInfo;
		public static System.Windows.Forms.StatusBarPanel g_sbpProgress;

		//main form
		public static FIA_Biosum_Manager.frmMain g_oFrmMain;
        public static System.Windows.Forms.Control g_oParentControl;

        public static FIA_Biosum_Manager.DelegateTools g_oDelegate;

		public static FIA_Biosum_Manager.Tables g_oTables = new Tables();

		public static FIA_Biosum_Manager.utils g_oUtils = new FIA_Biosum_Manager.utils();

		public static FIA_Biosum_Manager.env  g_oEnv = new env();

		public static System.Drawing.Color g_oGridViewBackgroundColor = Color.White;
		public static System.Drawing.Color g_oGridViewAlternateRowBackgroundColor=Color.LightGreen;
		public static System.Drawing.Color g_oGridViewRowBackgroundColor = Color.White;
		public static System.Drawing.Font  g_oGridViewFont= new System.Drawing.Font(
												"Microsoft Sans Serif", 8, System.Drawing.FontStyle.Regular);
		public static System.Drawing.Color g_oGridViewRowForegroundColor = Color.Black;
        public static System.Drawing.Color g_oGridViewSelectedRowBackgroundColor = Color.Magenta;

        //debugging values
        public static bool g_bDebug=true;
        public static int g_intDebugLevel = 3;
        

        //suppress table record counts
        public static bool g_bSuppressFVSInputTableRowCount = false;
        public static bool g_bSuppressFVSOutputTableRowCount = false;
        public static bool g_bSuppressProcessorScenarioTableRowCount = false;

        //OPCOST settings
        public static string g_strRDirectory = "";
        public static string g_strOPCOSTDirectory = "";

        //Path to Acrobat for Help
        public static string g_strAcrobatExe = "";

        //substitution macro variable
        public static FIA_Biosum_Manager.SQLMacroSubstitutionVariable_Collection g_oSQLMacroSubstitutionVariable_Collection= new SQLMacroSubstitutionVariable_Collection();
        public static FIA_Biosum_Manager.GeneralMacroSubstitutionVariable_Collection g_oGeneralMacroSubstitutionVariable_Collection = new GeneralMacroSubstitutionVariable_Collection();
        public const int PROJDIR = 0;
        public const int OLDPROJDIR = 1;

		public static string g_strAppVer = "5.11.2";
        public static string g_strBiosumDataDir = "\\FIABiosum";
        public static int g_intRefDbVer = 8;
        public static int g_intRefSqliteDbVer = 10;
        public static bool g_bUseOracleXE = false;

		private System.Windows.Forms.MenuItem mnuSettings;
        private MenuItem mnuTools;
        private MenuItem mnuToolsFCS;

		private bool m_bRefresh=false;
        private Help m_oHelp;
        //
        //splasher
        //
        public System.Threading.Thread standByAnimationThread;
        private MenuItem mnuToolsProjectRootFolder;
        public StandByAnimation.StandByAnimation standByAnimation;
        private MenuItem mnuReleaseNotes;
        private MenuItem mnuToolsPath;
        static readonly object _locker = new object();





        
               
       

		public frmMain()
		{
			//
			// Required for Windows Form Designer support
			//
           
			InitializeComponent();


            

            this.m_oEnv = new env();

            

            //create and initialize project form
            this.frmProject = new FIA_Biosum_Manager.frmDialog();
            
            this.frmProject.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.frmProject.BackColor = System.Drawing.SystemColors.Control;
			this.frmProject.ClientSize = new System.Drawing.Size(712, 488);
			this.frmProject.Location = new System.Drawing.Point(23, 23);
			this.frmProject.Name = "frmProject";

           

            this.frmProject.MdiParent = this;
			this.frmProject.Initialize_User_Control("PROJECT");
			this.frmProject.Visible=false;

            
            intGrpBoxLeftTopPosition = 0; //Math.Abs(this.btnDB.Top + this.btnDB.Height + 2) ;
			for (int x=1;;x++)
			{
                intGrpBoxLeftTopPosition=x;
				if (this.btnDB.Top + this.btnDB.Height + 2 < x)
				{
					break;
				}

			}
            

            intListHtPosition = this.Height - (this.btnDB.Height * 2);
			
			this.panel1.Top = intGrpBoxLeftTopPosition;
			this.panel1.BackColor = System.Drawing.Color.Gray;
			this.btnMain1.Visible=false;
			this.panel1.Enabled=false;
			this.panel1.Visible=false;

            

            this.InitializeMainFormPanelsAndButtons();
			this.panel1.Height=0;
			
			panel1.Height = this.grpboxLeft.Height - panel1.Top;

            
            this.m_pnlCurrent = this.m_pnlDb;
			this.m_pnlCurrent.Enabled=false;
			this.m_pnlCurrent.Size = this.panel1.Size;
			this.m_pnlCurrent.Height = this.panel1.Height ;
			this.m_pnlCurrent.Top = intGrpBoxLeftTopPosition;
			this.m_pnlCurrent.Visible=true;

            

            this.btnSave.Enabled=false;
			this.btnContacts.Enabled=false;
			this.btnOptimizer.Enabled = false;
			this.btnDB.Enabled=false;
			
			this.btnFVS.Enabled=false;
			this.btnProcessor.Enabled=false;
			this.panel1.Enabled=false;
			this.btnContacts.Enabled=false;
			this.btnNotes.Enabled=false;
			this.btnProject.Enabled=false;
			this.btnLinks.Enabled = false;
			this.mnuView.Enabled=false;
			
			this.mnuFileSaveProject.Enabled=false;
			

			/********************************************************************************
			 **create the user fia biosum application data directory if it does not exist
             ********************************************************************************/
			//make sure \documents and settings\ + userid + \application data\qatools exists
			string str = m_oEnv.strApplicationDataDirectory	+ "\\FIABiosum";
			if (!System.IO.Directory.Exists(str))
			{
				System.IO.Directory.CreateDirectory(str);
			}
			//********************************************************************************
            //load the menu with a list of the 5 most recent projects opened
            //********************************************************************************
			if (System.IO.File.Exists(this.m_oEnv.strApplicationDataDirectory + "\\recent.dat"))
			{
                this.mnuFileRecentProjects.Enabled=true;
				this.mnuFileRecentProjects1.Visible=false;
				this.mnuFileRecentProjects2.Visible=false;
				this.mnuFileRecentProjects3.Visible=false;
				this.mnuFileRecentProjects4.Visible=false;
				this.mnuFileRecentProjects5.Visible=false;
				try 
				{
					// Create an instance of StreamReader to read from a file.
					// The using statement also closes the StreamReader.
					System.IO.StreamReader sr = new System.IO.StreamReader(this.m_oEnv.strApplicationDataDirectory + "\\recent.dat");
				
					String line;
					int intMenuCount=0;
					// Read and display lines from the file until the end of 
					// the file is reached.
					while ((line = sr.ReadLine()) != null) 
					{
						switch (intMenuCount)
						{
							case 0:
								this.mnuFileRecentProjects1.Visible=true;
                                this.mnuFileRecentProjects1.Text = line;
								break;
							case 1:
								this.mnuFileRecentProjects2.Visible=true;
								this.mnuFileRecentProjects2.Text = line;
								break;
							case 2:
								this.mnuFileRecentProjects3.Visible=true;
								this.mnuFileRecentProjects3.Text = line;
								break;
							case 3:
								this.mnuFileRecentProjects4.Visible=true;
								this.mnuFileRecentProjects4.Text = line;
								break;
							case 4:
								this.mnuFileRecentProjects5.Visible=true;
								this.mnuFileRecentProjects5.Text = line;
								break;
						}
						intMenuCount++;
					}
				    sr.Close();
					sr = null;
				}
				catch  
				{
				}
			}
            //get local hard drives on the pc
            this.m_LocalHardDrive = new string[24];
			FIA_Biosum_Manager.utils p_oUtils = new FIA_Biosum_Manager.utils();
			this.m_LocalHardDrive = p_oUtils.getLocalHardDriveList();
			p_oUtils = null;

			if (this.frmProject.Width > this.Width - this.grpboxLeft.Width)
			{
				this.frmProject.Width = this.Width - this.grpboxLeft.Width - 20;

			}

			frmMain.g_oDelegate = new DelegateTools();

            
            this.btnDB.EnabledChanged += new System.EventHandler(this.ProcessButton_EnabledChanged);
			this.btnDB.Paint += new System.Windows.Forms.PaintEventHandler(this.ProcessButton_Paint);
			this.btnFVS.EnabledChanged += new System.EventHandler(this.ProcessButton_EnabledChanged);
			this.btnFVS.Paint += new System.Windows.Forms.PaintEventHandler(this.ProcessButton_Paint);
			this.btnProcessor.EnabledChanged += new System.EventHandler(this.ProcessButton_EnabledChanged);
			this.btnProcessor.Paint += new System.Windows.Forms.PaintEventHandler(this.ProcessButton_Paint);
			this.btnOptimizer.EnabledChanged += new System.EventHandler(this.ProcessButton_EnabledChanged);
			this.btnOptimizer.Paint += new System.Windows.Forms.PaintEventHandler(this.ProcessButton_Paint);

			btnDB.ForeColor = System.Drawing.SystemColors.GrayText;
			btnFVS.ForeColor = System.Drawing.SystemColors.GrayText;
			btnProcessor.ForeColor = System.Drawing.SystemColors.GrayText;
			btnOptimizer.ForeColor = System.Drawing.SystemColors.GrayText;

			Datasource.InititializeMacroVariables();

            GeneralMacroSubstitutionVariableItem oItem = new GeneralMacroSubstitutionVariableItem();
            oItem.Description = "Current Project Directory";
            oItem.Index = 0;
            oItem.VariableName = "ProjDir";
            oItem.VariableSubstitutionString = "";
            g_oGeneralMacroSubstitutionVariable_Collection.Add(oItem);

            oItem = new GeneralMacroSubstitutionVariableItem();
            oItem.Description = "Old Project Directory";
            oItem.Index = 1;
            oItem.VariableName = "OldProjDir";
            oItem.VariableSubstitutionString = "";
            g_oGeneralMacroSubstitutionVariable_Collection.Add(oItem);

            oItem = new GeneralMacroSubstitutionVariableItem();
            oItem.Description = "Application Data Folder";
            oItem.Index = 2;
            oItem.VariableName = "AppData";
            oItem.VariableSubstitutionString = frmMain.g_oEnv.strApplicationDataDirectory;
            g_oGeneralMacroSubstitutionVariable_Collection.Add(oItem);

            CheckForBiosumRefData();
            CheckForBiosumRefSqliteData();
            CheckForFcsFiles();
            
        }

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// 
		~frmMain()
		{
			
			
		}
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.mnuMain = new System.Windows.Forms.MainMenu(this.components);
            this.mnuFile = new System.Windows.Forms.MenuItem();
            this.mnuFileNewProject = new System.Windows.Forms.MenuItem();
            this.mnuFileOpenProject = new System.Windows.Forms.MenuItem();
            this.mnuFileSaveProject = new System.Windows.Forms.MenuItem();
            this.menuItem5 = new System.Windows.Forms.MenuItem();
            this.mnuFileRecentProjects = new System.Windows.Forms.MenuItem();
            this.mnuFileRecentProjects1 = new System.Windows.Forms.MenuItem();
            this.mnuFileRecentProjects2 = new System.Windows.Forms.MenuItem();
            this.mnuFileRecentProjects3 = new System.Windows.Forms.MenuItem();
            this.mnuFileRecentProjects4 = new System.Windows.Forms.MenuItem();
            this.mnuFileRecentProjects5 = new System.Windows.Forms.MenuItem();
            this.menuItem9 = new System.Windows.Forms.MenuItem();
            this.mnuFileExit = new System.Windows.Forms.MenuItem();
            this.mnuView = new System.Windows.Forms.MenuItem();
            this.mnuViewProject = new System.Windows.Forms.MenuItem();
            this.mnuViewNotes = new System.Windows.Forms.MenuItem();
            this.mnuViewLinks = new System.Windows.Forms.MenuItem();
            this.mnuViewContacts = new System.Windows.Forms.MenuItem();
            this.mnuSettings = new System.Windows.Forms.MenuItem();
            this.mnuTools = new System.Windows.Forms.MenuItem();
            this.mnuToolsFCS = new System.Windows.Forms.MenuItem();
            this.mnuToolsProjectRootFolder = new System.Windows.Forms.MenuItem();
            this.mnuToolsPath = new System.Windows.Forms.MenuItem();
            this.mnuHelp = new System.Windows.Forms.MenuItem();
            this.mnuHelpBiosummatic = new System.Windows.Forms.MenuItem();
            this.menuItem23 = new System.Windows.Forms.MenuItem();
            this.mnuHelpTechnicalSupport = new System.Windows.Forms.MenuItem();
            this.menuItem25 = new System.Windows.Forms.MenuItem();
            this.mnuReleaseNotes = new System.Windows.Forms.MenuItem();
            this.mnuHelpAbout = new System.Windows.Forms.MenuItem();
            this.tlbMain = new System.Windows.Forms.ToolBar();
            this.btnOpen = new System.Windows.Forms.ToolBarButton();
            this.btnSave = new System.Windows.Forms.ToolBarButton();
            this.btnProject = new System.Windows.Forms.ToolBarButton();
            this.btnNotes = new System.Windows.Forms.ToolBarButton();
            this.btnLinks = new System.Windows.Forms.ToolBarButton();
            this.btnContacts = new System.Windows.Forms.ToolBarButton();
            this.ctxMenu1 = new System.Windows.Forms.ContextMenu();
            this.imgList1 = new System.Windows.Forms.ImageList(this.components);
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.grpboxLeft = new System.Windows.Forms.GroupBox();
            this.btnOptimizer = new System.Windows.Forms.Button();
            this.btnFVS = new System.Windows.Forms.Button();
            this.btnProcessor = new System.Windows.Forms.Button();
            this.btnDB = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnMain1 = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.imgList2 = new System.Windows.Forms.ImageList(this.components);
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.txtDropDown = new System.Windows.Forms.TextBox();
            this.grpboxLeft.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // mnuMain
            // 
            this.mnuMain.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuFile,
            this.mnuView,
            this.mnuSettings,
            this.mnuTools,
            this.mnuHelp});
            // 
            // mnuFile
            // 
            this.mnuFile.Index = 0;
            this.mnuFile.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuFileNewProject,
            this.mnuFileOpenProject,
            this.mnuFileSaveProject,
            this.menuItem5,
            this.mnuFileRecentProjects,
            this.menuItem9,
            this.mnuFileExit});
            this.mnuFile.Text = "&File";
            // 
            // mnuFileNewProject
            // 
            this.mnuFileNewProject.Index = 0;
            this.mnuFileNewProject.Text = "New Project";
            this.mnuFileNewProject.Click += new System.EventHandler(this.mnuFileNewProject_Click);
            // 
            // mnuFileOpenProject
            // 
            this.mnuFileOpenProject.Index = 1;
            this.mnuFileOpenProject.Text = "&Open Project";
            this.mnuFileOpenProject.Click += new System.EventHandler(this.mnuFileOpenProject_Click);
            // 
            // mnuFileSaveProject
            // 
            this.mnuFileSaveProject.Index = 2;
            this.mnuFileSaveProject.Text = "&Save";
            // 
            // menuItem5
            // 
            this.menuItem5.Index = 3;
            this.menuItem5.Text = "-";
            // 
            // mnuFileRecentProjects
            // 
            this.mnuFileRecentProjects.Enabled = false;
            this.mnuFileRecentProjects.Index = 4;
            this.mnuFileRecentProjects.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuFileRecentProjects1,
            this.mnuFileRecentProjects2,
            this.mnuFileRecentProjects3,
            this.mnuFileRecentProjects4,
            this.mnuFileRecentProjects5});
            this.mnuFileRecentProjects.Text = "Recent Projects";
            // 
            // mnuFileRecentProjects1
            // 
            this.mnuFileRecentProjects1.Index = 0;
            this.mnuFileRecentProjects1.Text = "";
            this.mnuFileRecentProjects1.Click += new System.EventHandler(this.mnuFileRecentProjects1_Click);
            // 
            // mnuFileRecentProjects2
            // 
            this.mnuFileRecentProjects2.Index = 1;
            this.mnuFileRecentProjects2.Text = "";
            this.mnuFileRecentProjects2.Click += new System.EventHandler(this.mnuFileRecentProjects2_Click);
            // 
            // mnuFileRecentProjects3
            // 
            this.mnuFileRecentProjects3.Index = 2;
            this.mnuFileRecentProjects3.Text = "";
            this.mnuFileRecentProjects3.Click += new System.EventHandler(this.mnuFileRecentProjects3_Click);
            // 
            // mnuFileRecentProjects4
            // 
            this.mnuFileRecentProjects4.Index = 3;
            this.mnuFileRecentProjects4.Text = "";
            this.mnuFileRecentProjects4.Click += new System.EventHandler(this.mnuFileRecentProjects4_Click);
            // 
            // mnuFileRecentProjects5
            // 
            this.mnuFileRecentProjects5.Index = 4;
            this.mnuFileRecentProjects5.Text = "";
            this.mnuFileRecentProjects5.Click += new System.EventHandler(this.mnuFileRecentProjects5_Click);
            // 
            // menuItem9
            // 
            this.menuItem9.Index = 5;
            this.menuItem9.Text = "-";
            // 
            // mnuFileExit
            // 
            this.mnuFileExit.Index = 6;
            this.mnuFileExit.Text = "E&xit";
            this.mnuFileExit.Click += new System.EventHandler(this.mnuFileExit_Click);
            // 
            // mnuView
            // 
            this.mnuView.Index = 1;
            this.mnuView.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuViewProject,
            this.mnuViewNotes,
            this.mnuViewLinks,
            this.mnuViewContacts});
            this.mnuView.Text = "&View";
            // 
            // mnuViewProject
            // 
            this.mnuViewProject.Index = 0;
            this.mnuViewProject.Text = "Project Properties";
            this.mnuViewProject.Click += new System.EventHandler(this.mnuViewProject_Click);
            // 
            // mnuViewNotes
            // 
            this.mnuViewNotes.Index = 1;
            this.mnuViewNotes.Text = "Project Notes";
            this.mnuViewNotes.Click += new System.EventHandler(this.mnuViewNotes_Click);
            // 
            // mnuViewLinks
            // 
            this.mnuViewLinks.Index = 2;
            this.mnuViewLinks.Text = "Project Links";
            this.mnuViewLinks.Click += new System.EventHandler(this.mnuViewLinks_Click);
            // 
            // mnuViewContacts
            // 
            this.mnuViewContacts.Index = 3;
            this.mnuViewContacts.Text = "Project Contacts";
            this.mnuViewContacts.Click += new System.EventHandler(this.mnuViewContacts_Click);
            // 
            // mnuSettings
            // 
            this.mnuSettings.Index = 2;
            this.mnuSettings.Text = "Settings";
            this.mnuSettings.Click += new System.EventHandler(this.mnuSettings_Click);
            // 
            // mnuTools
            // 
            this.mnuTools.Index = 3;
            this.mnuTools.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuToolsFCS,
            this.mnuToolsProjectRootFolder,
            this.mnuToolsPath});
            this.mnuTools.Text = "Tools";
            // 
            // mnuToolsFCS
            // 
            this.mnuToolsFCS.Index = 0;
            this.mnuToolsFCS.Text = "Tree Volume and Biomass Calculator Troubleshooter Tool";
            this.mnuToolsFCS.Click += new System.EventHandler(this.mnuToolsFCS_Click);
            // 
            // mnuToolsProjectRootFolder
            // 
            this.mnuToolsProjectRootFolder.Enabled = false;
            this.mnuToolsProjectRootFolder.Index = 1;
            this.mnuToolsProjectRootFolder.Text = "Scan and Synchronize Project Root Folder Tool";
            this.mnuToolsProjectRootFolder.Click += new System.EventHandler(this.mnuToolsProjectRootFolder_Click);
            // 
            // mnuToolsPath
            // 
            this.mnuToolsPath.Index = 2;
            this.mnuToolsPath.Text = "Path Environment Variable Value";
            this.mnuToolsPath.Click += new System.EventHandler(this.mnuToolsPath_Click);
            // 
            // mnuHelp
            // 
            this.mnuHelp.Index = 4;
            this.mnuHelp.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.mnuHelpBiosummatic,
            this.menuItem23,
            this.mnuHelpTechnicalSupport,
            this.menuItem25,
            this.mnuReleaseNotes,
            this.mnuHelpAbout});
            this.mnuHelp.Text = "&Help";
            // 
            // mnuHelpBiosummatic
            // 
            this.mnuHelpBiosummatic.Index = 0;
            this.mnuHelpBiosummatic.Text = "FIA Biosum";
            // 
            // menuItem23
            // 
            this.menuItem23.Index = 1;
            this.menuItem23.Text = "-";
            // 
            // mnuHelpTechnicalSupport
            // 
            this.mnuHelpTechnicalSupport.Index = 2;
            this.mnuHelpTechnicalSupport.Text = "Product Support";
            this.mnuHelpTechnicalSupport.Click += new System.EventHandler(this.mnuHelpTechnicalSupport_Click);
            // 
            // menuItem25
            // 
            this.menuItem25.Index = 3;
            this.menuItem25.Text = "-";
            // 
            // mnuReleaseNotes
            // 
            this.mnuReleaseNotes.Index = 4;
            this.mnuReleaseNotes.Text = "Release Notes";
            this.mnuReleaseNotes.Click += new System.EventHandler(this.mnuReleaseNotes_Click);
            // 
            // mnuHelpAbout
            // 
            this.mnuHelpAbout.Index = 5;
            this.mnuHelpAbout.Text = "About FIA Biosum";
            this.mnuHelpAbout.Click += new System.EventHandler(this.mnuHelpAbout_Click);
            // 
            // tlbMain
            // 
            this.tlbMain.AutoSize = false;
            this.tlbMain.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
            this.btnOpen,
            this.btnSave,
            this.btnProject,
            this.btnNotes,
            this.btnLinks,
            this.btnContacts});
            this.tlbMain.ButtonSize = new System.Drawing.Size(52, 38);
            this.tlbMain.ContextMenu = this.ctxMenu1;
            this.tlbMain.DropDownArrows = true;
            this.tlbMain.ImageList = this.imgList1;
            this.tlbMain.Location = new System.Drawing.Point(0, 0);
            this.tlbMain.Margin = new System.Windows.Forms.Padding(4);
            this.tlbMain.Name = "tlbMain";
            this.tlbMain.ShowToolTips = true;
            this.tlbMain.Size = new System.Drawing.Size(981, 49);
            this.tlbMain.TabIndex = 1;
            this.tlbMain.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.tlbMain_ButtonClick);
            this.tlbMain.Click += new System.EventHandler(this.tlbMain_Click);
            // 
            // btnOpen
            // 
            this.btnOpen.ImageIndex = 0;
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Text = "Open";
            // 
            // btnSave
            // 
            this.btnSave.ImageIndex = 1;
            this.btnSave.Name = "btnSave";
            this.btnSave.Text = "Save";
            // 
            // btnProject
            // 
            this.btnProject.ImageIndex = 2;
            this.btnProject.Name = "btnProject";
            this.btnProject.Text = "Project";
            // 
            // btnNotes
            // 
            this.btnNotes.ImageIndex = 3;
            this.btnNotes.Name = "btnNotes";
            this.btnNotes.Text = "Notes";
            // 
            // btnLinks
            // 
            this.btnLinks.ImageIndex = 4;
            this.btnLinks.Name = "btnLinks";
            this.btnLinks.Text = "Links";
            // 
            // btnContacts
            // 
            this.btnContacts.ImageIndex = 5;
            this.btnContacts.Name = "btnContacts";
            this.btnContacts.Text = "Contacts";
            // 
            // ctxMenu1
            // 
            this.ctxMenu1.Popup += new System.EventHandler(this.contextMenu1_Popup);
            // 
            // imgList1
            // 
            this.imgList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgList1.ImageStream")));
            this.imgList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imgList1.Images.SetKeyName(0, "");
            this.imgList1.Images.SetKeyName(1, "");
            this.imgList1.Images.SetKeyName(2, "");
            this.imgList1.Images.SetKeyName(3, "");
            this.imgList1.Images.SetKeyName(4, "");
            this.imgList1.Images.SetKeyName(5, "");
            // 
            // grpboxLeft
            // 
            this.grpboxLeft.Controls.Add(this.btnOptimizer);
            this.grpboxLeft.Controls.Add(this.btnFVS);
            this.grpboxLeft.Controls.Add(this.btnProcessor);
            this.grpboxLeft.Controls.Add(this.btnDB);
            this.grpboxLeft.Controls.Add(this.panel1);
            this.grpboxLeft.Dock = System.Windows.Forms.DockStyle.Left;
            this.grpboxLeft.Location = new System.Drawing.Point(0, 49);
            this.grpboxLeft.Margin = new System.Windows.Forms.Padding(4);
            this.grpboxLeft.Name = "grpboxLeft";
            this.grpboxLeft.Padding = new System.Windows.Forms.Padding(4);
            this.grpboxLeft.Size = new System.Drawing.Size(180, 577);
            this.grpboxLeft.TabIndex = 7;
            this.grpboxLeft.TabStop = false;
            this.grpboxLeft.Resize += new System.EventHandler(this.grpboxLeft_Resize);
            // 
            // btnOptimizer
            // 
            this.btnOptimizer.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.btnOptimizer.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOptimizer.Location = new System.Drawing.Point(4, 465);
            this.btnOptimizer.Margin = new System.Windows.Forms.Padding(4);
            this.btnOptimizer.Name = "btnOptimizer";
            this.btnOptimizer.Size = new System.Drawing.Size(172, 34);
            this.btnOptimizer.TabIndex = 5;
            this.btnOptimizer.Text = "Optimizer";
            this.btnOptimizer.Click += new System.EventHandler(this.btnCoreAnalysis_Click);
            // 
            // btnFVS
            // 
            this.btnFVS.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.btnFVS.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVS.Location = new System.Drawing.Point(4, 499);
            this.btnFVS.Margin = new System.Windows.Forms.Padding(4);
            this.btnFVS.Name = "btnFVS";
            this.btnFVS.Size = new System.Drawing.Size(172, 38);
            this.btnFVS.TabIndex = 3;
            this.btnFVS.Text = "FVS";
            this.btnFVS.Click += new System.EventHandler(this.btnFVS_Click);
            // 
            // btnProcessor
            // 
            this.btnProcessor.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.btnProcessor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnProcessor.Location = new System.Drawing.Point(4, 537);
            this.btnProcessor.Margin = new System.Windows.Forms.Padding(4);
            this.btnProcessor.Name = "btnProcessor";
            this.btnProcessor.Size = new System.Drawing.Size(172, 36);
            this.btnProcessor.TabIndex = 2;
            this.btnProcessor.Text = "Processor";
            this.btnProcessor.Click += new System.EventHandler(this.btnProcessor_Click);
            // 
            // btnDB
            // 
            this.btnDB.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnDB.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDB.Location = new System.Drawing.Point(4, 19);
            this.btnDB.Margin = new System.Windows.Forms.Padding(4);
            this.btnDB.Name = "btnDB";
            this.btnDB.Size = new System.Drawing.Size(172, 34);
            this.btnDB.TabIndex = 0;
            this.btnDB.Text = "Database";
            this.btnDB.Click += new System.EventHandler(this.btnDB_Click);
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.btnMain1);
            this.panel1.Location = new System.Drawing.Point(12, 72);
            this.panel1.Margin = new System.Windows.Forms.Padding(4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(150, 348);
            this.panel1.TabIndex = 13;
            // 
            // btnMain1
            // 
            this.btnMain1.Font = new System.Drawing.Font("Times New Roman", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMain1.Location = new System.Drawing.Point(18, 12);
            this.btnMain1.Margin = new System.Windows.Forms.Padding(4);
            this.btnMain1.Name = "btnMain1";
            this.btnMain1.Size = new System.Drawing.Size(112, 84);
            this.btnMain1.TabIndex = 13;
            this.btnMain1.MouseEnter += new System.EventHandler(this.btnMain1_MouseEnter);
            this.btnMain1.MouseLeave += new System.EventHandler(this.btnMain1_MouseLeave);
            // 
            // imgList2
            // 
            this.imgList2.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgList2.ImageStream")));
            this.imgList2.TransparentColor = System.Drawing.Color.Transparent;
            this.imgList2.Images.SetKeyName(0, "");
            // 
            // imageList1
            // 
            this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // txtDropDown
            // 
            this.txtDropDown.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.txtDropDown.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDropDown.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtDropDown.HideSelection = false;
            this.txtDropDown.Location = new System.Drawing.Point(288, 132);
            this.txtDropDown.Margin = new System.Windows.Forms.Padding(4);
            this.txtDropDown.Multiline = true;
            this.txtDropDown.Name = "txtDropDown";
            this.txtDropDown.ReadOnly = true;
            this.txtDropDown.Size = new System.Drawing.Size(226, 34);
            this.txtDropDown.TabIndex = 11;
            this.txtDropDown.Visible = false;
            // 
            // frmMain
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(981, 626);
            this.Controls.Add(this.grpboxLeft);
            this.Controls.Add(this.tlbMain);
            this.Controls.Add(this.txtDropDown);
            this.Icon = ((System.Drawing.Icon) Properties.Resources.favicon);
            this.IsMdiContainer = true;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Menu = this.mnuMain;
            this.Name = "frmMain";
            this.Text = "FIA Biosum Manager";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Closing += new System.ComponentModel.CancelEventHandler(this.frmMain_Closing);
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.Resize += new System.EventHandler(this.frmMain_Resize);
            this.grpboxLeft.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new frmMain());
			
		}

		private void contextMenu1_Popup(object sender, System.EventArgs e)
		{
		
		}

		private void tlbMain_Click(object sender, System.EventArgs e)
		{
		
		}

		private void mnuFileExit_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void mnuFileOpenProject_Click(object sender,System.EventArgs e)
		{
			
		    this.frmProject.uc_project1.Open_Project();
			if (this.frmProject.uc_project1.m_intError == 0)
			{
         		this.OpenProject(this.frmProject.uc_project1.m_strNewProjectDirectory, this.frmProject.uc_project1.m_strNewProjectFile);

			}

		}
		
		private void frmMain_Resize(object sender, System.EventArgs e)
		{
			if  (sbpInfo != null)
			this.sbpInfo.Width=(this.Width / 2);
			
        
		}

		

		private void frmMain_Load(object sender, System.EventArgs e)
		{
			if (System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg"))
            {
                string strSection="";
                
				string strColorARGBList="";
				string[] strColorARGBArray=null;
				string strFontList="";
				string[] strFontArray=null;
				//Open the file in a stream reader.
				System.IO.StreamReader s = new System.IO.StreamReader(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg");
				//Read the rest of the data in the file.        
				string strWholeFile = s.ReadToEnd();
				string[] rows = strWholeFile.Split("\r\n".ToCharArray());
				foreach(string r in rows)
				{
					if (r.Trim().Length > 0)
					{
					    
                        
						if (r.Trim().ToUpper()=="END") break;
                        if (r.IndexOf("[", 0) >= 0 && r.IndexOf("]", 0) >= 0)
                            strSection = r.Trim().ToUpper();
						int intPos = r.IndexOf("=",0);
						if (intPos > 0)
						{
							string[] strSettingsArray = frmMain.g_oUtils.ConvertListToArray(r,"=");
                            if (strSection == "[GRIDVIEW]")
                            {
                                switch (strSettingsArray[0].Trim())
                                {
                                    case "BackgroundColor_A":
                                        strColorARGBList = strSettingsArray[1] + ",";
                                        break;
                                    case "BackgroundColor_R":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1] + ",";
                                        break;
                                    case "BackgroundColor_G":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1] + ",";
                                        break;
                                    case "BackgroundColor_B":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1];
                                        strColorARGBArray = frmMain.g_oUtils.ConvertListToArray(strColorARGBList, ",");
                                        frmMain.g_oGridViewBackgroundColor = Color.FromArgb(Convert.ToInt32(strColorARGBArray[0]),
                                            Convert.ToInt32(strColorARGBArray[1]),
                                            Convert.ToInt32(strColorARGBArray[2]),
                                            Convert.ToInt32(strColorARGBArray[3]));
                                        strColorARGBList = "";
                                        strColorARGBArray = null;
                                        break;
                                    case "RowBackgroundColor_A":
                                        strColorARGBList = strSettingsArray[1] + ",";
                                        break;
                                    case "RowBackgroundColor_R":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1] + ",";
                                        break;
                                    case "RowBackgroundColor_G":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1] + ",";
                                        break;
                                    case "RowBackgroundColor_B":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1];
                                        strColorARGBArray = frmMain.g_oUtils.ConvertListToArray(strColorARGBList, ",");
                                        frmMain.g_oGridViewRowBackgroundColor = Color.FromArgb(Convert.ToInt32(strColorARGBArray[0]),
                                                                                            Convert.ToInt32(strColorARGBArray[1]),
                                                                                            Convert.ToInt32(strColorARGBArray[2]),
                                                                                            Convert.ToInt32(strColorARGBArray[3]));
                                        strColorARGBList = "";
                                        strColorARGBArray = null;
                                        break;
                                    case "AlternatingRowBackgroundColor_A":
                                        strColorARGBList = strSettingsArray[1] + ",";
                                        break;
                                    case "AlternatingRowBackgroundColor_R":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1] + ",";
                                        break;
                                    case "AlternatingRowBackgroundColor_G":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1] + ",";
                                        break;
                                    case "AlternatingRowBackgroundColor_B":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1];
                                        strColorARGBArray = frmMain.g_oUtils.ConvertListToArray(strColorARGBList, ",");
                                        frmMain.g_oGridViewAlternateRowBackgroundColor = Color.FromArgb(Convert.ToInt32(strColorARGBArray[0]),
                                            Convert.ToInt32(strColorARGBArray[1]),
                                            Convert.ToInt32(strColorARGBArray[2]),
                                            Convert.ToInt32(strColorARGBArray[3]));
                                        strColorARGBList = "";
                                        strColorARGBArray = null;
                                        break;
                                    case "RowForegroundColor_A":
                                        strColorARGBList = strSettingsArray[1] + ",";
                                        break;
                                    case "RowForegroundColor_R":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1] + ",";
                                        break;
                                    case "RowForegroundColor_G":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1] + ",";
                                        break;
                                    case "RowForegroundColor_B":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1];
                                        strColorARGBArray = frmMain.g_oUtils.ConvertListToArray(strColorARGBList, ",");
                                        frmMain.g_oGridViewRowForegroundColor = Color.FromArgb(Convert.ToInt32(strColorARGBArray[0]),
                                            Convert.ToInt32(strColorARGBArray[1]),
                                            Convert.ToInt32(strColorARGBArray[2]),
                                            Convert.ToInt32(strColorARGBArray[3]));
                                        strColorARGBList = "";
                                        strColorARGBArray = null;
                                        break;
                                    case "SelectedRowBackgroundColor_A":
                                        strColorARGBList = strSettingsArray[1] + ",";
                                        break;
                                    case "SelectedRowBackgroundColor_R":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1] + ",";
                                        break;
                                    case "SelectedRowBackgroundColor_G":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1] + ",";
                                        break;
                                    case "SelectedRowBackgroundColor_B":
                                        strColorARGBList = strColorARGBList + strSettingsArray[1];
                                        strColorARGBArray = frmMain.g_oUtils.ConvertListToArray(strColorARGBList, ",");
                                        frmMain.g_oGridViewSelectedRowBackgroundColor = Color.FromArgb(Convert.ToInt32(strColorARGBArray[0]),
                                            Convert.ToInt32(strColorARGBArray[1]),
                                            Convert.ToInt32(strColorARGBArray[2]),
                                            Convert.ToInt32(strColorARGBArray[3]));
                                        strColorARGBList = "";
                                        strColorARGBArray = null;
                                        break;
                                    case "FontName":
                                        strFontList = strSettingsArray[1] + ",";
                                        break;
                                    case "FontSize":
                                        strFontList = strFontList + strSettingsArray[1] + ",";
                                        break;
                                    case "FontStyle":
                                        strFontList = strFontList + strSettingsArray[1];
                                        strFontArray = frmMain.g_oUtils.ConvertListToArray(strFontList, ",");
                                        frmMain.g_oGridViewFont = new Font(strFontArray[0].ToString(),
                                                                           Convert.ToInt16(strFontArray[1]),
                                                                           (System.Drawing.FontStyle)Convert.ToInt16(strFontArray[2]));
                                        break;

                                }
                            }
                            else if (strSection == "[DEBUG]")
                            {
                                switch (strSettingsArray[0].Trim())
                                {
                                    case "Debug":
                                        if (strSettingsArray[1].Trim().ToUpper() == "Y") frmMain.g_bDebug = true;
                                        else frmMain.g_bDebug = false;
                                        break;
                                    case "Level":
                                        if (strSettingsArray[1].Trim() == "1") frmMain.g_intDebugLevel = Convert.ToInt32(strSettingsArray[1]);
                                        else if (strSettingsArray[1].Trim() == "2") frmMain.g_intDebugLevel = Convert.ToInt32(strSettingsArray[1]);
                                        else if (strSettingsArray[1].Trim() == "3") frmMain.g_intDebugLevel = Convert.ToInt32(strSettingsArray[1]);
                                        else
                                            frmMain.g_intDebugLevel = 1;

                                        break;
                                }
                            }
                            else if (strSection == "[SUPPRESS TABLE RECORD COUNTS]")
                            {
                                switch (strSettingsArray[0].Trim())
                                {
                                    case "FVSInputForm":
                                        if (strSettingsArray[1].Trim().ToUpper() == "Y") frmMain.g_bSuppressFVSInputTableRowCount = true;
                                        else frmMain.g_bSuppressFVSInputTableRowCount = false;
                                        break;
                                    case "FVSOutputForm":
                                        if (strSettingsArray[1].Trim().ToUpper() == "Y") frmMain.g_bSuppressFVSOutputTableRowCount = true;
                                        else frmMain.g_bSuppressFVSOutputTableRowCount = false;
                                        break;
                                    case "ProcessorScenarioForm":
                                        if (strSettingsArray[1].Trim().ToUpper() == "Y") frmMain.g_bSuppressProcessorScenarioTableRowCount = true;
                                        else frmMain.g_bSuppressProcessorScenarioTableRowCount = false;
                                        break;
                                   
                                }
                            }
                            else if (strSection == "[OPCOST]")
                            {
                                switch (strSettingsArray[0].Trim())
                                {
                                    case "RFile":
                                        frmMain.g_strRDirectory = strSettingsArray[1].Trim();
                                        break;
                                    case "OPCOSTFile":
                                        frmMain.g_strOPCOSTDirectory = strSettingsArray[1].Trim();
                                        break;
                                 
                                }
                            }

						}

					}
				}
				s.Close();
				s=null;
			}

            // Search the registry for the Adobe executable
            const string ACROBAT_REG_PATH = "SOFTWARE\\CLASSES\\ACROBAT\\SHELL\\OPEN\\COMMAND\\";
            Microsoft.Win32.RegistryKey registry = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(ACROBAT_REG_PATH);
            if (registry != null)
            {
                var pathOtherWay = Convert.ToString(registry.GetValue(""));
                if (! string.IsNullOrEmpty(pathOtherWay))
                {
                    string[] arrPieces = pathOtherWay.Split('\"');
                    for (int i = 0; i < arrPieces.Length; i++)
                    {
                        if (arrPieces[i].ToUpper().IndexOf("ACROBAT") > -1)
                        {
                            g_strAcrobatExe = arrPieces[i];
                            break;
                        }
                    }
                }
            }
        }

        private void btnCoreAnalysis_Click(object sender, System.EventArgs e)
		{
			if (this.btnOptimizer.Enabled == true) 
			{

				this.btnOptimizer.Dock = DockStyle.Top;
				this.btnOptimizer.Enabled=false;
				this.btnDB.Dock = DockStyle.Bottom;
				this.btnDB.Enabled = true;
				this.btnFVS.Dock = DockStyle.Bottom;
				this.btnFVS.Enabled = true;
				this.btnProcessor.Dock = DockStyle.Bottom;
				this.btnProcessor.Enabled=true;
				this.m_pnlOptimizer.Size = this.m_pnlCurrent.Size;
				this.m_pnlOptimizer.Location = this.m_pnlCurrent.Location;
				this.m_pnlCurrent.Visible=false;
				this.m_pnlCurrent = this.m_pnlOptimizer;
				this.m_pnlCurrent.Visible=true;
				this.m_pnlCurrent.Refresh();
				this.ChildWindowVisible("Treatment Optimizer:");
			}			

		}
		private void btnFVS_Click(object sender, System.EventArgs e)
		{
			if (this.btnFVS.Enabled == true) 
			{
				this.btnOptimizer.Dock = DockStyle.Bottom;
				this.btnOptimizer.Enabled = true;
				this.btnDB.Dock = DockStyle.Bottom;
				this.btnDB.Enabled = true;
				this.btnFVS.Dock = DockStyle.Top;
				this.btnFVS.Enabled = false;
				this.btnProcessor.Dock = DockStyle.Bottom;
				this.btnProcessor.Enabled = true;
				this.m_pnlFvs.Size = this.m_pnlCurrent.Size;
				this.m_pnlFvs.Location = this.m_pnlCurrent.Location;
				this.m_pnlCurrent.Visible=false;
				this.m_pnlCurrent = this.m_pnlFvs;
				this.m_pnlCurrent.Visible=true;
				this.m_pnlCurrent.Refresh();

				this.ChildWindowVisible("FVS:");
			}
			


		}

		private void btnProcessor_Click(object sender, System.EventArgs e)
		{
			if (this.btnProcessor.Enabled == true) 
			{
				this.btnOptimizer.Dock = DockStyle.Bottom;
				this.btnOptimizer.Enabled = true;
				this.btnDB.Dock = DockStyle.Bottom;
				this.btnDB.Enabled = true;
				this.btnFVS.Dock = DockStyle.Bottom;
				this.btnFVS.Enabled = true;
				this.btnProcessor.Dock = DockStyle.Top;
				this.btnProcessor.Enabled = false;
				this.m_pnlProcessor.Size = this.m_pnlCurrent.Size;
				this.m_pnlProcessor.Location = this.m_pnlCurrent.Location;
				this.m_pnlCurrent.Visible=false;
				this.m_pnlCurrent = this.m_pnlProcessor;
				this.m_pnlCurrent.Visible=true;
				this.m_pnlCurrent.Refresh();

				this.ChildWindowVisible("Processor:");
			}
		
		}
		private void btnDB_Click(object sender, System.EventArgs e)
		{
			if (this.btnDB.Enabled == true) 
			{
				this.btnOptimizer.Dock = DockStyle.Bottom;
				this.btnOptimizer.Enabled = true;
				this.btnDB.Dock = DockStyle.Top;
				this.btnDB.Enabled = false;
				this.btnFVS.Dock = DockStyle.Bottom;
				this.btnFVS.Enabled = true;
				this.btnProcessor.Dock = DockStyle.Bottom;
				this.btnProcessor.Enabled = true;
				this.m_pnlDb.Size = this.m_pnlCurrent.Size;
				this.m_pnlDb.Location = this.m_pnlCurrent.Location;
				this.m_pnlCurrent.Visible=false;
				this.m_pnlCurrent = this.m_pnlDb;
				this.m_pnlCurrent.Visible=true;
				this.m_pnlCurrent.Refresh();
				this.ChildWindowVisible("Database:");
				this.Refresh();
			}
		}
		private bool IsChildWindowVisible(string strWinText)
		{
			foreach (Form child in this.MdiChildren)

			{
				
				if (child.Text.Trim().ToUpper() == strWinText.Trim().ToUpper())
				{
					return true;
				}
			}
			return false;
           
		}
		private void ChildWindowVisible(string strWinText)
		{
			foreach (Form child in this.MdiChildren)

			{
				
				if (child.Text.IndexOf(strWinText) >= 0) 
				{
					if (child.MaximizeBox == false)
					{
						child.WindowState = System.Windows.Forms.FormWindowState.Normal;
					}
					child.Visible=true;
				}
				else 
				{
					child.Visible=false;
				}
			}
		}
		public string getProjectDirectory()
		{
			
			return this.frmProject.uc_project1.m_strProjectDirectory;
		}
        public void OpenOptimizerScenario(string p_strType, frmOptimizerScenario p_frmOptimizerScenario)
		{
            FIA_Biosum_Manager.frmOptimizerScenario oFrmOptimizerScenario = new frmOptimizerScenario(this);

			DialogResult result;
			if (p_strType=="Open")
			{
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "=====================   OpenOptimizerScenario   =====================\r\n");


                oFrmOptimizerScenario.InitializeOpenScenario(frmProject.uc_project1.m_strDebugFile);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: frmOptimizerScenario.InitializeOpenScenario() successful! \r\n");
                oFrmOptimizerScenario.uc_scenario_open1.Height = oFrmOptimizerScenario.uc_scenario_open1.m_intFullHt;
                oFrmOptimizerScenario.uc_scenario_open1.Width = oFrmOptimizerScenario.uc_scenario_open1.m_intFullWd;
                oFrmOptimizerScenario.Height = oFrmOptimizerScenario.uc_scenario_open1.Height + oFrmOptimizerScenario.uc_scenario_open1.Top + 50;

                result = oFrmOptimizerScenario.ShowDialog();
				if (result == DialogResult.OK)
				{
					frmOptimizerScenario oFrmScenario = new frmOptimizerScenario(this);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Initialized new frmOptimizerScenario\r\n");
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Existing scenario name: " + oFrmOptimizerScenario.uc_scenario_open1.txtScenarioId.Text + "\r\n");
                    }
                    oFrmScenario.Text = "Treatment Optimizer: Optimization Scenario (" + oFrmOptimizerScenario.uc_scenario_open1.txtScenarioId.Text.Trim() + ")";
					oFrmScenario.m_bScenarioOpen = true;
                    oFrmScenario.HelpChapter = "EDIT_SCENARIO";
                    oFrmScenario.uc_datasource1.strDataSourceMDBFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;
					oFrmScenario.uc_datasource1.strDataSourceTable = "scenario_datasource";
                    oFrmScenario.uc_datasource1.strScenarioId = oFrmOptimizerScenario.uc_scenario_open1.txtScenarioId.Text.Trim();
					oFrmScenario.uc_datasource1.strProjectDirectory = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
					oFrmScenario.uc_datasource1.LoadValues(frmProject.uc_project1.m_strDebugFile);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Loaded datasource values\r\n");

                    oFrmScenario.uc_scenario1.strScenarioDescription = oFrmOptimizerScenario.uc_scenario_open1.strScenarioDescription;
                    oFrmScenario.uc_scenario1.strScenarioId = oFrmOptimizerScenario.uc_scenario_open1.strScenarioId;
                    oFrmScenario.uc_scenario1.strScenarioPath = oFrmOptimizerScenario.uc_scenario_open1.strScenarioPath;
                    oFrmScenario.uc_scenario_notes1.ReferenceOptimizerScenarioForm=oFrmScenario;
					oFrmScenario.uc_scenario_notes1.LoadValues();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Loaded notes\r\n");
                    oFrmScenario.tlbScenario.Buttons[5].Visible = true; //properties
                    oFrmScenario.tlbScenario.Buttons[7].Visible = true; //copy
					oFrmScenario.MdiParent = this;
					oFrmScenario.Show();
				}
			}
			else
			{
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Creating new scenario\r\n");
                oFrmOptimizerScenario.InitializeNewScenario();
                oFrmOptimizerScenario.MinimizeBox = false;
                result = oFrmOptimizerScenario.ShowDialog();
				if (result == DialogResult.OK)
				{
					frmOptimizerScenario oFrmScenario = new frmOptimizerScenario(this);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Initialized new frmOptimizerScenario\r\n");
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: New scenario name: " + oFrmOptimizerScenario.uc_scenario1.txtScenarioId.Text + "\r\n");

                    }
                    oFrmScenario.Text = "Treatment Optimizer: Optimization Scenario (" + oFrmOptimizerScenario.uc_scenario1.txtScenarioId.Text.Trim() + ")";
					oFrmScenario.m_bScenarioOpen = true;
                    oFrmScenario.HelpChapter = "EDIT_SCENARIO";
                    oFrmScenario.uc_datasource1.strDataSourceMDBFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;
					oFrmScenario.uc_datasource1.strDataSourceTable = "scenario_datasource";
                    oFrmScenario.uc_datasource1.strScenarioId = oFrmOptimizerScenario.uc_scenario1.txtScenarioId.Text.Trim();
					oFrmScenario.uc_datasource1.strProjectDirectory = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
					oFrmScenario.uc_datasource1.LoadValues(frmProject.uc_project1.m_strDebugFile);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Loaded datasource values\r\n");

                    oFrmScenario.uc_scenario1.strScenarioDescription = oFrmOptimizerScenario.uc_scenario1.strScenarioDescription;
                    oFrmScenario.uc_scenario1.strScenarioId = oFrmOptimizerScenario.uc_scenario1.strScenarioId;
                    oFrmScenario.uc_scenario1.strScenarioPath = oFrmOptimizerScenario.uc_scenario1.strScenarioPath;
                    oFrmScenario.tlbScenario.Buttons[5].Visible = true; //properties
                    oFrmScenario.tlbScenario.Buttons[7].Visible = true; //copy
					oFrmScenario.MdiParent = this;
					oFrmScenario.Show();
                    if (p_frmOptimizerScenario != null)
                    {
                        p_frmOptimizerScenario.DialogResult = DialogResult.Cancel;
                    }
				}			
			}			


		}

        public void OpenOptimizerScenarioSqlite(string p_strType, frmOptimizerScenario p_frmOptimizerScenario)
        {
            FIA_Biosum_Manager.frmOptimizerScenario oFrmOptimizerScenario = new frmOptimizerScenario(this);

            DialogResult result;
            if (p_strType == "Open")
            {
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "=====================   OpenOptimizerScenario   =====================\r\n");


                oFrmOptimizerScenario.InitializeOpenScenario(frmProject.uc_project1.m_strDebugFile);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: frmOptimizerScenario.InitializeOpenScenario() successful! \r\n");
                oFrmOptimizerScenario.uc_scenario_open1.Height = oFrmOptimizerScenario.uc_scenario_open1.m_intFullHt;
                oFrmOptimizerScenario.uc_scenario_open1.Width = oFrmOptimizerScenario.uc_scenario_open1.m_intFullWd;
                oFrmOptimizerScenario.Height = oFrmOptimizerScenario.uc_scenario_open1.Height + oFrmOptimizerScenario.uc_scenario_open1.Top + 50;

                result = oFrmOptimizerScenario.ShowDialog();
                if (result == DialogResult.OK)
                {
                    frmOptimizerScenario oFrmScenario = new frmOptimizerScenario(this);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Initialized new frmOptimizerScenario\r\n");
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Existing scenario name: " + oFrmOptimizerScenario.uc_scenario_open1.txtScenarioId.Text + "\r\n");
                    }
                    oFrmScenario.Text = "Treatment Optimizer: Optimization Scenario (" + oFrmOptimizerScenario.uc_scenario_open1.txtScenarioId.Text.Trim() + ")";
                    oFrmScenario.m_bScenarioOpen = true;
                    oFrmScenario.HelpChapter = "EDIT_SCENARIO";
                    oFrmScenario.uc_datasource1.strDataSourceMDBFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
                    oFrmScenario.uc_datasource1.strDataSourceTable = "scenario_datasource";
                    oFrmScenario.uc_datasource1.strScenarioId = oFrmOptimizerScenario.uc_scenario_open1.txtScenarioId.Text.Trim();
                    oFrmScenario.uc_datasource1.strProjectDirectory = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
                    oFrmScenario.uc_datasource1.LoadValuesSqlite();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Loaded datasource values\r\n");

                    oFrmScenario.uc_scenario1.strScenarioDescription = oFrmOptimizerScenario.uc_scenario_open1.strScenarioDescription;
                    oFrmScenario.uc_scenario1.strScenarioId = oFrmOptimizerScenario.uc_scenario_open1.strScenarioId;
                    oFrmScenario.uc_scenario1.strScenarioPath = oFrmOptimizerScenario.uc_scenario_open1.strScenarioPath;
                    oFrmScenario.uc_scenario_notes1.ReferenceOptimizerScenarioForm = oFrmScenario;
                    oFrmScenario.uc_scenario_notes1.LoadValuesSqlite();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Loaded notes\r\n");
                    oFrmScenario.tlbScenario.Buttons[5].Visible = true; //properties
                    oFrmScenario.tlbScenario.Buttons[7].Visible = true; //copy
                    oFrmScenario.MdiParent = this;
                    oFrmScenario.Show();
                }
            }
            else
            {
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Creating new scenario\r\n");
                oFrmOptimizerScenario.InitializeNewScenario();
                oFrmOptimizerScenario.MinimizeBox = false;
                result = oFrmOptimizerScenario.ShowDialog();
                if (result == DialogResult.OK)
                {
                    frmOptimizerScenario oFrmScenario = new frmOptimizerScenario(this);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Initialized new frmOptimizerScenario\r\n");
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: New scenario name: " + oFrmOptimizerScenario.uc_scenario1.txtScenarioId.Text + "\r\n");

                    }
                    oFrmScenario.Text = "Treatment Optimizer: Optimization Scenario (" + oFrmOptimizerScenario.uc_scenario1.txtScenarioId.Text.Trim() + ")";
                    oFrmScenario.m_bScenarioOpen = true;
                    oFrmScenario.HelpChapter = "EDIT_SCENARIO";
                    oFrmScenario.uc_datasource1.strDataSourceMDBFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
                    oFrmScenario.uc_datasource1.strDataSourceTable = "scenario_datasource";
                    oFrmScenario.uc_datasource1.strScenarioId = oFrmOptimizerScenario.uc_scenario1.txtScenarioId.Text.Trim();
                    oFrmScenario.uc_datasource1.strProjectDirectory = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
                    oFrmScenario.uc_datasource1.LoadValuesSqlite();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "OpenOptimizerScenario: Loaded datasource values\r\n");

                    oFrmScenario.uc_scenario1.strScenarioDescription = oFrmOptimizerScenario.uc_scenario1.strScenarioDescription;
                    oFrmScenario.uc_scenario1.strScenarioId = oFrmOptimizerScenario.uc_scenario1.strScenarioId;
                    oFrmScenario.uc_scenario1.strScenarioPath = oFrmOptimizerScenario.uc_scenario1.strScenarioPath;
                    oFrmScenario.tlbScenario.Buttons[5].Visible = true; //properties
                    oFrmScenario.tlbScenario.Buttons[7].Visible = true; //copy
                    oFrmScenario.MdiParent = this;
                    oFrmScenario.Show();
                    if (p_frmOptimizerScenario != null)
                    {
                        p_frmOptimizerScenario.DialogResult = DialogResult.Cancel;
                    }
                }
            }


        }

        public void OpenProcessorScenario(string p_strType,frmProcessorScenario p_frmProcessorScenario)
		{
			FIA_Biosum_Manager.frmProcessorScenario oFrmProcessorScenario = new frmProcessorScenario(this);

			DialogResult result;
			if (p_strType=="Open")
			{
				oFrmProcessorScenario.InitializeOpenScenario(frmProject.uc_project1.m_strDebugFile);


				oFrmProcessorScenario.uc_scenario_open1.Height = oFrmProcessorScenario.uc_scenario_open1.m_intFullHt;
				oFrmProcessorScenario.uc_scenario_open1.Width = oFrmProcessorScenario.uc_scenario_open1.m_intFullWd;
				oFrmProcessorScenario.Height = oFrmProcessorScenario.uc_scenario_open1.Height + oFrmProcessorScenario.uc_scenario_open1.Top + 50;
                oFrmProcessorScenario.MinimizeBox = false;

				result = oFrmProcessorScenario.ShowDialog();
				if (result == DialogResult.OK)
				{
					frmProcessorScenario oFrmScenario = new frmProcessorScenario(this);

					oFrmScenario.Text = "Processor: Scenario (" + oFrmProcessorScenario.uc_scenario_open1.strScenarioId.Trim() + ")";
					oFrmScenario.m_bScenarioOpen = true;

                    
					oFrmScenario.uc_datasource1.strDataSourceMDBFile =  frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\processor\\db\\scenario_processor_rule_definitions.mdb";
					oFrmScenario.uc_datasource1.strDataSourceTable = "scenario_datasource";
					oFrmScenario.uc_datasource1.strScenarioId = oFrmProcessorScenario.uc_scenario_open1.txtScenarioId.Text.Trim();
					oFrmScenario.uc_datasource1.strProjectDirectory = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
                    if (oFrmProcessorScenario.m_bUsingSqlite == false)
                    {
                        oFrmScenario.uc_datasource1.LoadValues(frmProject.uc_project1.m_strDebugFile);
                    }
                    else
                    {
                        oFrmScenario.uc_datasource1.LoadValuesSqlite();
                    }
                    oFrmScenario.uc_scenario1.strScenarioDescription = oFrmProcessorScenario.uc_scenario_open1.strScenarioDescription;
					oFrmScenario.uc_scenario1.strScenarioId = oFrmProcessorScenario.uc_scenario_open1.strScenarioId;
					oFrmScenario.uc_scenario1.strScenarioPath = oFrmProcessorScenario.uc_scenario_open1.strScenarioPath;
					oFrmScenario.uc_scenario_notes1.ReferenceProcessorScenarioForm=oFrmScenario;
					oFrmScenario.uc_scenario_notes1.ScenarioType="processor";
                    oFrmScenario.uc_datasource1.UsingSqlite = oFrmProcessorScenario.m_bUsingSqlite;
                    if (oFrmProcessorScenario.m_bUsingSqlite == false)
                    {
                        oFrmScenario.uc_scenario_notes1.LoadValues();
                    }
                    else
                    {
                        oFrmScenario.uc_scenario_notes1.LoadValuesSqlite();
                    }
					oFrmScenario.MdiParent = this;
                    oFrmScenario.m_oProcessorScenarioItem.ScenarioId = oFrmScenario.uc_scenario1.strScenarioId;
                    oFrmScenario.m_oProcessorScenarioItem.DbPath = oFrmScenario.uc_scenario1.strScenarioPath;
                    oFrmScenario.m_oProcessorScenarioItem.Description = oFrmScenario.uc_scenario1.strScenarioDescription;
                    oFrmScenario.m_oProcessorScenarioItem.DbFileName = oFrmScenario.uc_datasource1.strDataSourceMDBFile;
					oFrmScenario.Show();
                    if (p_frmProcessorScenario != null)
                    {
                        p_frmProcessorScenario.DialogResult = DialogResult.Cancel;
                    }
				}
			}
			else
			{
				oFrmProcessorScenario.InitializeNewScenario();
                oFrmProcessorScenario.MinimizeBox = false;			 
				result = oFrmProcessorScenario.ShowDialog();
				if (result == DialogResult.OK)
				{
                   
					frmProcessorScenario oFrmScenario = new frmProcessorScenario(this);
					oFrmScenario.Text = "Processor: Scenario (" + oFrmProcessorScenario.uc_scenario1.txtScenarioId.Text.Trim() + ")";
					oFrmScenario.m_bScenarioOpen = true;
					oFrmScenario.uc_datasource1.strDataSourceMDBFile =  frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\processor\\db\\scenario_processor_rule_definitions.mdb";
					oFrmScenario.uc_datasource1.strDataSourceTable = "scenario_datasource";
					oFrmScenario.uc_datasource1.strScenarioId = oFrmProcessorScenario.uc_scenario1.txtScenarioId.Text.Trim();
					oFrmScenario.uc_datasource1.strProjectDirectory = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
                    oFrmScenario.uc_datasource1.UsingSqlite = oFrmProcessorScenario.m_bUsingSqlite;
                    if (oFrmProcessorScenario.m_bUsingSqlite == false)
                    {
                        oFrmScenario.uc_datasource1.LoadValues(frmProject.uc_project1.m_strDebugFile);
                    }
                    else
                    {
                        oFrmScenario.uc_datasource1.LoadValuesSqlite();
                    }
                    
					oFrmScenario.uc_scenario1.strScenarioDescription = oFrmProcessorScenario.uc_scenario1.strScenarioDescription;
					oFrmScenario.uc_scenario1.strScenarioId = oFrmProcessorScenario.uc_scenario1.strScenarioId;
					oFrmScenario.uc_scenario1.strScenarioPath = oFrmProcessorScenario.uc_scenario1.strScenarioPath;
					oFrmScenario.MdiParent = this;
					oFrmScenario.Show();
                    if (p_frmProcessorScenario != null)
                    {
                        p_frmProcessorScenario.DialogResult = DialogResult.Cancel;
                    }
				}
				

			}
			


		}

        public void button_click(string strText)
        {

            if (this.btnOptimizer.Enabled == false)
            {
                if (strText.Trim().ToUpper() == "DEFINE CALCULATED VARIABLES")
                {
                    //check to see if the form has already been loaded
                    if (this.IsChildWindowVisible("Treatment Optimizer: Define Calculated Variables") == false)
                    {

                        this.m_frmCoreUserVariables = new frmDialog(this);
                        this.m_frmCoreUserVariables.MaximizeBox = false;
                        this.m_frmCoreUserVariables.MinimizeBox = false;
                        this.m_frmCoreUserVariables.BackColor = System.Drawing.SystemColors.Control;
                        this.m_frmCoreUserVariables.Text = "Treatment Optimizer: Define Calculated Variables";
                        this.m_frmCoreUserVariables.MdiParent = this;
                        this.m_frmCoreUserVariables.Initialize_Core_User_Variables_User_Control();

                        this.m_frmCoreUserVariables.DisposeOfFormWhenClosing = true;

                        this.m_frmCoreUserVariables.Width = this.m_frmCoreUserVariables.uc_core_scenario_weighted_average1.m_DialogWd;
                        this.m_frmCoreUserVariables.Height = this.m_frmCoreUserVariables.uc_core_scenario_weighted_average1.m_DialogHt;
                        this.m_frmCoreUserVariables.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                        this.m_frmCoreUserVariables.uc_core_scenario_weighted_average1.Top = 0;
                        this.m_frmCoreUserVariables.uc_core_scenario_weighted_average1.Left = 0;

                        this.m_frmCoreUserVariables.uc_core_scenario_weighted_average1.Visible = true;
                        this.m_frmCoreUserVariables.MinimizeMainForm = true;
                        this.m_frmCoreUserVariables.ParentControl = frmMain.g_oFrmMain;
                        this.m_frmCoreUserVariables.Show();

                    }
                    else
                    {
                        if (this.m_frmCoreMerge.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                            this.m_frmCoreMerge.WindowState = System.Windows.Forms.FormWindowState.Normal;

                        this.m_frmCoreMerge.Focus();

                    }
                }
                else if (strText.Trim().ToUpper() == "LOAD GIS DATA")
                {
                    //GisTools oGisTools = new GisTools();
                    //bool bTablesHaveData = false;
                    //bool bTablesExist = oGisTools.CheckForExistingDataSqlite(this.frmProject.uc_project1.m_strProjectDirectory, out bTablesHaveData);
                    //bool bCreateBackups = false;
                    //string strMessage = "";
                    //bool bSuccess = true;

                    //string strMasterDb = frmMain.g_oEnv.strApplicationDataDirectory.Trim() + frmMain.g_strBiosumDataDir +
                    //    "\\" + Tables.TravelTime.DefaultMasterTravelTimeDbFile;
                    //if (!System.IO.File.Exists(strMasterDb))
                    //{
                    //    MessageBox.Show("The source gis_travel_times_master.db is required but does not exist in the " +
                    //        frmMain.g_oEnv.strApplicationDataDirectory.Trim() + frmMain.g_strBiosumDataDir + " folder. \r\n\r\n" +
                    //        "Please download a copy of this database into the FIABiosum folder!!", "FIA Biosum");
                    //    return;
                    //}
                    //if (bTablesHaveData == true)
                    //{
                    //    strMessage = "BioSum has found existing data in your gis data tables. Do you wish to overwrite existing data? " +
                    //        "This process cannot be reversed!!";
                    //    DialogResult res = MessageBox.Show(strMessage, "FIA BioSum", MessageBoxButtons.YesNo);
                    //    if (res != DialogResult.Yes)
                    //    {
                    //        MessageBox.Show("GIS data load terminated!!", "FIA BioSum");
                    //        return;
                    //    }
                    //    else
                    //    {
                    //        strMessage = "Would you like BioSum to make a copy of your existing GIS database? The name of the database backup will include today's date.";
                    //        res = MessageBox.Show(strMessage, "FIA BioSum", MessageBoxButtons.YesNo);
                    //        if (res == DialogResult.Yes)
                    //        {
                    //            bCreateBackups = true;
                    //        }
                    //    }
                    //}
                    //// Check for existence of MoveDist_ft_REPLACEMENT field
                    //string strSourceField = "MoveDist_ft_REPLACEMENT";
                    //if (oGisTools.CheckPlotGisTable(strMasterDb, strSourceField))
                    //{
                    //    strMessage = "Do you want to update the plot yarding distance from the plot_gis table in the master travel times database ?";
                    //    DialogResult res2 = MessageBox.Show(strMessage, "FIA BioSum", MessageBoxButtons.YesNo);
                    //    if (res2 != DialogResult.Yes)
                    //    {
                    //        strSourceField = "";
                    //    }
                    //}
                    //else
                    //{
                    //    strSourceField = "";
                    //}

                    //this.ActivateStandByAnimation(
                    //     this.WindowState,
                    //     this.Left,
                    //     this.Height,
                    //     this.Width,
                    //     this.Top);
                    //g_sbpInfo.Text = "Loading gis data...Stand by";

                    //if (bCreateBackups == true)
                    //{
                    //    g_sbpInfo.Text = "Backing up old gis data...Stand by";
                    //    bSuccess = oGisTools.BackupGisData();
                    //}
                    //if (bSuccess == true)
                    //{
                    //    int intRowCount = oGisTools.LoadSqliteGisData(strSourceField, frmProject.uc_project1.m_strDebugFile);
                    //    if (intRowCount < 1)
                    //    {
                    //        MessageBox.Show("An error occurred while loading the GIS data!!", "FIA BioSum");
                    //        this.DeactivateStandByAnimation();
                    //        return;
                    //    }
                    //}
                    //else
                    //{
                    //    MessageBox.Show("An error occurred while backing up the tables. GIS data load terminated!!", "FIA BioSum");
                    //    g_sbpInfo.Text = "Ready";
                    //    this.DeactivateStandByAnimation();
                    //    return;
                    //}
                    //g_sbpInfo.Text = "Ready";
                    //this.DeactivateStandByAnimation();
                    //MessageBox.Show("If you updated existing GIS data, verify the selected sites in Treatment Optimizer. GIS data successfully loaded!");
                    // END EXISTING
                    this.m_frmLoadGis = new frmDialog(this);
                    this.m_frmLoadGis.MaximizeBox = false;
                    this.m_frmLoadGis.MinimizeBox = false;
                    this.m_frmLoadGis.BackColor = System.Drawing.SystemColors.Control;
                    this.m_frmLoadGis.Text = "Load GIS Data";
                    this.m_frmLoadGis.MdiParent = this;
                    this.m_frmLoadGis.Initialize_Load_Gis_Data_User_Control();

                    this.m_frmLoadGis.DisposeOfFormWhenClosing = true;

                    this.m_frmLoadGis.Width = this.m_frmLoadGis.uc_optimizer_load_gis_data1.Width + 25;
                    this.m_frmLoadGis.Height = this.m_frmLoadGis.uc_optimizer_load_gis_data1.Height + 40;
                    this.m_frmLoadGis.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                    this.m_frmLoadGis.uc_optimizer_load_gis_data1.Top = 0;
                    this.m_frmLoadGis.uc_optimizer_load_gis_data1.Left = 0;

                    this.m_frmLoadGis.uc_optimizer_load_gis_data1.Visible = true;
                    this.m_frmLoadGis.MinimizeMainForm = true;
                    this.m_frmLoadGis.ParentControl = frmMain.g_oFrmMain;
                    this.m_frmLoadGis.Show();
                }
                else if (strText.Trim().ToUpper() == "OPTIMIZATION SCENARIO")
                {
                    System.Text.StringBuilder strFullPath;

                    string strProjDir = getProjectDirectory();
                    strFullPath = new System.Text.StringBuilder(strProjDir);
                    strFullPath.Append("\\");
                    strFullPath.Append(Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile);
                    SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
                    string strConn = oDataMgr.GetConnectionString(strFullPath.ToString());
                    long lngCount = oDataMgr.getRecordCount(strConn, "select count(*) from scenario", "scenario");
                    if (oDataMgr.m_intError == 0)
                    {
                        if (lngCount > 0)
                        {
                            OpenOptimizerScenarioSqlite("Open", null);
                        }
                        else
                        {

                            OpenOptimizerScenarioSqlite("New", null);
                        }
                    }
                    oDataMgr = null;

                }
                else if (strText.Trim().ToUpper() == "UPGRADE FOR V5.11.0")
                {
                    //version_control oVersCtl = new version_control();
                    //oVersCtl.UpdateDatasources_5_11_0(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim());
                    //MessageBox.Show("Upgrade Complete!");
                    //this.m_frmSqlite = new frmDialog(this);
                    //this.m_frmSqlite.MaximizeBox = false;
                    //this.m_frmSqlite.MinimizeBox = false;
                    //this.m_frmSqlite.BackColor = System.Drawing.SystemColors.Control;
                    //this.m_frmSqlite.Text = "Treatment Optimizer: Export to SQLITE";
                    //this.m_frmSqlite.MdiParent = this;
                    //this.m_frmSqlite.Initialize_Optimizer_Sqlite_User_Control();

                    //this.m_frmSqlite.DisposeOfFormWhenClosing = true;

                    //this.m_frmSqlite.Width = this.m_frmSqlite.uc_optimizer_sqlite_export1.Width + 25;
                    //this.m_frmSqlite.Height = this.m_frmSqlite.uc_optimizer_sqlite_export1.Height + 40;
                    //this.m_frmSqlite.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                    //this.m_frmSqlite.uc_optimizer_sqlite_export1.Top = 0;
                    //this.m_frmSqlite.uc_optimizer_sqlite_export1.Left = 0;

                    //this.m_frmSqlite.uc_optimizer_sqlite_export1.ReferenceFormDialog = this.m_frmSqlite;
                    //this.m_frmSqlite.uc_optimizer_sqlite_export1.Visible = true;
                    //this.m_frmSqlite.MinimizeMainForm = true;
                    //this.m_frmSqlite.ParentControl = frmMain.g_oFrmMain;
                    //this.m_frmSqlite.Show();
                }
                else if (strText.Trim().ToUpper() == "JOIN DATA FROM MULTIPLE SCENARIOS")
                {
                    //check to see if the form has already been loaded
                    if (this.IsChildWindowVisible("Treatment Optimizer: Join Data From Multiple Scenarios") == false)
                    {

                        this.m_frmCoreMerge = new frmDialog(this);
                        this.m_frmCoreMerge.MaximizeBox = false;
                        this.m_frmCoreMerge.BackColor = System.Drawing.SystemColors.Control;
                        this.m_frmCoreMerge.Text = "Treatment Optimizer: Join Data From Multiple Scenarios";
                        this.m_frmCoreMerge.MdiParent = this;
                        this.m_frmCoreMerge.Initialize_Join_Scenario_User_Control();
                        this.m_frmCoreMerge.uc_merge_tables1.Top = 0;
                        this.m_frmCoreMerge.uc_merge_tables1.Left = 0;
                        int intHt = 0;



                        int intHt2 = this.m_frmCoreMerge.uc_merge_tables1.groupBox1.Top + this.m_frmCoreMerge.uc_merge_tables1.lblTitle.Height + this.m_frmCoreMerge.uc_merge_tables1.grpboxOpen.Top + this.m_frmCoreMerge.uc_merge_tables1.grpboxOpen.Height;
                        int intTop = this.m_frmCoreMerge.uc_merge_tables1.groupBox1.Top;
                        while (intTop + intHt2 + 20
                            >= intHt)
                        {
                            intHt += 10;

                        }
                        this.m_frmCoreMerge.Height = intHt;


                        this.m_frmCoreMerge.DisposeOfFormWhenClosing = true;
                        this.m_frmCoreMerge.Height = this.m_frmCoreMerge.uc_merge_tables1.Height + this.m_frmCoreMerge.uc_merge_tables1.lblTitle.Height;
                        this.m_frmCoreMerge.Width = ((this.m_frmCoreMerge.uc_merge_tables1.Left + this.m_frmCoreMerge.uc_merge_tables1.groupBox1.Left + this.m_frmCoreMerge.uc_merge_tables1.grpboxOpen.Left) * 2) + this.m_frmCoreMerge.uc_merge_tables1.grpboxOpen.Width + 5;

                        this.m_frmCoreMerge.Left = 0;
                        this.m_frmCoreMerge.Top = 0;
                        this.m_frmCoreMerge.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                        this.m_frmCoreMerge.Show();

                    }
                    else
                    {
                        if (this.m_frmCoreMerge.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                            this.m_frmCoreMerge.WindowState = System.Windows.Forms.FormWindowState.Normal;

                        this.m_frmCoreMerge.Focus();

                    }

                }
            }
            else if (this.btnDB.Enabled == false)
            {
                switch (strText.Trim().ToUpper())
                {
                    case "MANAGE TABLES":
                        StartManageTablesDialog();
                        break;

                    case "MANAGE SQLITE TABLES":
                        StartManageSQLiteTablesDialog();
                        break;

                    case "PLOT DATA":


                        //check to see if the form has already been loaded
                        if (this.IsChildWindowVisible("Database: Plot Data") == false)
                        {

                            this.m_frmPlotData = new frmDialog(this);
                            this.m_frmPlotData.BackColor = System.Drawing.SystemColors.Control;
                            this.m_frmPlotData.Text = "Database: Plot Data";
                            this.m_frmPlotData.MinimizeBox = true;
                            this.m_frmPlotData.MaximizeBox = false;
                            this.m_frmPlotData.MdiParent = this;
                            this.m_frmPlotData.Initialize_Plot_Data_Add_Edit_User_Control();

                            this.m_frmPlotData.Height = 0;
                            this.m_frmPlotData.Width = 0;
                            if (this.m_frmPlotData.uc_plot_add_edit1.Top + this.m_frmPlotData.uc_plot_add_edit1.Height > this.m_frmPlotData.ClientSize.Height + 2)
                            {
                                for (int x = 1; ; x++)
                                {
                                    this.m_frmPlotData.Height = x;
                                    if (this.m_frmPlotData.uc_plot_add_edit1.Top +
                                        this.m_frmPlotData.uc_plot_add_edit1.Height <
                                        this.m_frmPlotData.ClientSize.Height)
                                    {
                                        break;
                                    }
                                }

                            }
                            if (this.m_frmPlotData.uc_plot_add_edit1.Left + this.m_frmPlotData.uc_plot_add_edit1.Width > this.m_frmPlotData.ClientSize.Width + 2)
                            {
                                for (int x = 1; ; x++)
                                {
                                    this.m_frmPlotData.Width = x;
                                    if (this.m_frmPlotData.uc_plot_add_edit1.Left +
                                        this.m_frmPlotData.uc_plot_add_edit1.Width <
                                        this.m_frmPlotData.ClientSize.Width)
                                    {
                                        break;
                                    }
                                }

                            }
                            this.m_frmPlotData.MaximumSize = this.m_frmPlotData.Size;

                            this.m_frmPlotData.uc_plot_add_edit1.Dock = System.Windows.Forms.DockStyle.Fill;
                            this.m_frmPlotData.uc_plot_add_edit1.Top = 0;
                            this.m_frmPlotData.uc_plot_add_edit1.Left = 0;

                            this.m_frmPlotData.uc_plot_add_edit1.Visible = true;
                            this.m_frmPlotData.DisposeOfFormWhenClosing = true;
                            this.m_frmPlotData.Show();
                            this.m_frmPlotData.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;


                        }
                        else
                        {
                            if (this.m_frmPlotData.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                                this.m_frmPlotData.WindowState = System.Windows.Forms.FormWindowState.Normal;

                            this.m_frmPlotData.Focus();

                        }



                        break;
                    case "WOOD PROCESSING SITES":
                        StartPSiteDialog(this);
                        break;

                    case "PROJECT DATA SOURCES":
                        //check to see if the form has already been loaded
                        if (this.IsChildWindowVisible("Database: Project Data Sources") == false)
                        {
                            frmMain.g_sbpInfo.Text = "Loading Project Data Sources...Stand By";

                            this.m_frmDataSource = new frmDialog(this);
                            this.m_frmDataSource.MaximizeBox = true;
                            this.m_frmDataSource.BackColor = System.Drawing.SystemColors.Control;
                            this.m_frmDataSource.Text = "Database: Project Data Sources";
                            this.m_frmDataSource.MdiParent = this;
                            FIA_Biosum_Manager.uc_datasource p_uc = new uc_datasource(this.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\project.mdb");
                            this.m_frmDataSource.Controls.Add(p_uc);
                            p_uc.Dock = System.Windows.Forms.DockStyle.Fill;

                            p_uc.strProjectDirectory = this.frmProject.uc_project1.txtRootDirectory.Text.Trim();




                            this.m_frmDataSource.Height = 0;
                            this.m_frmDataSource.Width = 0;
                            if (p_uc.Top + p_uc.Height > this.m_frmDataSource.ClientSize.Height + 2)
                            {
                                for (int x = 1; ; x++)
                                {
                                    this.m_frmDataSource.Height = x;
                                    if (p_uc.Top +
                                        p_uc.Height <
                                        this.m_frmDataSource.ClientSize.Height)
                                    {
                                        break;
                                    }
                                }

                            }
                            if (p_uc.Left + p_uc.Width > this.m_frmDataSource.ClientSize.Width + 2)
                            {
                                for (int x = 1; ; x++)
                                {
                                    this.m_frmDataSource.Width = x;
                                    if (p_uc.Left +
                                        p_uc.Width <
                                        this.m_frmDataSource.ClientSize.Width)
                                    {
                                        break;
                                    }
                                }

                            }


                            p_uc.populate_listview_grid();
                            p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
                            this.m_frmDataSource.Left = 0;
                            this.m_frmDataSource.Top = 0;
                            this.m_frmDataSource.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                            this.m_frmDataSource.DisposeOfFormWhenClosing = true;
                            frmMain.g_sbpInfo.Text = "Ready";
                            this.m_frmDataSource.Show();

                        }
                        else
                        {
                            if (this.m_frmDataSource.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                                this.m_frmDataSource.WindowState = System.Windows.Forms.FormWindowState.Normal;

                            this.m_frmDataSource.Focus();

                        }
                        break;

                    case "GENERATE TRAVEL TIMES":
                        FIA_Biosum_Travel_Times_Generator.generate_travel_times p_trvltm = new FIA_Biosum_Travel_Times_Generator.generate_travel_times(this);
                        p_trvltm.create_travel_times();
                        break;

                }
            }
            else if (this.btnProcessor.Enabled == false)
            {
                switch (strText.Trim().ToUpper())
                {
                    case "TREE SPECIES":
                        this.LoadProcessorTreeSpcForm(this, "Processor");
                        break;

                    case "START BIOSUM PROCESSOR":
                        StartBiosumProcessorDialog();
                        break;
                    default:
                        break;

                }
            }
            else if (this.btnFVS.Enabled == false)
            {
                switch (strText.Trim().ToUpper())
                {
                    case "PLOT FVS VARIANTS":
                        StartPlotFVSVariantsDialog(this);

                        break;

                    case "RX":
                        StartRxDialog(this);
                        break;
                    case "RX PACKAGE":
                        StartRxPackageDialog(this);

                        break;
                    case "FVS INPUT":
                        StartFVSInputDataDialog();


                        break;
                    case "TREE SPECIES CONVERSION":

                        //check to see if the form has already been loaded
                        if (this.IsChildWindowVisible("FVS: Tree Species Conversion") == false)
                        {
                            this.ActivateStandByAnimation(this.WindowState, this.Left, this.Height, this.Width, this.Top);
                            this.m_frmFvsTreeSpcCvt = new frmDialog(this);
                            this.m_frmFvsTreeSpcCvt.MaximizeBox = true;
                            this.m_frmFvsTreeSpcCvt.BackColor = System.Drawing.SystemColors.Control;
                            this.m_frmFvsTreeSpcCvt.Text = "FVS: Tree Species Conversion";
                            FIA_Biosum_Manager.uc_fvs_tree_spc_conversion p_uc = new uc_fvs_tree_spc_conversion(this.frmProject.uc_project1.txtRootDirectory.Text.Trim());
                            if (p_uc.m_intError < 0)
                            {
                                this.m_frmFvsTreeSpcCvt.Dispose();
                                return;
                            }
                            this.m_frmFvsTreeSpcCvt.Controls.Add(p_uc);
                            this.m_frmFvsTreeSpcCvt.TreeSpeciesConversionUserControl = p_uc;
                            this.m_frmFvsTreeSpcCvt.Height = 0;
                            this.m_frmFvsTreeSpcCvt.Width = 0;
                            if (p_uc.Top + p_uc.Height > this.m_frmFvsTreeSpcCvt.ClientSize.Height + 2)
                            {
                                for (int x = 1; ; x++)
                                {
                                    this.m_frmFvsTreeSpcCvt.Height = x;
                                    if (p_uc.Top +
                                        p_uc.Height <
                                        this.m_frmFvsTreeSpcCvt.ClientSize.Height)
                                    {
                                        break;
                                    }
                                }

                            }
                            if (p_uc.Left + p_uc.Width > this.m_frmFvsTreeSpcCvt.ClientSize.Width + 2)
                            {
                                for (int x = 1; ; x++)
                                {
                                    this.m_frmFvsTreeSpcCvt.Width = x;
                                    if (p_uc.Left +
                                        p_uc.Width <
                                        this.m_frmFvsTreeSpcCvt.ClientSize.Width)
                                    {
                                        break;
                                    }
                                }

                            }
                            p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
                            p_uc.loadvalues();


                            this.m_frmFvsTreeSpcCvt.Left = 0;
                            this.m_frmFvsTreeSpcCvt.Top = 0;
                            this.m_frmFvsTreeSpcCvt.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                            this.m_frmFvsTreeSpcCvt.DisposeOfFormWhenClosing = true;
                            this.m_frmFvsTreeSpcCvt.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                            this.DeactivateStandByAnimation();
                            this.m_frmFvsTreeSpcCvt.ShowDialog();

                        }
                        else
                        {
                            if (this.m_frmFvsTreeSpcCvt.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                                this.m_frmFvsTreeSpcCvt.WindowState = System.Windows.Forms.FormWindowState.Normal;

                            this.m_frmFvsTreeSpcCvt.Focus();

                        }

                        break;
                    case "TREE SPECIES":
                        this.LoadProcessorTreeSpcForm(this, "FVS");
                        break;
                    case "CREATE FVSOUT MDBS":
                        // Call StartFVSCreateMdbsDialog() method here
                        StartFVSCreateMdbsDialog();
                        break;
                    case "FVS OUTPUT DATA":
                        StartFVSOutputDataDialog();

                        break;

                    default:
                        break;
                }

            }
        }
        public void StartPlotFVSVariantsDialog(System.Windows.Forms.Control p_oParentControl)
        {
            //check to see if the form has already been loaded
            if (this.IsChildWindowVisible("FVS: Plot FVS Variant") == false)
            {
                frmMain.g_sbpInfo.Text = "Loading Plot FVS Variants...Stand By";
               
                this.m_frmFvsVariant = new frmDialog(this);
                this.m_frmFvsVariant.MaximizeBox = true;
                this.m_frmFvsVariant.BackColor = System.Drawing.SystemColors.Control;
                this.m_frmFvsVariant.Text = "FVS: Plot FVS Variant";

                FIA_Biosum_Manager.uc_plot_fvs_variant p_uc = new uc_plot_fvs_variant(this.frmProject.uc_project1.txtRootDirectory.Text.Trim());
                if (p_uc.m_intError < 0)
                {
                    this.m_frmFvsVariant.Dispose();
                    return;
                }
                this.m_frmFvsVariant.Controls.Add(p_uc);
                this.m_frmFvsVariant.PlotFvsVariantUserControl = p_uc;
                this.m_frmFvsVariant.Height = 0;
                this.m_frmFvsVariant.Width = 0;
                if (p_uc.Top + p_uc.Height > this.m_frmFvsVariant.ClientSize.Height + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmFvsVariant.Height = x;
                        if (p_uc.Top +
                            p_uc.Height <
                            this.m_frmFvsVariant.ClientSize.Height)
                        {
                            break;
                        }
                    }

                }
                if (p_uc.Left + p_uc.Width > this.m_frmFvsVariant.ClientSize.Width + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmFvsVariant.Width = x;
                        if (p_uc.Left +
                            p_uc.Width <
                            this.m_frmFvsVariant.ClientSize.Width)
                        {
                            break;
                        }
                    }

                }
                p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
                p_uc.loadvalues();


                this.m_frmFvsVariant.Left = 0;
                this.m_frmFvsVariant.Top = 0;
                this.m_frmFvsVariant.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                this.m_frmFvsVariant.DisposeOfFormWhenClosing = true;
                this.m_frmFvsVariant.MinimizeMainForm = true;
                this.m_frmFvsVariant.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                frmMain.g_sbpInfo.Text = "Ready";
                p_oParentControl.Enabled = false;
                m_frmFvsVariant.ParentControl = p_oParentControl;
                this.m_frmFvsVariant.Show();

            }
            else
            {
                if (this.m_frmFvsVariant.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    this.m_frmFvsVariant.WindowState = System.Windows.Forms.FormWindowState.Normal;

                this.m_frmFvsVariant.Focus();

            }
        }
        public void StartRxDialog(Control p_oParentControl)
        {
            //check to see if the form has already been loaded
            if (this.IsChildWindowVisible("FVS: Treatments") == false)
            {
                frmMain.g_sbpInfo.Text = "Loading Treatment Definitions...Stand By";
                this.m_frmRx = new frmDialog(this);
                this.m_frmRx.MaximizeBox = true;
                this.m_frmRx.BackColor = System.Drawing.SystemColors.Control;
                this.m_frmRx.Text = "FVS: Treatments";
                this.m_frmRx.Initialize_Rx_User_Control();



                this.m_frmRx.Height = 0;
                this.m_frmRx.Width = 0;
                if (this.m_frmRx.uc_rx_list1.Top + this.m_frmRx.uc_rx_list1.Height > this.m_frmRx.ClientSize.Height + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmRx.Height = x;
                        if (this.m_frmRx.uc_rx_list1.Top +
                            this.m_frmRx.uc_rx_list1.Height <
                            this.m_frmRx.ClientSize.Height)
                        {
                            break;
                        }
                    }

                }
                if (this.m_frmRx.uc_rx_list1.Left + this.m_frmRx.uc_rx_list1.Width > this.m_frmRx.ClientSize.Width + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmRx.Width = x;
                        if (this.m_frmRx.uc_rx_list1.Left +
                            this.m_frmRx.uc_rx_list1.Width <
                            this.m_frmRx.ClientSize.Width)
                        {
                            break;
                        }
                    }

                }

                this.m_frmRx.uc_rx_list1.Dock = System.Windows.Forms.DockStyle.Fill;
                this.m_frmRx.uc_rx_list1.loadvalues();
                this.m_frmRx.DisposeOfFormWhenClosing = true;
                this.m_frmRx.MinimizeMainForm = true;
                this.m_frmRx.Left = 0;
                this.m_frmRx.Top = 0;
                this.m_frmRx.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                frmMain.g_sbpInfo.Text = "Ready";
                p_oParentControl.Enabled = false;
                m_frmRx.ParentControl = p_oParentControl;
                this.m_frmRx.Show();



            }
            else
            {
                if (this.m_frmRx.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    this.m_frmRx.WindowState = System.Windows.Forms.FormWindowState.Normal;

                this.m_frmRx.Focus();

            }

        }
        public void StartRxPackageDialog(Control p_oParentControl)
        {
            //check to see if the form has already been loaded
            if (this.IsChildWindowVisible("FVS: Treatment Packages") == false)
            {
                frmMain.g_sbpInfo.Text = "Loading Treatment Package Definitions...Stand By";
               
                this.m_frmRxPackage = new frmDialog(this);
                this.m_frmRxPackage.MaximizeBox = true;
                this.m_frmRxPackage.BackColor = System.Drawing.SystemColors.Control;
                this.m_frmRxPackage.Text = "FVS: Treatment Packages";
               
                this.m_frmRxPackage.Initialize_Rx_Package_User_Control();


                this.m_frmRxPackage.Height = 0;
                this.m_frmRxPackage.Width = 0;
                if (this.m_frmRxPackage.uc_rx_package_list1.Top + this.m_frmRxPackage.uc_rx_package_list1.Height > this.m_frmRxPackage.ClientSize.Height + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmRxPackage.Height = x;
                        if (this.m_frmRxPackage.uc_rx_package_list1.Top +
                            this.m_frmRxPackage.uc_rx_package_list1.Height <
                            this.m_frmRxPackage.ClientSize.Height)
                        {
                            break;
                        }
                    }

                }
                if (this.m_frmRxPackage.uc_rx_package_list1.Left + this.m_frmRxPackage.uc_rx_package_list1.Width > this.m_frmRxPackage.ClientSize.Width + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmRxPackage.Width = x;
                        if (this.m_frmRxPackage.uc_rx_package_list1.Left +
                            this.m_frmRxPackage.uc_rx_package_list1.Width <
                            this.m_frmRxPackage.ClientSize.Width)
                        {
                            break;
                        }
                    }

                }

                this.m_frmRxPackage.uc_rx_package_list1.Dock = System.Windows.Forms.DockStyle.Fill;
                this.m_frmRxPackage.uc_rx_package_list1.loadvalues();
                this.m_frmRxPackage.DisposeOfFormWhenClosing = true;
                this.m_frmRxPackage.MinimizeMainForm = true;
                this.m_frmRxPackage.Left = 0;
                this.m_frmRxPackage.Top = 0;
                this.m_frmRxPackage.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                frmMain.g_sbpInfo.Text = "Ready";
                p_oParentControl.Enabled = false;
                m_frmRxPackage.ParentControl = p_oParentControl;
                this.m_frmRxPackage.Show();



            }
            else
            {
                if (this.m_frmRxPackage.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    this.m_frmRxPackage.WindowState = System.Windows.Forms.FormWindowState.Normal;

                this.m_frmRxPackage.Focus();

            }
        }
        public void StartManageTablesDialog()
        {
            //check to see if the form has already been loaded
            if (this.IsChildWindowVisible("Database: Manage Tables") == false)
            {
                this.m_frmDb = new frmDialog(this);
                this.m_frmDb.MaximizeBox = true;
                this.m_frmDb.BackColor = System.Drawing.SystemColors.Control;
                this.m_frmDb.Text = "Database: Manage Tables";
                this.m_frmDb.MdiParent = this;

                FIA_Biosum_Manager.uc_db p_uc = new uc_db(this.frmProject.uc_project1.txtRootDirectory.Text.Trim());
                if (p_uc.m_intError < 0)
                {
                    this.m_frmDb.Dispose();
                    return;
                }
                this.m_frmDb.Controls.Add(p_uc);
                this.m_frmDb.DbUserControl = p_uc;
                this.m_frmDb.Height = 0;
                this.m_frmDb.Width = 0;
                if (p_uc.Top + p_uc.Height > this.m_frmDb.ClientSize.Height + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmDb.Height = x;
                        if (p_uc.Top +
                            p_uc.Height <
                            this.m_frmDb.ClientSize.Height)
                        {
                            break;
                        }
                    }

                }
                if (p_uc.Left + p_uc.Width > this.m_frmDb.ClientSize.Width + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmDb.Width = x;
                        if (p_uc.Left +
                            p_uc.Width <
                            this.m_frmDb.ClientSize.Width)
                        {
                            break;
                        }
                    }

                }
                p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
                p_uc.loadvalues();


                this.m_frmDb.Left = 0;
                this.m_frmDb.Top = 0;
                this.m_frmDb.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                this.m_frmDb.DisposeOfFormWhenClosing = true;
                this.m_frmDb.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                this.m_frmDb.Show();

            }
            else
            {
                if (this.m_frmDb.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    this.m_frmDb.WindowState = System.Windows.Forms.FormWindowState.Normal;

                this.m_frmDb.Focus();

            }

        }
        public void StartManageSQLiteTablesDialog()
        {
            //check to see if the form has already been loaded
            if (this.IsChildWindowVisible("Database: Manage SQLite Tables") == false)
            {
                this.m_frmDbSQLite = new frmDialog(this);
                this.m_frmDbSQLite.MaximizeBox = true;
                this.m_frmDbSQLite.BackColor = System.Drawing.SystemColors.Control;
                this.m_frmDbSQLite.Text = "Database: Manage SQLite Tables";
                this.m_frmDbSQLite.MdiParent = this;

                FIA_Biosum_Manager.uc_db_sqlite p_uc = new uc_db_sqlite(this.frmProject.uc_project1.txtRootDirectory.Text.Trim());
                if (p_uc.m_intError < 0)
                {
                    this.m_frmDbSQLite.Dispose();
                    return;
                }
                this.m_frmDbSQLite.Controls.Add(p_uc);
                this.m_frmDbSQLite.DbSQLiteUserControl = p_uc;
                this.m_frmDbSQLite.Height = 0;
                this.m_frmDbSQLite.Width = 0;
                if (p_uc.Top + p_uc.Height > this.m_frmDbSQLite.ClientSize.Height + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmDbSQLite.Height = x;
                        if (p_uc.Top +
                            p_uc.Height <
                            this.m_frmDbSQLite.ClientSize.Height)
                        {
                            break;
                        }
                    }

                }
                if (p_uc.Left + p_uc.Width > this.m_frmDbSQLite.ClientSize.Width + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmDbSQLite.Width = x;
                        if (p_uc.Left +
                            p_uc.Width <
                            this.m_frmDbSQLite.ClientSize.Width)
                        {
                            break;
                        }
                    }

                }
                p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
                p_uc.loadvalues();


                this.m_frmDbSQLite.Left = 0;
                this.m_frmDbSQLite.Top = 0;
                this.m_frmDbSQLite.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                this.m_frmDbSQLite.DisposeOfFormWhenClosing = true;
                this.m_frmDbSQLite.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                this.m_frmDbSQLite.Show();

            }
            else
            {
                if (this.m_frmDbSQLite.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    this.m_frmDbSQLite.WindowState = System.Windows.Forms.FormWindowState.Normal;

                this.m_frmDbSQLite.Focus();

            }

        }
        public void StartPSiteDialog(Control p_oParentControl)
        {
            //check to see if the form has already been loaded
            if (this.IsChildWindowVisible("Database: Wood Processing Sites") == false)
            {
                frmMain.g_sbpInfo.Text = "Loading Wood Processing Sites...Stand By";
                this.m_frmPSite = new frmDialog(this);
                this.m_frmPSite.MaximizeBox = true;
                this.m_frmPSite.BackColor = System.Drawing.SystemColors.Control;
                this.m_frmPSite.Text = "Database: Wood Processing Sites";
                FIA_Biosum_Manager.uc_gis_psite p_uc = new uc_gis_psite(this.frmProject.uc_project1.txtRootDirectory.Text.Trim());
                if (p_uc.m_intError < 0)
                {
                    this.m_frmPSite.Dispose();
                    return;
                }
                this.m_frmPSite.Controls.Add(p_uc);
                this.m_frmPSite.ProcessingSiteUserControl = p_uc;
                this.m_frmPSite.Height = 0;
                this.m_frmPSite.Width = 0;
                if (p_uc.Top + p_uc.Height > this.m_frmPSite.ClientSize.Height + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmPSite.Height = x;
                        if (p_uc.Top +
                            p_uc.Height <
                            this.m_frmPSite.ClientSize.Height)
                        {
                            break;
                        }
                    }

                }
                if (p_uc.Left + p_uc.Width > this.m_frmPSite.ClientSize.Width + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmPSite.Width = x;
                        if (p_uc.Left +
                            p_uc.Width <
                            this.m_frmPSite.ClientSize.Width)
                        {
                            break;
                        }
                    }

                }
                p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
                p_uc.loadvalues();


                this.m_frmPSite.Left = 0;
                this.m_frmPSite.Top = 0;
                this.m_frmPSite.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                this.m_frmPSite.DisposeOfFormWhenClosing = true;
                this.m_frmPSite.MinimizeMainForm = true;
                this.m_frmPSite.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                frmMain.g_sbpInfo.Text = "Ready";
                p_oParentControl.Enabled = false;
                m_frmPSite.ParentControl = p_oParentControl;
                this.m_frmPSite.Show();

            }
            else
            {
                if (this.m_frmPSite.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    this.m_frmPSite.WindowState = System.Windows.Forms.FormWindowState.Normal;

                this.m_frmPSite.Focus();

            }

						



        }
        public void StartBiosumProcessorDialog()
        {
            System.Text.StringBuilder strFullPath;
            long lngCount = -1;

            string strProjDir = getProjectDirectory();
            string strScenarioDir = strProjDir.Trim() + "\\processor";
            strFullPath = new System.Text.StringBuilder(strScenarioDir);
            strFullPath.Append("\\");
            string strDbFullPath = strFullPath.ToString() + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            strFullPath.Append(Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsDbFile);
            if (System.IO.File.Exists(strDbFullPath))
            {
                SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(dataMgr.GetConnectionString(strDbFullPath)))
                {
                    con.Open();
                    lngCount = dataMgr.getRecordCount(con, "select count(*) from scenario", "scenario");
                }
            }
            else
            {
                ado_data_access oAdo = new ado_data_access();
                string strConn = oAdo.getMDBConnString(strFullPath.ToString(), "admin", "");
                lngCount = oAdo.getRecordCount(strConn, "select count(*) from scenario", "scenario");
            }

            if (lngCount > 0)
            {
                OpenProcessorScenario("Open", null);
            }
            else
            {
                OpenProcessorScenario("New", null);
            }
        }

        public void StartFVSCreateMdbsDialog()
        {
            // TODO: 
            //check to see if the form has already been loaded
            if (this.IsChildWindowVisible("FVS: Create MDBs") == false)
            {
                frmMain.g_sbpInfo.Text = "Loading Create MDBs feature...Stand By";
                this.m_frmCreateFvsOutputMdbs = new frmDialog(this);
                this.m_frmCreateFvsOutputMdbs.MaximizeBox = true;
                this.m_frmCreateFvsOutputMdbs.BackColor = System.Drawing.SystemColors.Control;
                this.m_frmCreateFvsOutputMdbs.Text = "FVS: Create MDBs";
                FIA_Biosum_Manager.uc_fvs_create_mdbs p_uc = new uc_fvs_create_mdbs(this.frmProject.uc_project1.txtRootDirectory.Text.Trim());
                if (p_uc.m_intError < 0)
                {
                    this.m_frmCreateFvsOutputMdbs.Dispose();
                    return;
                }
                ActivateStandByAnimation(this.WindowState, this.Left, this.Height, this.Width, this.Top);
                this.m_frmCreateFvsOutputMdbs.Controls.Add(p_uc);
                this.m_frmCreateFvsOutputMdbs.FvsCreateMdbsUserControl = p_uc;
                this.m_frmCreateFvsOutputMdbs.Height = 0;
                this.m_frmCreateFvsOutputMdbs.Width = 0;

                if (p_uc.Top + p_uc.Height > this.m_frmCreateFvsOutputMdbs.ClientSize.Height + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmCreateFvsOutputMdbs.Height = x;
                        if (p_uc.Top +
                            p_uc.Height <
                            this.m_frmCreateFvsOutputMdbs.ClientSize.Height)
                        {
                            break;
                        }
                    }

                }
                if (p_uc.Left + p_uc.Width > this.m_frmCreateFvsOutputMdbs.ClientSize.Width + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmCreateFvsOutputMdbs.Width = x;
                        if (p_uc.Left +
                            p_uc.Width <
                            this.m_frmCreateFvsOutputMdbs.ClientSize.Width)
                        {
                            break;
                        }
                    }

                }
                p_uc.Dock = System.Windows.Forms.DockStyle.Fill;

                this.m_frmCreateFvsOutputMdbs.Left = 0;
                this.m_frmCreateFvsOutputMdbs.Top = 0;
                this.m_frmCreateFvsOutputMdbs.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                this.m_frmCreateFvsOutputMdbs.DisposeOfFormWhenClosing = true;
                this.m_frmCreateFvsOutputMdbs.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                frmMain.g_sbpInfo.Text = "Ready";
                this.m_frmCreateFvsOutputMdbs.MinimizeMainForm = true;
                this.m_frmCreateFvsOutputMdbs.ParentControl = this;
                this.Enabled = false;
                this.DeactivateStandByAnimation();
                this.m_frmCreateFvsOutputMdbs.Show(this);

            }
            else
            {
                if (this.m_frmCreateFvsOutputMdbs.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    this.m_frmCreateFvsOutputMdbs.WindowState = System.Windows.Forms.FormWindowState.Normal;

                this.m_frmCreateFvsOutputMdbs.Focus();

            }

        }

        public void StartFVSOutputDataDialog()
        {
            //check to see if the form has already been loaded
            if (this.IsChildWindowVisible("FVS: Process FVS Output") == false)
            {
                frmMain.g_sbpInfo.Text = "Loading FVS Output...Stand By";
                this.m_frmFvsOutput = new frmDialog(this);
                this.m_frmFvsOutput.MaximizeBox = true;
                this.m_frmFvsOutput.BackColor = System.Drawing.SystemColors.Control;
                this.m_frmFvsOutput.Text = "FVS: Process FVS Output";
                FIA_Biosum_Manager.uc_fvs_output p_uc = new uc_fvs_output(this.frmProject.uc_project1.txtRootDirectory.Text.Trim());
                if (p_uc.m_intError < 0)
                {
                    this.m_frmFvsOutput.Dispose();
                    return;
                }
                ActivateStandByAnimation(this.WindowState, this.Left, this.Height, this.Width, this.Top);
                this.m_frmFvsOutput.Controls.Add(p_uc);
                this.m_frmFvsOutput.FvsOutProcessorInUserControl = p_uc;
                this.m_frmFvsOutput.Height = 0;
                this.m_frmFvsOutput.Width = 0;

                if (p_uc.Top + p_uc.Height > this.m_frmFvsOutput.ClientSize.Height + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmFvsOutput.Height = x;
                        if (p_uc.Top +
                            p_uc.Height <
                            this.m_frmFvsOutput.ClientSize.Height)
                        {
                            break;
                        }
                    }

                }
                if (p_uc.Left + p_uc.Width > this.m_frmFvsOutput.ClientSize.Width + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmFvsOutput.Width = x;
                        if (p_uc.Left +
                            p_uc.Width <
                            this.m_frmFvsOutput.ClientSize.Width)
                        {
                            break;
                        }
                    }

                }
                p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
                p_uc.loadvalues();


                this.m_frmFvsOutput.Left = 0;
                this.m_frmFvsOutput.Top = 0;
                this.m_frmFvsOutput.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                this.m_frmFvsOutput.DisposeOfFormWhenClosing = true;
                this.m_frmFvsOutput.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                frmMain.g_sbpInfo.Text = "Ready";
                this.m_frmFvsOutput.MinimizeMainForm = true;
                this.m_frmFvsOutput.ParentControl = this;
                this.Enabled = false;
                this.DeactivateStandByAnimation();
                this.m_frmFvsOutput.Show(this);

            }
            else
            {
                if (this.m_frmFvsOutput.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    this.m_frmFvsOutput.WindowState = System.Windows.Forms.FormWindowState.Normal;

                this.m_frmFvsOutput.Focus();

            }

        }
        public void StartFVSInputDataDialog()
        {
            //check to see if the form has already been loaded
            if (this.IsChildWindowVisible("FVS: Create FVS Input") == false)
            {
                ActivateStandByAnimation(this.WindowState, this.Left, this.Height, this.Width, this.Top);
                frmMain.g_sbpInfo.Text = "Loading FVS Input Data...Stand By";
                
                this.m_frmFvsInput = new frmDialog(this);
                this.m_frmFvsInput.MaximizeBox = true;
                this.m_frmFvsInput.BackColor = System.Drawing.SystemColors.Control;
                this.m_frmFvsInput.Text = "FVS: Create FVS Input";
               
                FIA_Biosum_Manager.uc_fvs_input p_uc = new uc_fvs_input();
                p_uc.ReferenceMainForm = this;
                p_uc.ReferenceParentDialogForm = this.m_frmFvsInput;
                this.m_frmFvsInput.Controls.Add(p_uc);
                this.m_frmFvsInput.FVSInputUserControl = p_uc;

                p_uc.strProjectDirectory = this.frmProject.uc_project1.txtRootDirectory.Text.Trim();
                p_uc.strProjectId = this.frmProject.uc_project1.txtProjectId.Text.Trim();




                this.m_frmFvsInput.Height = 0;
                this.m_frmFvsInput.Width = 0;
                if (p_uc.Top + p_uc.Height > this.m_frmFvsInput.ClientSize.Height + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmFvsInput.Height = x;
                        if (p_uc.Top +
                            p_uc.Height <
                            this.m_frmFvsInput.ClientSize.Height)
                        {
                            break;
                        }
                    }

                }
                if (p_uc.Left + p_uc.Width > this.m_frmFvsInput.ClientSize.Width + 2)
                {
                    for (int x = 1; ; x++)
                    {
                        this.m_frmFvsInput.Width = x;
                        if (p_uc.Left +
                            p_uc.Width <
                            this.m_frmFvsInput.ClientSize.Width)
                        {
                            break;
                        }
                    }

                }
                p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
                p_uc.loadvalues();


                this.m_frmFvsInput.Left = 0;
                this.m_frmFvsInput.Top = 0;
                this.m_frmFvsInput.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                this.m_frmFvsInput.DisposeOfFormWhenClosing = true;
                
                this.m_frmFvsInput.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                frmMain.g_sbpInfo.Text = "Ready";
                this.m_frmFvsInput.MinimizeMainForm = true;
                this.Enabled = false;
                this.m_frmFvsInput.ParentControl = this;
                this.DeactivateStandByAnimation();
                this.m_frmFvsInput.Show(this);

            }
            else
            {
                if (this.m_frmFvsInput.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    this.m_frmFvsInput.WindowState = System.Windows.Forms.FormWindowState.Normal;

                this.m_frmFvsInput.Focus();

            }
        }
		private void tlbMain_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			
			switch (e.Button.Text)
			{
				case "Open":
					this.mnuFileOpenProject_Click(sender, e);
					break;
                case "Save":
                    this.SaveAll(false);
					break;
                case "Project":
					this.mnuViewProject_Click(sender,e);
					break;
                case "Notes":
					this.mnuViewNotes_Click(sender,e);
                    break;
                case "Links":
					this.mnuViewLinks_Click(sender,e);
					break;
                case "Contacts":
					this.mnuViewContacts_Click(sender,e);
                    break;
			}
		}

		private void mnuViewProject_Click(object sender, System.EventArgs e)
		{
			
			this.frmProject.uc_project_document_links1.Visible=false;
			this.frmProject.uc_project_notes1.Visible=false;
			this.frmProject.uc_contact_list1.Visible=false;
			this.frmProject.uc_project1.Visible=true;
			this.frmProject.uc_project1.m_strAction="VIEW";
			this.frmProject.uc_project1.lblTitle.Text = "Project Properties";
			
			this.tlbMain.Buttons[1].Enabled=true;
            resizeProjectForm(this);
			
			this.frmProject.Visible = true;
			if (this.frmProject.WindowState==System.Windows.Forms.FormWindowState.Minimized)
			{
				this.frmProject.WindowState = System.Windows.Forms.FormWindowState.Normal;
			}
			this.frmProject.uc_project1.resize_uc_project();
			this.frmProject.uc_project1.m_oResizeForm.ControlToResize = frmProject;
			
			this.frmProject.uc_project1.m_oResizeForm.ResizeControl();

			this.frmProject.Focus();
			
			
		}
		public void OpenProject(string strNewProjectDirectory, string strNewProjectFile)
		{
            
			//check to make sure this project is not already open
			//lets see if this project is already open
			FIA_Biosum_Manager.utils p_oUtils = new FIA_Biosum_Manager.utils();
			if (p_oUtils.FindWindowLike((IntPtr)0, "FIA Biosum Manager (" + this.frmProject.uc_project1.m_strNewProjectId + ")","*",true,true) > 0)
			{
				MessageBox.Show("!!Project Already Open!!","Project Open",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}
			p_oUtils = null;

            

			version_control oVersCtl = new version_control();
			
			if (this.frmProject.uc_project1.m_intError == 0  && m_ProjectOpen == false) 
			{
                frmProject.uc_project1.m_strDebugFile = frmMain.g_oEnv.strTempDir + @"\FIA_Biosum_DebugLog_" + this.frmProject.uc_project1.m_strNewProjectId.Trim() + "_" + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".txt";
				frmMain.g_sbpInfo.Text = "Loading Project...Stand By";
				this.frmProject.uc_project1.m_strProjectDirectory = strNewProjectDirectory; 
				this.frmProject.uc_project1.m_strProjectFile = strNewProjectFile; 
				this.frmProject.uc_project1.lblTitle.Text = "Project Properties";
				this.Text = "Fia Biosum Manager (" + this.frmProject.uc_project1.m_strNewProjectId.Trim() + ")";
				this.frmProject.uc_project1.txtProjectId.Text = this.frmProject.uc_project1.m_strNewProjectId;
				this.frmProject.uc_project1.txtName.Text = this.frmProject.uc_project1.m_strNewName;
				this.frmProject.uc_project1.txtDate.Text = this.frmProject.uc_project1.m_strNewDate;
				this.frmProject.uc_project1.txtCompany.Text = this.frmProject.uc_project1.m_strNewCompany;
				this.frmProject.uc_project1.txtDescription.Text = this.frmProject.uc_project1.m_strNewDescription ;
				this.frmProject.uc_project1.txtShared.Text = this.frmProject.uc_project1.m_strNewShared;
				this.frmProject.uc_project1.txtRootDirectory.Text = this.frmProject.uc_project1.m_strNewRootDirectory;
				this.frmProject.uc_project1.m_strProjectId = this.frmProject.uc_project1.m_strNewProjectId;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "=====================   Opening FIA BIOSUM Project   =====================\r\n");
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "**Project Properties**\r\n");
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "Project ID:                 " + frmProject.uc_project1.txtProjectId.Text.Trim() + "\r\n");
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "Project Name:               " + frmProject.uc_project1.txtName.Text.Trim() + "\r\n");
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "Project Root Directory:     " + frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\r\n");
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "Project File:               " + frmProject.uc_project1.m_strProjectFile + "\r\n");
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "Project Date:               " + frmProject.uc_project1.txtDate.Text.Trim() + "\r\n");
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "Current Date/Time:          " + DateTime.Now.ToString() + "\r\n");
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "Company:                    " + frmProject.uc_project1.txtCompany.Text.Trim() + "\r\n");
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "Shared:                     " + frmProject.uc_project1.txtShared.Text.Trim() + "\r\n");
                    frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "Description:                " + frmProject.uc_project1.txtDescription.Text.Trim() + "\r\n");
                }


				this.frmProject.uc_project1.SetProjectPathEnvironmentVariables();
				if (frmProject.uc_project1.m_strAction != "NEW")
				{
					oVersCtl.ReferenceMainForm=this;
					oVersCtl.ReferenceProjectDirectory=this.frmProject.uc_project1.m_strProjectDirectory;
                    // Warn if application version < project version
                    string strProjVersionFile = this.frmProject.uc_project1.m_strProjectDirectory.Trim() + "\\application.version";
                    string strProjVersion = "";
                    if (System.IO.File.Exists(strProjVersionFile))
                    {
                        using (System.IO.StreamReader reader = new System.IO.StreamReader(strProjVersionFile))
                        {
                            string strTemp = reader.ReadLine();
                            if (strTemp != null)
                            {
                                strProjVersion = strTemp.Trim();
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(strProjVersion))
                    {
                        int retVal = oVersCtl.VersionCompare(frmMain.g_strAppVer, strProjVersion);
                        if (retVal < 0)
                        {
                            System.Text.StringBuilder sb = new System.Text.StringBuilder();
                            sb.Append("WARNING: You are trying to open a project from v" + strProjVersion.Trim() + " with an older version of BioSum! ");
                            sb.Append("This is not recommended. Click 'Yes' if you wish to continue.");
                            System.Windows.Forms.DialogResult res =
                                System.Windows.Forms.MessageBox.Show(sb.ToString(), "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo);
                            if (res != System.Windows.Forms.DialogResult.Yes)
                            {
                                return;
                            }
                        }
                    }
                    oVersCtl.PerformVersionCheck();
				}
              

				btnDB.ForeColor = Color.Red;
				this.btnContacts.Enabled=true;
				this.btnOptimizer.Enabled = true;
				this.btnDB.Enabled=false;
				this.btnFVS.Enabled=true;
				this.btnProcessor.Enabled=true;
				this.m_pnlCurrent.Enabled=true;
				this.btnContacts.Enabled=true;
				this.btnNotes.Enabled=true;
				this.btnProject.Enabled=true;
				this.btnSave.Enabled=true;
                this.mnuToolsProjectRootFolder.Enabled = true;
				if (this.frmProject.uc_project1.txtPersonal.Text.Trim().Length == 0 &&
					this.frmProject.uc_project1.txtShared.Text.Trim().Length == 0)
				{
					this.btnLinks.Enabled = false;
					this.btnNotes.Enabled=false;
				}
				else 
				{
					this.btnLinks.Enabled=true;
					this.btnNotes.Enabled=true;
				}
				this.mnuView.Enabled=true;
				
				this.mnuFileSaveProject.Enabled=true;
				m_ProjectOpen = true;
				this.recentfiles();

                
			}
			else if (this.frmProject.uc_project1.m_intError == 0)
			{
				frmMain frmTemp = new frmMain();
                frmTemp.frmProject.uc_project1.m_strDebugFile = frmMain.g_oEnv.strTempDir + @"\FIA_Biosum_DebugLog_" + this.frmProject.uc_project1.m_strNewProjectId.Trim() + "_" + String.Format("{0:yyyyMMdd}", DateTime.Now) + ".txt";
				frmTemp.frmProject.uc_project1.m_strProjectDirectory = strNewProjectDirectory; 
				frmTemp.frmProject.uc_project1.m_strProjectFile = strNewProjectFile; 
				frmTemp.frmProject.uc_project1.lblTitle.Text = "Project Properties";
				frmTemp.Text = "FIA Biosum Manager (" + this.frmProject.uc_project1.m_strNewProjectId + ")";
				frmTemp.frmProject.uc_project1.txtProjectId.Text = this.frmProject.uc_project1.m_strNewProjectId;
				frmTemp.frmProject.uc_project1.txtName.Text = this.frmProject.uc_project1.m_strNewName;
				frmTemp.frmProject.uc_project1.txtDate.Text = this.frmProject.uc_project1.m_strNewDate;
				frmTemp.frmProject.uc_project1.txtCompany.Text = this.frmProject.uc_project1.m_strNewCompany;
				frmTemp.frmProject.uc_project1.txtDescription.Text = this.frmProject.uc_project1.m_strNewDescription ;
				frmTemp.frmProject.uc_project1.txtShared.Text = this.frmProject.uc_project1.m_strNewShared;
				frmTemp.frmProject.uc_project1.txtRootDirectory.Text = this.frmProject.uc_project1.m_strNewRootDirectory;
				frmTemp.frmProject.uc_project1.m_strProjectId = this.frmProject.uc_project1.m_strNewProjectId;

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(frmTemp.frmProject.uc_project1.m_strDebugFile, "=====================   Opening FIA BIOSUM Project   =====================\r\n");
                    frmMain.g_oUtils.WriteText(frmTemp.frmProject.uc_project1.m_strDebugFile, "**Project Properties**\r\n");
                    frmMain.g_oUtils.WriteText(frmTemp.frmProject.uc_project1.m_strDebugFile, "Project ID:                 " + frmProject.uc_project1.txtProjectId.Text.Trim() + "\r\n");
                    frmMain.g_oUtils.WriteText(frmTemp.frmProject.uc_project1.m_strDebugFile, "Project Name:               " + frmProject.uc_project1.txtName.Text.Trim() + "\r\n");
                    frmMain.g_oUtils.WriteText(frmTemp.frmProject.uc_project1.m_strDebugFile, "Project Root Directory:     " + frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\r\n");
                    frmMain.g_oUtils.WriteText(frmTemp.frmProject.uc_project1.m_strDebugFile, "Project File:               " + frmProject.uc_project1.m_strProjectFile + "\r\n");
                    frmMain.g_oUtils.WriteText(frmTemp.frmProject.uc_project1.m_strDebugFile, "Project Date:               " + frmProject.uc_project1.txtDate.Text.Trim() + "\r\n");
                    frmMain.g_oUtils.WriteText(frmTemp.frmProject.uc_project1.m_strDebugFile, "Current Date/Time           " + DateTime.Now.ToString() + "\r\n") ;
                    frmMain.g_oUtils.WriteText(frmTemp.frmProject.uc_project1.m_strDebugFile, "Company:                    " + frmProject.uc_project1.txtCompany.Text.Trim() + "\r\n");
                    frmMain.g_oUtils.WriteText(frmTemp.frmProject.uc_project1.m_strDebugFile, "Shared:                     " + frmProject.uc_project1.txtShared.Text.Trim() + "\r\n");
                    frmMain.g_oUtils.WriteText(frmTemp.frmProject.uc_project1.m_strDebugFile, "Description:                " + frmProject.uc_project1.txtDescription.Text.Trim() + "\r\n");
                }

				frmTemp.frmProject.uc_project1.SetProjectPathEnvironmentVariables();
				if (frmTemp.frmProject.uc_project1.m_strAction != "NEW")
				{
					oVersCtl.ReferenceMainForm=frmTemp;
					oVersCtl.ReferenceProjectDirectory=frmTemp.frmProject.uc_project1.m_strProjectDirectory;
					oVersCtl.PerformVersionCheck();
				}
              
				btnDB.ForeColor = Color.Red;
				frmTemp.btnContacts.Enabled=true;
				frmTemp.btnOptimizer.Enabled = true;
				frmTemp.btnDB.Enabled=false;
			
				frmTemp.btnFVS.Enabled=true;
				frmTemp.btnProcessor.Enabled=true;
				frmTemp.m_pnlCurrent.Enabled=true;
				frmTemp.btnContacts.Enabled=true;
				frmTemp.btnNotes.Enabled=true;
				frmTemp.btnProject.Enabled=true;
				frmTemp.btnSave.Enabled=true;
                mnuToolsProjectRootFolder.Enabled = true;

				if (frmTemp.frmProject.uc_project1.txtPersonal.Text.Trim().Length == 0 &&
					frmTemp.frmProject.uc_project1.txtShared.Text.Trim().Length == 0)
				{
					frmTemp.btnLinks.Enabled = false;
				}
				else 
				{
					frmTemp.btnLinks.Enabled=true;
				}
				
				frmTemp.mnuView.Enabled=true;
				
				frmTemp.mnuFileSaveProject.Enabled=true;
				frmTemp.m_ProjectOpen = true;
				frmTemp.recentfiles();
				frmTemp.Show();
			}
		    this.frmProject.uc_project1.m_strNewProjectFile = "";
		    this.frmProject.uc_project1.m_strNewProjectDirectory = "";
		    this.frmProject.uc_project1.m_strNewProjectId="";
            this.frmProject.uc_project1.m_strNewName="";
		    this.frmProject.uc_project1.m_strNewDate="";
		    this.frmProject.uc_project1.m_strNewCompany="";
		    this.frmProject.uc_project1.m_strNewDescription="";
            this.frmProject.uc_project1.m_strNewShared="";
		    this.frmProject.uc_project1.m_strNewRootDirectory="";
			this.frmProject.uc_project1.m_strNewProjectVersion="";
			frmMain.g_sbpInfo.Text = "Ready";
			

		}

		private void mnuFileNewProject_Click(object sender, System.EventArgs e)
		{

			
			if (this.m_ProjectOpen == false) 
			{
				this.frmProject.uc_project1.lblTitle.Text = "New Project";
				this.frmProject.uc_project1.New_Project();
                resizeProjectForm(this);
			}
			else 
			{
				frmMain frmTemp = new frmMain();
				frmTemp.m_ProjectOpen = false;
				frmTemp.frmProject.uc_project1.lblTitle.Text = "New Project";
				frmTemp.frmProject.uc_project1.New_Project();
                resizeProjectForm(frmTemp);
				frmTemp.Show();
			}
		 
		}
		private void recentfiles()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(this.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//frmMain.recentfiles \r\n");
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }
			string[]  strCurrentMenuItem = new string[5];
			string[]  strNewMenuItem     = new string[5];
			int intCount=0;
			int a=0;

			for (a=0; a <= 4; a++)
			{
				strCurrentMenuItem[a]="";
				strNewMenuItem[a]="";
			}


			try 
			{
				// Create an instance of StreamReader to read from a file.
				// The using statement also closes the StreamReader.
				System.IO.StreamReader sr = new System.IO.StreamReader(this.m_oEnv.strApplicationDataDirectory + "\\recent.dat");
			
				String line;
				int intMenuCount=0;
	
				while ((line = sr.ReadLine()) != null) 
				{
					switch (intMenuCount)
					{
						case 0:
							strCurrentMenuItem[0] = line;
							break;
						case 1:
							strCurrentMenuItem[1] = line;
							break;
						case 2:
							strCurrentMenuItem[2] = line;
							break;
						case 3:
							strCurrentMenuItem[3] = line;
							break;
						case 4:
							strCurrentMenuItem[4] = line;
							break;
					}
					intMenuCount++;
				}
				sr.Close();
				
			
			}
			catch  
			{
				
			}
			

            strNewMenuItem[0] = this.frmProject.uc_project1.m_strProjectDirectory + "\\db\\" + 
			                    this.frmProject.uc_project1.m_strProjectFile;

			for (a=0; a <= 4; a++)
			{
				if (strCurrentMenuItem[a].Trim().Length > 0) 
				{
					if (strNewMenuItem[0].Trim().ToUpper() == 
						strCurrentMenuItem[a].Trim().ToUpper())
					{
					}
					else 
					{
						intCount++;
						strNewMenuItem[intCount] = strCurrentMenuItem[a];
						if (intCount == 4) break;
					}
				}
			}

			System.IO.StreamWriter sw = new System.IO.StreamWriter((this.m_oEnv.strApplicationDataDirectory + "\\recent.dat"));

			this.mnuFileRecentProjects1.Text = strNewMenuItem[0];
			sw.WriteLine(strNewMenuItem[0]);
			this.mnuFileRecentProjects1.Enabled=true;
			this.mnuFileRecentProjects1.Visible=true;
			
			if (strNewMenuItem[1].Trim().Length > 0)
			{
				this.mnuFileRecentProjects2.Text = strNewMenuItem[1];
				sw.WriteLine(strNewMenuItem[1]);
				this.mnuFileRecentProjects2.Enabled=true;
				this.mnuFileRecentProjects2.Visible=true;
			}

			if (strNewMenuItem[2].Trim().Length > 0)
			{
				this.mnuFileRecentProjects3.Text = strNewMenuItem[2];
				sw.WriteLine(strNewMenuItem[2]);
				this.mnuFileRecentProjects3.Enabled=true;
				this.mnuFileRecentProjects3.Visible=true;
			}

			if (strNewMenuItem[3].Trim().Length > 0)
			{
				this.mnuFileRecentProjects4.Text = strNewMenuItem[3];
				sw.WriteLine(strNewMenuItem[3]);
				this.mnuFileRecentProjects4.Enabled=true;
				this.mnuFileRecentProjects4.Visible=true;
			}

			if (strNewMenuItem[4].Trim().Length > 0)
			{
				this.mnuFileRecentProjects5.Text = strNewMenuItem[4];
				sw.WriteLine(strNewMenuItem[4]);
				this.mnuFileRecentProjects5.Enabled=true;
				this.mnuFileRecentProjects5.Visible=true;
			}
            sw.Close();
			sw = null;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//frmMain.recentfiles: Leaving \r\n");
		}

		private void mnuFileRecentProjects1_Click(object sender, System.EventArgs e)
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(this.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//frmMain.mnuFileRecentProjects1_Click \r\n");
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }
			this.frmProject.uc_project1.Open_Project_No_Dialog(this.mnuFileRecentProjects1.Text);
			if (this.frmProject.uc_project1.m_intError==0) 
				this.OpenProject(this.frmProject.uc_project1.m_strNewProjectDirectory,
					             this.frmProject.uc_project1.m_strNewProjectFile);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//frmMain.mnuFileRecentProjects1_Click: Leaving \r\n");
               
           
			
		}

		private void mnuFileRecentProjects2_Click(object sender, System.EventArgs e)
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(this.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//frmMain.mnuFileRecentProjects2_Click \r\n");
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }
			this.frmProject.uc_project1.Open_Project_No_Dialog(this.mnuFileRecentProjects2.Text);
			if (this.frmProject.uc_project1.m_intError==0) 
				this.OpenProject(this.frmProject.uc_project1.m_strNewProjectDirectory,
					this.frmProject.uc_project1.m_strNewProjectFile);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//frmMain.mnuFileRecentProjects2_Click: Leaving \r\n");
		}

		private void mnuFileRecentProjects3_Click(object sender, System.EventArgs e)
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(this.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//frmMain.mnuFileRecentProjects3_Click \r\n");
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }
			this.frmProject.uc_project1.Open_Project_No_Dialog(this.mnuFileRecentProjects3.Text);
			if (this.frmProject.uc_project1.m_intError==0) 
				this.OpenProject(this.frmProject.uc_project1.m_strNewProjectDirectory,
					this.frmProject.uc_project1.m_strNewProjectFile);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//frmMain.mnuFileRecentProjects3_Click: Leaving \r\n");
		}

		private void mnuFileRecentProjects4_Click(object sender, System.EventArgs e)
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(this.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//frmMain.mnuFileRecentProjects4_Click \r\n");
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }   
			this.frmProject.uc_project1.Open_Project_No_Dialog(this.mnuFileRecentProjects4.Text);
			if (this.frmProject.uc_project1.m_intError==0) 
				this.OpenProject(this.frmProject.uc_project1.m_strNewProjectDirectory,
					this.frmProject.uc_project1.m_strNewProjectFile);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//frmMain.mnuFileRecentProjects4_Click: Leaving \r\n");
		}

		private void mnuFileRecentProjects5_Click(object sender, System.EventArgs e)
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(this.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//frmMain.mnuFileRecentProjects5_Click \r\n");
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }
			this.frmProject.uc_project1.Open_Project_No_Dialog(this.mnuFileRecentProjects5.Text);
			if (this.frmProject.uc_project1.m_intError==0) 
				this.OpenProject(this.frmProject.uc_project1.m_strNewProjectDirectory,
					this.frmProject.uc_project1.m_strNewProjectFile);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmProject.uc_project1.m_strDebugFile, "//frmMain.mnuFileRecentProjects5_Click: Leaving \r\n");
		}

		private void mnuViewLinks_Click(object sender, System.EventArgs e)
		{
			int intAvailHt=0;;
			int intAvailWd=0;
			this.frmProject.uc_project_notes1.Visible=false;
			this.frmProject.uc_project1.Visible=false;
            this.frmProject.uc_scenario1.Visible=false;
			this.frmProject.uc_contact_list1.Visible=false;
			this.frmProject.uc_project_document_links1.loadvalues(this.frmProject.uc_project1.txtShared.Text,this.frmProject.uc_project1.txtPersonal.Text,false);
			this.frmProject.uc_project_document_links1.Visible=true;
			this.frmProject.Visible = true;
			if (this.frmProject.WindowState==System.Windows.Forms.FormWindowState.Minimized)
			{
				this.frmProject.WindowState = System.Windows.Forms.FormWindowState.Normal;
			}
			else
			{
				
			}

			this.frmProject.Focus();
			intAvailWd = this.ClientSize.Width - this.grpboxLeft.Left - this.grpboxLeft.Width - 20;
			intAvailHt = this.ClientSize.Height - this.tlbMain.Top - this.tlbMain.Height - 20;
			this.frmProject.Height = intAvailHt; 
			this.frmProject.Width = intAvailWd; 
			this.frmProject.Left = 0;
			this.frmProject.Top = 0;


		}

		private void frmMain_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{

            frmMain.SaveSettings();
			this.SaveAll(true);

			//delete any temporary files
			string[] strFiles = new string[100];
			strFiles = System.IO.Directory.GetFiles(this.m_oEnv.strTempDir,"fia_biosum_*.*");
			for (int x=0; x <= strFiles.GetUpperBound(0); x++)
			{
				if (strFiles[x].Trim().Length > 0)
				{
					try 
					{
					
						System.IO.File.Delete(strFiles[x]);
					}  
					catch 
					{
					}
				}
			}

            if (this.frmProject.uc_project1.txtRootDirectory.Text.Trim().Length > 0)
            {
                                
                //delete any temporary FRCS files
                if (System.IO.Directory.Exists(frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\FRCS"))
                {
                    strFiles = new string[100];
                    strFiles = System.IO.Directory.GetFiles(frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\FRCS", "fia_biosum_*.xls");
                    for (int x = 0; x <= strFiles.GetUpperBound(0); x++)
                    {
                        if (strFiles[x].Trim().Length > 0)
                        {
                            try
                            {

                                System.IO.File.Delete(strFiles[x]);
                            }
                            catch
                            {
                            }
                        }
                    }
                }
                 
            }
			


		}
		
        public static void SaveSettings()
        {
            //save settings file
            if (!System.IO.Directory.Exists(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum"))
                System.IO.Directory.CreateDirectory(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum");

            if (System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg"))
                System.IO.File.Delete(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg");
            //
            //GRIDVIEW 
            //
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "[GRIDVIEW]\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "BackgroundColor_A=" + frmMain.g_oGridViewBackgroundColor.A.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "BackgroundColor_R=" + frmMain.g_oGridViewBackgroundColor.R.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "BackgroundColor_G=" + frmMain.g_oGridViewBackgroundColor.G.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "BackgroundColor_B=" + frmMain.g_oGridViewBackgroundColor.B.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "RowBackgroundColor_A=" + frmMain.g_oGridViewRowBackgroundColor.A.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "RowBackgroundColor_R=" + frmMain.g_oGridViewRowBackgroundColor.R.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "RowBackgroundColor_G=" + frmMain.g_oGridViewRowBackgroundColor.G.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "RowBackgroundColor_B=" + frmMain.g_oGridViewRowBackgroundColor.B.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "RowForegroundColor_A=" + frmMain.g_oGridViewRowForegroundColor.A.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "RowForegroundColor_R=" + frmMain.g_oGridViewRowForegroundColor.R.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "RowForegroundColor_G=" + frmMain.g_oGridViewRowForegroundColor.G.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "RowForegroundColor_B=" + frmMain.g_oGridViewRowForegroundColor.B.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "AlternatingRowBackgroundColor_A=" + frmMain.g_oGridViewAlternateRowBackgroundColor.A.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "AlternatingRowBackgroundColor_R=" + frmMain.g_oGridViewAlternateRowBackgroundColor.R.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "AlternatingRowBackgroundColor_G=" + frmMain.g_oGridViewAlternateRowBackgroundColor.G.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "AlternatingRowBackgroundColor_B=" + frmMain.g_oGridViewAlternateRowBackgroundColor.B.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "SelectedRowBackgroundColor_A=" + frmMain.g_oGridViewSelectedRowBackgroundColor.A.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "SelectedRowBackgroundColor_R=" + frmMain.g_oGridViewSelectedRowBackgroundColor.R.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "SelectedRowBackgroundColor_G=" + frmMain.g_oGridViewSelectedRowBackgroundColor.G.ToString().Trim() + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "SelectedRowBackgroundColor_B=" + frmMain.g_oGridViewSelectedRowBackgroundColor.B.ToString().Trim() + "\r\n");


            if (frmMain.g_oGridViewFont != null)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "FontName=" + frmMain.g_oGridViewFont.Name.Trim() + "\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "FontSize=" + Convert.ToInt16(frmMain.g_oGridViewFont.Size).ToString().Trim() + "\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "FontStyle=" + Convert.ToInt16(frmMain.g_oGridViewFont.Style).ToString().Trim() + "\r\n");
            }
            //
            //DEBUG
            //
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "[DEBUG]\r\n");
            if (frmMain.g_bDebug) frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "Debug=Y\r\n");
            else frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "Debug=N\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "Level=" + frmMain.g_intDebugLevel.ToString().Trim() + "\r\n");
            //
            //SUPPRESS TABLE RECORD COUNTS
            //
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "[SUPPRESS TABLE RECORD COUNTS]\r\n");
            if (frmMain.g_bSuppressFVSInputTableRowCount) frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "FVSInputForm=Y\r\n");
            else frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "FVSInputForm=N\r\n");
            if (frmMain.g_bSuppressFVSOutputTableRowCount) frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "FVSOutputForm=Y\r\n");
            else frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "FVSOutputForm=N\r\n");
            if (frmMain.g_bSuppressProcessorScenarioTableRowCount) frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "ProcessorScenarioForm=Y\r\n");
            else frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "ProcessorScenarioForm=N\r\n");
            //
            //OPCOST
            //
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "[OPCOST]\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "RFile=" + frmMain.g_strRDirectory + "\r\n");
            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "OPCOSTFile=" + frmMain.g_strOPCOSTDirectory + "\r\n");

            frmMain.g_oUtils.WriteText(frmMain.g_oEnv.strApplicationDataDirectory.Trim() + "\\FIABiosum\\settings.cfg", "END");
        }
        private void SaveAll(bool p_bPrompt)
		{
			    bool bPromptMsg=false;
				DialogResult result;
				foreach (Form child in this.MdiChildren)

				{
                    if (child.Text.IndexOf("Treatment Optimizer: Optimization") >= 0) 
					{
						/*************************************************************
						 **cast the child form to get a reference to its controls,
						 **properties and methods
						 *************************************************************/
						FIA_Biosum_Manager.frmOptimizerScenario  temp= ((FIA_Biosum_Manager.frmOptimizerScenario)child);
						if (temp.m_bSave == true) 
						{
							if (bPromptMsg==false && p_bPrompt)
							{
								result = MessageBox.Show("Save Changes to " + child.Text.Trim() + " Y/N","FIA Biosum",System.Windows.Forms.MessageBoxButtons.YesNo,System.Windows.Forms.MessageBoxIcon.Question);
								if (result == System.Windows.Forms.DialogResult.No)
								{
									break;
								}
								else bPromptMsg=true;
							}
							temp.SaveRuleDefinitions();
							
						}
						temp = null;

					}
                    else if (child.Text.IndexOf("Processor: Scenario") >= 0)
                    {
                        /*************************************************************
                         **cast the child form to get a reference to its controls,
                         **properties and methods
                         *************************************************************/
                        FIA_Biosum_Manager.frmProcessorScenario temp = ((FIA_Biosum_Manager.frmProcessorScenario)child);
                        if (temp.m_bSave == true)
                        {
                            if (bPromptMsg == false && p_bPrompt)
                            {
                                result = MessageBox.Show("Save Changes to " + child.Text.Trim() + " Y/N", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                                if (result == System.Windows.Forms.DialogResult.No)
                                {
                                    break;
                                }
                                else bPromptMsg = true;
                            }
                            temp.SaveRuleDefinitions();

                        }
                        temp = null;

                    }
                    else if (child.Text.IndexOf("Project") >= 0)
                    {
                        //cast the child form to get a reference to its controls, properties and methods
                        FIA_Biosum_Manager.frmDialog temp = child as FIA_Biosum_Manager.frmDialog;
                        if (temp != null)
                        {
                        if (temp.uc_project1.btnSave.Enabled == true)
                        {
                            if (bPromptMsg == false && p_bPrompt)
                            {
                                result = MessageBox.Show("Save Changes to " + child.Text.Trim() + " Y/N", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                                if (result == System.Windows.Forms.DialogResult.No)
                                {
                                    break;
                                }
                                else bPromptMsg = true;
                            }
                            temp.uc_project1.SaveProjectProperties();
                        }
                        if (temp.uc_project_notes1.btnSave.Enabled == true)
                        {
                            if (bPromptMsg == false && p_bPrompt)
                            {
                                result = MessageBox.Show("Save Changes to " + child.Text.Trim() + " Y/N", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                                if (result == System.Windows.Forms.DialogResult.No)
                                {
                                    break;
                                }
                                else bPromptMsg = true;
                            }
                            temp.uc_project_notes1.savevalues();
                        }
                        if (temp.uc_contact_list1.btnSave.Enabled == true)
                        {
                            if (bPromptMsg == false && p_bPrompt)
                            {
                                result = MessageBox.Show("Save Changes to " + child.Text.Trim() + " Y/N", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                                if (result == System.Windows.Forms.DialogResult.No)
                                {
                                    break;
                                }
                                else bPromptMsg = true;
                            }
                            temp.uc_contact_list1.savevalues();
                        }
                        temp = null;
                    }
                    }

                    else if (child.Text.IndexOf("Treatment Optimizer: Edit Harvest Costs") >= 0)
                    {
                        FIA_Biosum_Manager.frmGridView temp = ((FIA_Biosum_Manager.frmGridView)child);
                        if (bPromptMsg == false && p_bPrompt)
                        {
                            result = MessageBox.Show("Save Changes to " + child.Text.Trim() + " Y/N", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                            if (result == System.Windows.Forms.DialogResult.No)
                            {
                                break;
                            }
                            else bPromptMsg = true;
                        }
                        temp.SaveAll();
                    }
                    else if (child.Text.IndexOf("FVS: Treatments") >= 0)
                    {
                        FIA_Biosum_Manager.frmDialog temp = ((FIA_Biosum_Manager.frmDialog)child);
                        if (temp.uc_rx_list1.btnSave.Enabled)
                        {
                            if (bPromptMsg == false && p_bPrompt)
                            {
                                result = MessageBox.Show("Save Changes to " + child.Text.Trim() + " Y/N", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                                if (result == System.Windows.Forms.DialogResult.No)
                                {
                                    break;
                                }
                                else bPromptMsg = true;
                            }
                            temp.uc_rx_list1.savevalues();
                        }

                    }
                    else if (child.Text.IndexOf("Processor: Tree Diameter Groups") >= 0)
                    {
                        FIA_Biosum_Manager.frmDialog temp = ((FIA_Biosum_Manager.frmDialog)child);
                        if (temp.uc_processor_scenario_tree_diam_groups_list1.btnSave.Enabled)
                        {
                            if (bPromptMsg == false && p_bPrompt)
                            {
                                result = MessageBox.Show("Save Changes to " + child.Text.Trim() + " Y/N", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                                if (result == System.Windows.Forms.DialogResult.No)
                                {
                                    break;
                                }
                                else bPromptMsg = true;
                            }
                            temp.uc_processor_scenario_tree_diam_groups_list1.savevalues();
                        }
                    }
				}


		}

		private void btnMain1_MouseLeave(object sender, System.EventArgs e)
		{
			this.btnMain1.BackColor = System.Drawing.Color.Gray;
		}

		private void btnMain1_MouseEnter(object sender, System.EventArgs e)
		{
    		this.btnMain1.BackColor = System.Drawing.Color.LightGray;
		}
		
		private void InitializeMainFormPanelsAndButtons()
		{
			
            //CORE ANALYSIS PANEL
            this.m_pnlOptimizer = new Panel();
			this.grpboxLeft.Controls.Add(this.m_pnlOptimizer);
            this.PanelProperties(this.panel1,ref this.m_pnlOptimizer);
			this.m_pnlOptimizer.Visible=false;
			this.m_pnlOptimizer.Name="CORE";
            //user-defined PRE/POST variables
            this.m_btnOptimizerUserVariables = new btnMainForm(this);
            this.m_pnlOptimizer.Controls.Add(this.m_btnOptimizerUserVariables);
            this.m_btnOptimizerUserVariables.Size = this.btnMain1.Size;
            this.m_btnOptimizerUserVariables.Location = this.btnMain1.Location;
            this.m_btnOptimizerUserVariables.Text = "Define Calculated Variables";
            //Load psites and travel times
            this.m_btnOptimizerLoadGisData = new btnMainForm(this);
            this.m_pnlOptimizer.Controls.Add(this.m_btnOptimizerLoadGisData);
            this.m_btnOptimizerLoadGisData.Size = this.btnMain1.Size;
            this.m_btnOptimizerLoadGisData.Left = this.m_btnOptimizerUserVariables.Left;
            this.m_btnOptimizerLoadGisData.Top = this.m_btnOptimizerUserVariables.Top + this.m_btnOptimizerUserVariables.Height + 5;
            this.m_btnOptimizerLoadGisData.Text = "Load GIS Data";
            //Optimization scenario
            this.m_btnOptimizerScenario = new btnMainForm(this);
            this.m_pnlOptimizer.Controls.Add(this.m_btnOptimizerScenario);
            this.m_btnOptimizerScenario.Size = this.btnMain1.Size;
            this.m_btnOptimizerScenario.Left = this.m_btnOptimizerUserVariables.Left;
            this.m_btnOptimizerScenario.Top = this.m_btnOptimizerLoadGisData.Top + this.m_btnOptimizerLoadGisData.Height + 5;
            this.m_btnOptimizerScenario.Text = "Optimization Scenario";
            //Generate SQLITE database
            //this.m_btnOptimizerSqlite = new btnMainForm(this);
            //this.m_pnlOptimizer.Controls.Add(this.m_btnOptimizerSqlite);
            //this.m_btnOptimizerSqlite.Size = this.btnMain1.Size;
            //this.m_btnOptimizerSqlite.Left = this.m_btnOptimizerUserVariables.Left;
            //this.m_btnOptimizerSqlite.Top = this.m_btnOptimizerScenario.Top + this.m_btnOptimizerScenario.Height + 5;
            //this.m_btnOptimizerSqlite.Text = "Upgrade for v5.11.0";
			//merge scenarios
			this.m_btnCoreMerge = new btnMainForm(this);
			this.m_pnlOptimizer.Controls.Add(this.m_btnCoreMerge);
			this.m_btnCoreMerge.Size = this.btnMain1.Size;
			this.m_btnCoreMerge.Left  = this.m_btnOptimizerScenario.Left;
			this.m_btnCoreMerge.Top = this.m_btnOptimizerScenario.Top + this.m_btnOptimizerScenario.Height + 5;
			this.m_btnCoreMerge.Text = "Join Data From Multiple Scenarios";
            this.m_btnCoreMerge.Hide();

			//PROCESSOR PANEL
			this.m_pnlProcessor = new Panel();
			this.grpboxLeft.Controls.Add(this.m_pnlProcessor);
			this.PanelProperties(this.panel1,ref this.m_pnlProcessor);
			this.m_pnlProcessor.Visible=false;
			this.m_pnlOptimizer.Name="PROCESSOR";

			//tree species
			this.m_btnProcessorTreeSpc = new btnMainForm(this);
			this.m_pnlProcessor.Controls.Add(this.m_btnProcessorTreeSpc);
			this.m_btnProcessorTreeSpc.Size = this.btnMain1.Size;
            this.m_btnProcessorTreeSpc.Location = this.btnMain1.Location;
			this.m_btnProcessorTreeSpc.Left  = this.btnMain1.Left;
			this.m_btnProcessorTreeSpc.Text = "Tree Species";
            this.m_btnProcessorTreeSpc.strToolTip = "Step 1 - Assess Data Readiness For This Item: \n" +
                                                             "1)Check If Each FIA Tree Species Code, FVS Variant, And FVS Species Code Combination Is Present In The Tree Species Table"; 
				               
           
			//start biosum processor button			
			this.m_btnProcessorStart = new btnMainForm(this);
			this.m_pnlProcessor.Controls.Add(this.m_btnProcessorStart);
			this.m_btnProcessorStart.Size = this.btnMain1.Size;
			this.m_btnProcessorStart.Location = this.btnMain1.Location;
            this.m_btnProcessorStart.Left = this.m_btnProcessorTreeSpc.Left;
            this.m_btnProcessorStart.Top = this.m_btnProcessorTreeSpc.Top + this.m_btnProcessorTreeSpc.Height + 5;
			this.m_btnProcessorStart.Text = "Start Biosum Processor";
			this.m_btnProcessorStart.strToolTip = "Step 2 - Execute Processor";


            //DATABASE PANEL AND BUTTONS
            this.m_pnlDb = new Panel();
			this.grpboxLeft.Controls.Add(this.m_pnlDb);
            this.PanelProperties(this.panel1,ref this.m_pnlDb);
            this.m_pnlDb.Visible=false;
			this.m_pnlDb.Name = "DATABASE";
			//plot data button
			this.m_btnDbPlotData = new btnMainForm(this);
			this.m_pnlDb.Controls.Add(this.m_btnDbPlotData);
			this.m_btnDbPlotData.Size = this.btnMain1.Size;
			this.m_btnDbPlotData.Location = this.btnMain1.Location;
			this.m_btnDbPlotData.Text = "Plot Data";

			//processing sites
            // Hide button in v5.11.0. No longer needed
			//this.m_btnDbPSite = new btnMainForm(this);
			//this.m_pnlDb.Controls.Add(this.m_btnDbPSite);
			//this.m_btnDbPSite.Size = this.btnMain1.Size;
			//this.m_btnDbPSite.Left = this.m_btnDbPlotData.Left;
			//this.m_btnDbPSite.Top = this.m_btnDbPlotData.Top + this.m_btnDbPlotData.Height + 5;
			//this.m_btnDbPSite.Text = "Wood Processing Sites";

			//project data sources
			this.m_btnDbDataSource = new btnMainForm(this);
			this.m_pnlDb.Controls.Add(this.m_btnDbDataSource);
			this.m_btnDbDataSource.Size = this.btnMain1.Size;
			this.m_btnDbDataSource.Left = this.m_btnDbPlotData.Left;
			this.m_btnDbDataSource.Top = this.m_btnDbPlotData.Top + this.m_btnDbPlotData.Height + 5;
			this.m_btnDbDataSource.Text = "Project Data Sources";
			this.m_btnDbDataSource.Enabled=true;
			//table management
			this.m_btnDbTableMgmt = new btnMainForm(this);
			this.m_pnlDb.Controls.Add(this.m_btnDbTableMgmt);
			this.m_btnDbTableMgmt.Size = this.btnMain1.Size;
			this.m_btnDbTableMgmt.Left  = this.m_btnDbPlotData.Left;
			this.m_btnDbTableMgmt.Top = this.m_btnDbDataSource.Top + this.m_btnDbDataSource.Height + 5;
			this.m_btnDbTableMgmt.Text = "Manage Tables";
            //table management
            this.m_btnSQLiteTableMgmt = new btnMainForm(this);
            this.m_pnlDb.Controls.Add(this.m_btnSQLiteTableMgmt);
            this.m_btnSQLiteTableMgmt.Size = this.btnMain1.Size;
            this.m_btnSQLiteTableMgmt.Left = this.m_btnDbPlotData.Left;
            this.m_btnSQLiteTableMgmt.Top = this.m_btnDbTableMgmt.Top + this.m_btnDbTableMgmt.Height + 5;
            this.m_btnSQLiteTableMgmt.Text = "Manage SQLite Tables";
			//generate travel times
			this.m_btnDbRandomTravelTimes = new btnMainForm(this);
			this.m_pnlDb.Controls.Add(this.m_btnDbRandomTravelTimes);
			this.m_btnDbRandomTravelTimes.Size = this.btnMain1.Size;
			this.m_btnDbRandomTravelTimes.Left  = this.m_btnDbPlotData.Left;
			this.m_btnDbRandomTravelTimes.Top = this.m_btnSQLiteTableMgmt.Top + this.m_btnSQLiteTableMgmt.Height + 5;
			this.m_btnDbRandomTravelTimes.Text = "Generate Random Travel Times";
			this.m_btnDbRandomTravelTimes.Enabled=true;
			this.m_btnDbRandomTravelTimes.Visible=false;

			//FRCS PANEL
			this.m_pnlFrcs = new Panel();
			this.grpboxLeft.Controls.Add(this.m_pnlFrcs);
			this.PanelProperties(this.panel1,ref this.m_pnlFrcs);
			this.m_pnlFrcs.Visible=false;
			this.m_pnlFrcs.Name = "FRCS";



			//FVS PANEL
			this.m_pnlFvs = new Panel();
			this.grpboxLeft.Controls.Add(this.m_pnlFvs);
			this.PanelProperties(this.panel1,ref this.m_pnlFvs);
			this.m_pnlFvs.Visible=false;
			this.m_pnlFvs.Name = "FVS";



			////plot fvs variant
			//this.m_btnFvsVariant = new btnMainForm(this);
			//this.m_pnlFvs.Controls.Add(this.m_btnFvsVariant);
			//this.m_btnFvsVariant.Size = this.btnMain1.Size;
			//this.m_btnFvsVariant.Location = this.btnMain1.Location;
			//this.m_btnFvsVariant.Text = "Plot FVS Variants";
			//this.m_btnFvsVariant.strToolTip = "Step 1 - Assign An FVS Variant To Each Plot";

			//rx input and edit
			this.m_btnFvsRx = new btnMainForm(this);
			this.m_pnlFvs.Controls.Add(this.m_btnFvsRx);
			this.m_btnFvsRx.Size = this.btnMain1.Size;
            this.m_btnFvsRx.Location = this.btnMain1.Location;
            this.m_btnFvsRx.Text = "Rx";
			this.m_btnFvsRx.strToolTip = "Step 2 - Assign A Label Identifier To Each FVS Treatment";

			//rx package input and edit
			this.m_btnFvsRxPackage = new btnMainForm(this);
			this.m_pnlFvs.Controls.Add(this.m_btnFvsRxPackage);
			this.m_btnFvsRxPackage.Size = this.btnMain1.Size;
			this.m_btnFvsRxPackage.Left  = this.m_btnFvsRx.Left;
			this.m_btnFvsRxPackage.Top = this.m_btnFvsRx.Top + this.m_btnFvsRx.Height + 5;
			this.m_btnFvsRxPackage.Text = "Rx Package";
			this.m_btnFvsRxPackage.strToolTip = "Step 3 - Assign one or more FVS treatments to a package";


			//fvs output button
			this.m_btnFvsTreeSpc = new btnMainForm(this);
			this.m_pnlFvs.Controls.Add(this.m_btnFvsTreeSpc);
			this.m_btnFvsTreeSpc.Size = this.btnMain1.Size;
			this.m_btnFvsTreeSpc.Left = this.m_btnFvsRx.Left;
			this.m_btnFvsTreeSpc.Top = this.m_btnFvsRxPackage.Top + this.m_btnFvsRxPackage.Height + 5;
			this.m_btnFvsTreeSpc.Text = "Tree Species";
			this.m_btnFvsTreeSpc.strToolTip = "Step 4 - Assess Data Readiness: Check If Each Tree Species And FVS Variant Combination Is In The Tree Species Table";


			//fvs input
			this.m_btnFvsInput = new btnMainForm(this);
			this.m_pnlFvs.Controls.Add(this.m_btnFvsInput);
			this.m_btnFvsInput.Size = this.btnMain1.Size;
			this.m_btnFvsInput.Left = this.m_btnFvsRx.Left;
			this.m_btnFvsInput.Top = this.m_btnFvsTreeSpc.Top + this.m_btnFvsTreeSpc.Height + 5;
			this.m_btnFvsInput.Text = "FVS Input";
			this.m_btnFvsInput.strToolTip = "Step 5 - Create FVS Input";

            // No longer needed as of v5.11.0
            //Convert SQLite to MS Access FVSOUT databases
            //this.m_btnCreateFvsOutputMdbs = new btnMainForm(this);
            //this.m_pnlFvs.Controls.Add(this.m_btnCreateFvsOutputMdbs);
            //this.m_btnCreateFvsOutputMdbs.Size = this.btnMain1.Size;
            //this.m_btnCreateFvsOutputMdbs.Left = this.m_btnFvsVariant.Left;
            //this.m_btnCreateFvsOutputMdbs.Top = this.m_btnFvsInput.Top + this.m_btnFvsInput.Height + 5;
            //this.m_btnCreateFvsOutputMdbs.Text = "Create FVSOUT MDBs";
            //this.m_btnCreateFvsOutputMdbs.strToolTip = "Step 5.5 - Populate FVSOUT MDBs with SQLite FVSOUT.DB data";

            //fvs output button
            this.m_btnFvsOutput = new btnMainForm(this);
			this.m_pnlFvs.Controls.Add(this.m_btnFvsOutput);
			this.m_btnFvsOutput.Size = this.btnMain1.Size;
			this.m_btnFvsOutput.Left = this.m_btnFvsRx.Left;
			this.m_btnFvsOutput.Top = this.m_btnFvsInput.Top + this.m_btnFvsInput.Height + 5;
            this.m_btnFvsOutput.Text = "FVS Output Data";
			this.m_btnFvsOutput.strToolTip = "Step 6 - Update FVS PRE/POST And Processor Tree Tables With FVS Output data";


			//CURRENT PANEL
			this.m_pnlCurrent = new Panel();
			this.grpboxLeft.Controls.Add(this.m_pnlCurrent);
			this.PanelProperties(this.panel1,ref this.m_pnlCurrent);
			this.m_pnlCurrent.Visible=false;


			//
			//statusbar and progressbar
			//
			this.sbpInfo = new StatusBarPanel();
			this.sbpProgress = new StatusBarPanel();
			ProgressStatus1 = new ProgressStatus();


			this.sbpInfo.Alignment = System.Windows.Forms.HorizontalAlignment.Right;
			this.sbpInfo.Text="Ready";
			this.sbpInfo.Width=(this.Width / 2);
			this.sbpProgress.AutoSize=System.Windows.Forms.StatusBarPanelAutoSize.Spring;
			ProgressStatus1.Panels.Add(sbpProgress);
			ProgressStatus1.Panels.Add(sbpInfo);
			ProgressStatus1.ShowPanels=true;
			ProgressStatus1.SetProgressBar=0;
			ProgressStatus1.ProgressBar1.Minimum=0;
			ProgressStatus1.ProgressBar1.Maximum=100;
			ProgressStatus1.ProgressBar1.Show();
			ProgressStatus1.ProgressBar1.Value=50;
			this.Controls.Add(ProgressStatus1);
			this.ProgressStatus1.ProgressBar1.Hide();
			frmMain.g_sbpInfo = this.sbpInfo;
			frmMain.g_sbpProgress = this.sbpProgress;
			frmMain.g_oFrmMain=this;

			

		}
		private void PanelProperties(System.Windows.Forms.Panel p_pnlSource, ref System.Windows.Forms.Panel p_pnlTarget)
		{
			p_pnlTarget.AutoScroll = p_pnlSource.AutoScroll;
			p_pnlTarget.BackColor = p_pnlSource.BackColor;
			p_pnlTarget.BorderStyle = p_pnlSource.BorderStyle;
			p_pnlTarget.ForeColor = p_pnlSource.ForeColor;
			p_pnlTarget.Location = p_pnlSource.Location;
			p_pnlTarget.Size = p_pnlSource.Size;
		}

		private void mnuHelpAbout_Click(object sender, System.EventArgs e)
		{
			frmAbout p_frmAbout = new frmAbout();
			p_frmAbout.MaximizeBox = false;
			p_frmAbout.MinimizeBox = false;
			p_frmAbout.Show();


			int intHt=p_frmAbout.Height;
			int intWd=p_frmAbout.Width;


			int intHt2=p_frmAbout.btnOK.Height;
			int intTop=p_frmAbout.btnOK.Top;
			while (intTop + intHt2 + 20
					    >=  intHt)
			{
					intHt += 10;
			
			}
			p_frmAbout.Height = intHt;

			//see if the width is less than the longest group box
			
			if (p_frmAbout.Width <= p_frmAbout.grpboxDesc.Width)
			{
				//resize the width of the form so that the whole group box
				//is displayed
				int intWd2=p_frmAbout.grpboxDesc.Width;
				int intLeft=p_frmAbout.grpboxDesc.Left;
				while (intWd2 + intLeft + 20 >= intWd)
				{
					intWd += 10;
			
				}
				

			}

            p_frmAbout.Height = intHt;
			p_frmAbout.Width = intWd;

			
			p_frmAbout.resize_frmAbout();
			p_frmAbout.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			

		}

        private void mnuReleaseNotes_Click(object sender, System.EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(Help.DefaultReleaseNotesFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "RELEASENOTES", "PAGE1" });
        }

		private void mnuHelpTechnicalSupport_Click(object sender, System.EventArgs e)
		{
			frmTechSupport p_frmTechSupport = new frmTechSupport();
			p_frmTechSupport.MaximizeBox = false;
			p_frmTechSupport.MinimizeBox = false;
			p_frmTechSupport.Show();
			int intHt=p_frmTechSupport.Height;
			int intWd=p_frmTechSupport.Width;


			int intHt2=p_frmTechSupport.btnOK.Height;
			int intTop=p_frmTechSupport.btnOK.Top;
			while (intTop + intHt2 + 20
				>=  intHt)
			{
				intHt += 10;
			
			}
			p_frmTechSupport.Height = intHt;

			//see if the width is less than the longest group box
			
			if (p_frmTechSupport.Width <= 
				  p_frmTechSupport.btnOK.Left + p_frmTechSupport.btnOK.Width)
			{
				//resize the width of the form so that the whole button
				//is displayed
				int intWd2=p_frmTechSupport.btnOK.Width;
				int intLeft=p_frmTechSupport.btnOK.Left;
				while (intWd2 + intLeft + 20 >= intWd)
				{
					intWd += 10;
			
				}
				

			}

			p_frmTechSupport.Height = intHt;
			p_frmTechSupport.Width = intWd;
		    p_frmTechSupport.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		
		}

		private void mnuViewNotes_Click(object sender, System.EventArgs e)
		{

			int intAvailHt=0;;
			int intAvailWd=0;
			
			this.frmProject.uc_project1.Visible=false;
			this.frmProject.uc_scenario1.Visible=false;
			this.frmProject.uc_project_document_links1.Visible=false;
			this.frmProject.uc_contact_list1.Visible=false;
            this.frmProject.uc_project_notes1.loadnotes();				
			this.frmProject.uc_project_notes1.Visible=true;
			this.frmProject.Visible = true;
			if (this.frmProject.WindowState==System.Windows.Forms.FormWindowState.Minimized)
			{
				this.frmProject.WindowState = System.Windows.Forms.FormWindowState.Normal;
			}
			else
			{
				
			}

			this.frmProject.Focus();
			intAvailWd = this.ClientSize.Width - this.grpboxLeft.Left - this.grpboxLeft.Width - 20;
			intAvailHt = this.ClientSize.Height - this.tlbMain.Top - this.tlbMain.Height - 20;
			this.frmProject.Height = intAvailHt; 
			this.frmProject.Width = intAvailWd; 
			this.frmProject.Left = 0;
			this.frmProject.Top = 0;

		
		}
		public void LoadProcessorTreeSpcForm(Control p_oParentControl,string p_strPanel)
		{
			string strTitle;
			if (p_strPanel.Trim() == "FVS")
			{
				strTitle = "FVS: Tree Species";
			}
			else
				strTitle = "Processor: Tree Species";

			//check to see if the form has already been loaded
			if (this.IsChildWindowVisible(strTitle) == false) 
			{
                this.ActivateStandByAnimation(this.WindowState, this.Left, this.Height, this.Width, this.Top);
				frmMain.g_sbpInfo.Text = "Loading Tree Species...Stand By";
				this.m_frmProcessorSpc = new frmDialog(this);
				this.m_frmProcessorSpc.MaximizeBox = true;
				this.m_frmProcessorSpc.BackColor = System.Drawing.SystemColors.Control;
				this.m_frmProcessorSpc.Text = strTitle;
				FIA_Biosum_Manager.uc_processor_tree_spc p_uc = new uc_processor_tree_spc(this.frmProject.uc_project1.txtRootDirectory.Text.Trim());
				if (p_uc.m_intError < 0) 
				{
					this.m_frmProcessorSpc.Dispose();
					return;
				}
				this.m_frmProcessorSpc.Controls.Add(p_uc);
				this.m_frmProcessorSpc.ProcessorTreeSpcUserControl = p_uc;
				this.m_frmProcessorSpc.Height=0;
				this.m_frmProcessorSpc.Width=0;
				if (p_uc.Top + p_uc.Height > this.m_frmProcessorSpc.ClientSize.Height + 2)
				{
					for (int x=1;;x++)
					{
						this.m_frmProcessorSpc.Height = x;
						if (p_uc.Top + 
							p_uc.Height < 
							this.m_frmProcessorSpc.ClientSize.Height)
						{
							break;
						}
					}

				}
				if (p_uc.Left + p_uc.Width > this.m_frmProcessorSpc.ClientSize.Width + 2)
				{
					for (int x=1;;x++)
					{
						this.m_frmProcessorSpc.Width = x;
						if (p_uc.Left + 
							p_uc.Width < 
							this.m_frmProcessorSpc.ClientSize.Width)
						{
							break;
						}
					}

				}
				p_uc.Dock = System.Windows.Forms.DockStyle.Fill;
				//p_uc.loadvalues();
							

				p_uc.SetMenuOptions(p_strPanel);

				this.m_frmProcessorSpc.Left = 0;
				this.m_frmProcessorSpc.Top = 0;
				this.m_frmProcessorSpc.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
                this.m_frmProcessorSpc.MinimizeMainForm = true;
				this.m_frmProcessorSpc.DisposeOfFormWhenClosing = true;
				this.m_frmProcessorSpc.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

				frmMain.g_sbpInfo.Text = "Ready";

                p_oParentControl.Enabled = false;
                m_frmProcessorSpc.ParentControl = p_oParentControl;
                DeactivateStandByAnimation();
				this.m_frmProcessorSpc.Show(p_oParentControl);

			}
			else
			{
				if (this.m_frmProcessorSpc.WindowState == System.Windows.Forms.FormWindowState.Minimized)
					this.m_frmProcessorSpc.WindowState = System.Windows.Forms.FormWindowState.Normal;

				this.m_frmProcessorSpc.Focus();
					
			}

		}

		private void mnuViewContacts_Click(object sender, System.EventArgs e)
		{
			
			this.frmProject.uc_project_notes1.Visible=false;
			this.frmProject.uc_project1.Visible=false;
			this.frmProject.uc_scenario1.Visible=false;
			this.frmProject.uc_project_document_links1.Visible=false;
			this.frmProject.uc_contact_list1.loadvalues(this.frmProject.uc_project1.txtRootDirectory.Text.Trim());
			this.frmProject.uc_contact_list1.Visible=true;
			this.frmProject.Visible = true;
			if (this.frmProject.WindowState==System.Windows.Forms.FormWindowState.Minimized)
			{
				this.frmProject.WindowState = System.Windows.Forms.FormWindowState.Normal;
			}
			else
			{
				
			}

			this.frmProject.Focus();
			this.frmProject.uc_contact_list1.m_oResizeForm.ControlToResize = frmProject;
			
			this.frmProject.uc_contact_list1.m_oResizeForm.ResizeControl();

			
		}

		private void grpboxLeft_Resize(object sender, System.EventArgs e)
		{
            try
            {
                this.m_pnlCurrent.Height = this.grpboxLeft.Height - this.m_pnlCurrent.Top - (this.btnOptimizer.Height * 3) - 10;
            }
            catch 
            {
            }
		}

		
	
		private void ProcessButton_EnabledChanged(object sender, System.EventArgs e)
		{
			
			System.Windows.Forms.Button btn = ( Button ) sender;
			if (btn.Dock==System.Windows.Forms.DockStyle.Top)
			{
				btn.ForeColor = Color.Red;
			}
			else
			{
				if (btn.Enabled)
					btn.ForeColor = Color.Black;
				else
					btn.ForeColor = System.Drawing.SystemColors.GrayText;
			}
		}
	
		
		private void ProcessButton_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
			System.Windows.Forms.Button btn = ( Button ) sender;
			System.Drawing.Brush drawBrush = new SolidBrush( btn.ForeColor );
			System.Drawing.StringFormat sf = new StringFormat();
			sf.Alignment = System.Drawing.StringAlignment.Center;
			sf.LineAlignment = System.Drawing.StringAlignment.Center;
			string s = btn.Text;
			float floatX = (int)(btn.Width * .5);
			float floatY =  (int)(btn.Height * .5);
			e.Graphics.DrawString( s, btn.Font, drawBrush,floatX,floatY, sf );
			drawBrush.Dispose();
			sf.Dispose();
			

		}

		private void mnuSettings_Click(object sender, System.EventArgs e)
		{
			FIA_Biosum_Manager.frmSettings frmTemp = new frmSettings();
			frmTemp.ShowDialog();
		}

	
		public class ProgressStatus : System.Windows.Forms.StatusBar
		{
			public System.Windows.Forms.ProgressBar ProgressBar1;
			private int _intProgressBar = -1;

			public ProgressStatus()
			{
				ProgressBar1 = new ProgressBar();
				ProgressBar1.Hide();
				this.Controls.Add(ProgressBar1);
				this.DrawItem += new StatusBarDrawItemEventHandler(ProgressStatus1_DrawItem);
			}
			public int SetProgressBar
			{
				get {return _intProgressBar;}
				set 
				{
					_intProgressBar = value;
					this.Panels[_intProgressBar].Style = System.Windows.Forms.StatusBarPanelStyle.OwnerDraw;
				}

			}
			public void ProgressStatus1_DrawItem(object sender, System.Windows.Forms.StatusBarDrawItemEventArgs e)
			{
				ProgressBar1.Location = new Point(e.Bounds.X,e.Bounds.Y);
				ProgressBar1.Size = new Size(e.Bounds.Width,e.Bounds.Height);
			}


		}

        private void mnuToolsFCS_Click(object sender, EventArgs e)
        {
            this.ActivateStandByAnimation(
                this.WindowState,
                this.Left,
                this.Height,
                this.Width,
                this.Top);
            frmFCSTreeVolumeEdit oForm = new frmFCSTreeVolumeEdit();
            this.DeactivateStandByAnimation();
            if (oForm.DialogResult != System.Windows.Forms.DialogResult.Abort)
            {
                oForm.MdiParent = this;

                oForm.Show();
            }
        }
        public void ActivateStandByAnimation()
        {

            lock (_locker)
            {
                if (standByAnimation == null)
                {
                    standByAnimation = new StandByAnimation.StandByAnimation();
                    standByAnimationThread = new System.Threading.Thread(standByAnimation.StartSplashing);
                    standByAnimationThread.IsBackground = true;
                    standByAnimationThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    standByAnimationThread.Start();
                }
            }
        }
        public void ActivateStandByAnimation(System.Windows.Forms.FormWindowState p_oWindowState, double p_dblLeft, double p_dblHeight, double p_dblWidth, double p_dblTop)
        {

            lock (_locker)
            {
                if (standByAnimation == null)
                {
                    standByAnimation = new StandByAnimation.StandByAnimation(p_oWindowState, p_dblLeft, p_dblHeight, p_dblWidth, p_dblTop);
                    standByAnimationThread = new System.Threading.Thread(standByAnimation.StartSplashing);
                    standByAnimationThread.IsBackground = true;
                    standByAnimationThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    standByAnimationThread.Start();
                }
            }
        }
        public void DeactivateStandByAnimation()
        {
           
                if (standByAnimation != null)
                {
                    standByAnimation.StopSplashing();
                }
                standByAnimationThread.Join();
                standByAnimation = null;
            
        }

        private void mnuToolsProjectRootFolder_Click(object sender, EventArgs e)
        {
            this.ActivateStandByAnimation(
               this.WindowState,
               this.Left,
               this.Height,
               this.Width,
               this.Top);
            frmScanAndSynchronizeProjectRootFolderTool oForm = new frmScanAndSynchronizeProjectRootFolderTool();
            oForm.MdiParent = this;
            this.DeactivateStandByAnimation();
            oForm.Show();
        }

        private void resizeProjectForm(frmMain parentForm)
        {
            if (parentForm.frmProject.Height < parentForm.frmProject.uc_project1.m_intFullHt)
                parentForm.frmProject.Height = parentForm.frmProject.uc_project1.m_intFullHt;
            if (parentForm.frmProject.Width < parentForm.frmProject.uc_project1.m_intFullWh)
                parentForm.frmProject.Width = parentForm.frmProject.uc_project1.m_intFullWh;

        }

        /// <summary>
        /// Manages the biosum_ref.accdb that lives in the current user's AppData directory
        /// </summary>
        private void CheckForBiosumRefData()
        {
            string strDestFile = frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceDbFile;
            try
            {
                bool bCopyDatabase = true;
                string strSourceFile = frmMain.g_oEnv.strAppDir + "\\db\\" + Tables.Reference.DefaultBiosumReferenceDbFile;
                if (System.IO.File.Exists(strDestFile) == true)
                {
                    // Check to see if the version is correct
                    ado_data_access oAdo = new ado_data_access();
                    string strConn = oAdo.getMDBConnString(strDestFile, "", "");
                    using (var oConn = new System.Data.OleDb.OleDbConnection(strConn))
                    {
                        oConn.Open();
                        // If the table doesn't exist, the database needs to be refreshed
                        if (oAdo.TableExist(oConn, Tables.Reference.DefaultBiosumReferenceVersionTableName))
                        {
                            string strSql = "select version_num from " + Tables.Reference.DefaultBiosumReferenceVersionTableName;
                            double dblVersion = oAdo.getSingleDoubleValueFromSQLQuery(oConn, strSql,
                                Tables.Reference.DefaultBiosumReferenceVersionTableName);
                            if (dblVersion >= g_intRefDbVer)
                            {
                                // if the version number is greater than or equal do the current version, the database doesn't need to be refreshed
                                bCopyDatabase = false;
                            }
                        }
                    }
                    oAdo = null;
                } 
               
                if (bCopyDatabase == true)
                    // Copy it the database from the app install directory
                    System.IO.File.Copy(strSourceFile, strDestFile, true);
            }
            catch (Exception)
            {
                MessageBox.Show("!! An error occurred while accessing biosum_ref.accdb at " + 
                    System.IO.Path.GetDirectoryName(strDestFile) + " !!", "FIA Biosum");
            }
        }

        private void CheckForBiosumRefSqliteData()
        {
            string strDestFile = frmMain.g_oEnv.strApplicationDataDirectory.Trim() +
                frmMain.g_strBiosumDataDir + "\\" + Tables.Reference.DefaultBiosumReferenceSqliteFile;
            try
            {
                bool bCopyDatabase = true;
                string strSourceFile = frmMain.g_oEnv.strAppDir + "\\db\\" + Tables.Reference.DefaultBiosumReferenceSqliteFile;
                if (System.IO.File.Exists(strDestFile) == true)
                {
                    // Check to see if the version is correct
                    SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
                    string strConn = oDataMgr.GetConnectionString(strDestFile);
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                    {
                        conn.Open();
                        if (!oDataMgr.TableExist(conn, Tables.Reference.DefaultBiosumReferenceVersionTableName))
                        {
                            oDataMgr.m_strSQL = "SELECT version_num FROM " + Tables.Reference.DefaultBiosumReferenceVersionTableName;
                            double dblVersion = oDataMgr.getSingleDoubleValueFromSQLQuery(conn, oDataMgr.m_strSQL, 
                                Tables.Reference.DefaultBiosumReferenceVersionTableName);
                            if (dblVersion >= g_intRefSqliteDbVer)
                            {
                                bCopyDatabase = false;
                            }
                        }
                    }
                    oDataMgr = null;
                }
                if (bCopyDatabase == true)
                    // Copy it the database from the app install directory
                    System.IO.File.Copy(strSourceFile, strDestFile, true);
            }
            catch (Exception)
            {
                MessageBox.Show("!! An error occurred while accessing biosum_ref.db at " +
                    System.IO.Path.GetDirectoryName(strDestFile) + " !!", "FIA Biosum");
            }
        }

        /// <summary>
        /// Manages the required files for FCS 0.3. These files are packaged in the application fcs directory
        /// and copied to the AppData\FIABioSum folder
        /// </summary>
        private void CheckForFcsFiles()
        {
            string[] arrRequiredFiles = {Tables.VolumeAndBiomass.BioSumCompsJar, Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase,
                                         Tables.VolumeAndBiomass.DefaultSqliteConfigDatabase, Tables.VolumeAndBiomass.FcsTreeCalcBat};
            string strTargetPath = "";
            try
            {
                for (int i = 0; i < arrRequiredFiles.Length; i++)
                {
                    strTargetPath = frmMain.g_oEnv.strApplicationDataDirectory.Trim() + frmMain.g_strBiosumDataDir + @"\" + arrRequiredFiles[i];
                    string strSourcePath = frmMain.g_oEnv.strAppDir + @"\fcs\" + arrRequiredFiles[i];
                    if (!System.IO.File.Exists(strTargetPath) == true)
                    {
                        System.IO.File.Copy(strSourcePath, strTargetPath);
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("!! An error occurred while copying fcs file to " +
                    System.IO.Path.GetDirectoryName(strTargetPath) + " !!", "FIA Biosum");
            }
        }

        private void mnuToolsPath_Click(object sender, EventArgs e)
        {
            MessageBox.Show(System.Environment.GetEnvironmentVariable("PATH").ToString());
        }
	}
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Xps.Packaging;
using System.Windows;
using System.Data;
using System.Reflection;
using System.Windows.Controls;
using MS.Internal.PresentationUI;
using System.IO.Packaging;
using MS.Internal.Documents.Application;
using System.Threading;
using System.Windows.Threading;


namespace FIA_Biosum_Manager
{
    public class Help
    {
        private System.Data.DataSet m_dsHelp;
        private System.Data.DataTable m_dtHelp;
        private Thread m_oHelpThread=null;
        public XPSDocumentViewer xpsDocumentViewer = null;
        System.Windows.Xps.Packaging.XpsDocument m_oXpsDocument = null;

        public int m_intCurrentPageNumber=-1;
        FrameworkElement fe;

        public const int COL_DEFAULT_HELP_PARENT = 0;
        public const int COL_DEFAULT_HELP_CHILD = 1;
        private string _strXPSFile = null;
        private env _env = null;

        static public string DefaultDatabaseXPSFile { get { return "DATABASE_Help.xps"; } }
        static public string DefaultProcessorXPSFile { get { return "PROCESSOR_Help.xps"; } }
        static public string DefaultFvsXPSFile { get { return "FVS_Help.xps"; } }
        static public string DefaultTreatmentOptimizerFile { get { return "TREATMENT_OPTIMIZER_Help.xps"; } }
        static public string DefaultMainFile { get { return "MAIN_Help.xps"; } }
        static public string DefaultReleaseNotesFile { get { return "RELEASE_NOTES.xps"; } }

        
        //call default constructor after initializing variables
        public Help(string p_strXPSFile, env p_env)
        {
            _strXPSFile = p_strXPSFile;
            _env = p_env;
            _strXSD = p_env.strAppDir + "\\Help\\Help.xsd";
            _strXML = p_env.strAppDir + "\\Help\\Help.xml";

            //default values
            _strPrimaryKeyColumns = new string[2];
            _strPrimaryKeyColumns[COL_DEFAULT_HELP_PARENT] = "PARENT_ITEM";
            _strPrimaryKeyColumns[COL_DEFAULT_HELP_CHILD] = "CHILD_ITEM";
            _strPrimaryKeyValues = new string[2];
            _strPrimaryKeyValues[COL_DEFAULT_HELP_PARENT] = "MAINWINDOW";
            _strPrimaryKeyValues[COL_DEFAULT_HELP_CHILD] = "MAINWINDOW";
            LoadXml();
 
        }

        public void LoadXml(string p_strXSDFile, string p_strXMLFile,string p_strTableName,string[] p_strPrimaryKeyColumns)
        {
            _strXSD = p_strXSDFile;
            _strXML = p_strXMLFile;
            _strTable = p_strTableName;
            _strPrimaryKeyColumns = p_strPrimaryKeyColumns;
            LoadXml();
           
        }
        public void LoadXml()
        {
            this.m_dsHelp = new System.Data.DataSet();
            this.m_dsHelp.ReadXmlSchema(XSDFile);
            this.m_dsHelp.ReadXml(XMLFile, System.Data.XmlReadMode.InferSchema);
            m_dtHelp = m_dsHelp.Tables[TableName];

            //define the primary key for the table
            if (PrimaryKeyColumns != null)
            {
                DataColumn[] colPk = new DataColumn[PrimaryKeyColumns.Length];
                for (int x = 0; x <= PrimaryKeyColumns.Length - 1; x++)
                {
                    colPk[x] = m_dtHelp.Columns[PrimaryKeyColumns[x]];
                }
                m_dtHelp.PrimaryKey = colPk;
            }

        }
       /// <summary>
       /// Look up the item page number from the primary key values and
       /// create a new instance of the help document viewer
       /// </summary>
       /// <param name="p_strPrimaryKeyValues"></param>
        public void ShowHelp(string[] p_strPrimaryKeyValues)
        {
            //shutdown the current document viewer and thread
            if (xpsDocumentViewer != null)
            {
                ShutdownThread();
            }

            //load the PageNumber variable from the key values
            GetItemPageNumber(p_strPrimaryKeyValues);

            //start a new thread containing the document viewer
            m_oHelpThread = new Thread(new ThreadStart(this.ShowHelp));
            m_oHelpThread.IsBackground = true;
            m_oHelpThread.SetApartmentState(System.Threading.ApartmentState.STA);
            m_oHelpThread.Start();

            
            
        }
        /// <summary>
        /// Close the document viewer and thread
        /// </summary>
        private void ShutdownThread()
        {
            if (xpsDocumentViewer != null)
            {

                xpsDocumentViewer.Dispatcher.Invoke(DispatcherPriority.Normal, (Action)delegate()
                {
                    xpsDocumentViewer.Close();
                });

                xpsDocumentViewer = null;
            }
            if (m_oHelpThread.IsAlive) m_oHelpThread.Abort();
            
        }
        /// <summary>
        /// Create a new instance of the document viewer, open the xps document, and navigate to the 
        /// designated page.
        /// </summary>
        private void ShowHelp()
        {

            
            xpsDocumentViewer = new XPSDocumentViewer();
           
            System.Windows.Xps.Packaging.XpsDocument xpsDoc = new System.Windows.Xps.Packaging.XpsDocument(_env.strAppDir + "\\Help\\" + _strXPSFile, System.IO.FileAccess.Read);
            xpsDocumentViewer.xpsViewer1.Document = xpsDoc.GetFixedDocumentSequence();
            xpsDocumentViewer.ReferenceHelp = this;

            
            
            m_intCurrentPageNumber = PageNumber;
            System.Windows.Documents.DocumentPage oPage = xpsDocumentViewer.xpsViewer1.Document.DocumentPaginator.GetPage(PageNumber);
           
            xpsDocumentViewer.xpsViewer1.GoToPage(PageNumber);
            
            xpsDocumentViewer.WindowState = WindowState.Normal;
            xpsDocumentViewer.Top = frmMain.g_oFrmMain.Top + 200;
            xpsDocumentViewer.Height = frmMain.g_oFrmMain.ClientSize.Height - 200;
            xpsDocumentViewer.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            xpsDocumentViewer.IsEnabled = true;
            xpsDocumentViewer.Visibility = Visibility.Visible;
            
            m_oXpsDocument = xpsDoc;
            
           

            xpsDocumentViewer.ShowDialog();

            
           
        }
        /// <summary>
        /// navigate to the designated page
        /// </summary>
        /// <param name="p_intPageNumber"></param>
        public void GoToPage(int p_intPageNumber)
        {
            PageNumber = p_intPageNumber;
            GoToPage();
           
        }
        /// <summary>
        /// Navigate to the current page number. If the document is already open then
        /// just navigate to the page. If the document viewer is not currently
        /// instantiated then create a new document viewer in its own thread
        /// </summary>
        public void GoToPage()
        {
            m_intCurrentPageNumber = PageNumber;
            if (xpsDocumentViewer != null)
            {
                xpsDocumentViewer.Dispatcher.Invoke(DispatcherPriority.Normal, (Action)delegate()
                {
                    xpsDocumentViewer.xpsViewer1.GoToPage(PageNumber);
                    if (xpsDocumentViewer.WindowState == WindowState.Minimized)
                        xpsDocumentViewer.WindowState = WindowState.Normal;
                });

            }
            else
            {
                m_oHelpThread = new Thread(new ThreadStart(this.ShowHelp));
                m_oHelpThread.IsBackground = true;
                m_oHelpThread.SetApartmentState(System.Threading.ApartmentState.STA);
                m_oHelpThread.Start();
            }

           
            
        }
        public void GoToPage(string[] p_strPrimaryKeyValues)
        {
            GetItemPageNumber(p_strPrimaryKeyValues);
            if (PageNumber > 0)
            {
                GoToPage();
            }

        }
        /// <summary>
        /// Load the PageNumber property with the page number 
        /// assigned to the passed parameter
        /// </summary>
        /// <param name="p_strPrimaryKeyValues"></param>
        private void GetItemPageNumber(string[] p_strPrimaryKeyValues)
        {
            PageNumber = -1;
            System.Object[] oSearch = new Object[PrimaryKeyColumns.Length];
            for (int x = 0; x <= PrimaryKeyColumns.Length - 1; x++)
            {
                oSearch[x] = p_strPrimaryKeyValues[x];
            }
            //search for the record
            System.Data.DataRow oRow = m_dtHelp.Rows.Find(oSearch);
            if (oRow != null)
            {
                PageNumber = Convert.ToInt32(oRow["PAGENUMBER"]);
            }

        }

       
            
        private int _intPageNumber=0;
        public int PageNumber
        {
            get { return _intPageNumber; }
            set { _intPageNumber = value; }
        }
        private string _strXSD = null;
        public string XSDFile
        {
            set {_strXSD=value;}
            get {return _strXSD;}
        }
        private string _strXML = null;
        public string XMLFile
        {
            set { _strXML = value; }
            get { return _strXML; }
        }
        private string _strTable = "BIOSUM_HELP";
        public string TableName
        {
            set { _strTable = value; }
            get { return _strTable; }
        }
        private string[] _strPrimaryKeyColumns = null;
        public string[] PrimaryKeyColumns
        {
            get {return _strPrimaryKeyColumns;}
            set {_strPrimaryKeyColumns=value;}
        }
        private string[] _strPrimaryKeyValues = null;
        public string[] PrimaryKeyValues
        {
            get { return _strPrimaryKeyValues; }
            set { _strPrimaryKeyValues = value; }
        }
        public string XPSFile
        {
            get { return _strXPSFile; }
            set { _strXPSFile = value; }
        }

        private string _strApplicationDirectory;
        public string ApplicationDirectory
        {
            get { return _strApplicationDirectory; }
            set { _strApplicationDirectory = value; }
        }
    }
}

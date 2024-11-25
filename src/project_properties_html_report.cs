using System;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for project_properties_html_report.
	/// </summary>
	public class project_properties_html_report
	{
		public int m_intError=0;
		public string m_strError="";
		public string m_strFile="";
		public string m_strJavaScriptFile="";
		private int m_intLineCount=0;
		private string m_strHtmlPageHeader="";
		private string m_strHtmlCurrentTable="";
		private bool m_bNewPage=true;
		private string m_strLine="";
	    int LINESPERPAGE=50;

		private string _strWindowTitle="";
		private string _strProjectName="";
		private string _strReportHeader="";
		private string _strTreatmentHeader="";
		private string _strPackageHeader="";
		private RxItem_Collection _RxItem_Collection;
		private RxPackageItem_Collection _RxPackage_Collection;
        private bool _bTreatments=true;
		private bool _bPackages=true;


		public project_properties_html_report()
		{
			//
			// TODO: Add constructor logic here
			//
			m_strFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"txt");
			m_strJavaScriptFile = m_strFile.Substring(0,m_strFile.IndexOf(".",0)) + "_java.txt";


			
			

			//frmMain.g_oUtils.WriteText(m_strFile,"<HTML> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptFormat("<SCRIPT LANGUAGE='JavaScript'> \n"));
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptFormat("var linefeed = '\\n'; \r\n"));
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,"function printer_friendly() \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,"{ \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptFormat("window.printWindow = " + 
																	"window.open(''," +
																	"'Debug'," + 
																	"'left=0,top=0,width=500,height=300,scrollbars=yes," + 
																	"status=yes,resizable=yes,menubar=yes'); \n"));
			//window.calendarWindow.opener = self;
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,"window.printWindow.document.open(); \n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,"window.printWindow.document.write(); \n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,"output = window.printWindow.document; \n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<HTML>"));

			 
			
			



			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<!------>"));
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<style type=text/css>"));
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(".pagebreak {page-break-after: always;}"));
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</style>"));


		}
		public void CreateReport()
		{
			
			CoverPage();
			if (ProcessPackages)
			{
				this.Packages();
			}
			if (ProcessTreatments)
			{
				this.Treatments();
			}
			EndReport();
			
		}
		private void CoverPage()
		{

			frmMain.g_oUtils.WriteText(m_strFile,"<body  bgcolor='#ffffff' link=#0000FF vlink=#0000FF alink=#0000FF> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<body  bgcolor=#ffffff link=#0000ff vlink=#0000ff alink=#0000ff>"));
			
			frmMain.g_oUtils.WriteText(m_strFile,"<!--COVER PAGE-->\r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<!--COVER PAGE-->"));

			frmMain.g_oUtils.WriteText(m_strFile,"<head>\r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<head>"));

			frmMain.g_oUtils.WriteText(m_strFile,"<HTTP-EQUIV='PRAGMA' CONTENT='NO-CACHE'>\r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<HTTP-EQUIV=PRAGMA CONTENT=NO-CACHE>"));

			frmMain.g_oUtils.WriteText(m_strFile,"<title>\r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<title>"));

			frmMain.g_oUtils.WriteText(m_strFile,WindowTitle + " \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat(WindowTitle));



			frmMain.g_oUtils.WriteText(m_strFile,"</title> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("</title>"));
			

			frmMain.g_oUtils.WriteText(m_strFile,"<!--report header--> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<!--report header-->"));

			frmMain.g_oUtils.WriteText(m_strFile,"<FONT SIZE='+3'><B> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<FONT SIZE=+3><B>"));

			frmMain.g_oUtils.WriteText(m_strFile,"<CENTER>" + ReportHeader + "</CENTER> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<CENTER>" + ReportHeader + "</CENTER>"));
			m_intLineCount=m_intLineCount + 5;

			frmMain.g_oUtils.WriteText(m_strFile,"</B></FONT><br><br> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("</B></FONT><br><br>"));
			m_intLineCount=m_intLineCount + 2;

			frmMain.g_oUtils.WriteText(m_strFile,"<FONT SIZE='-1' COLOR='white'> \r\n");
			frmMain.g_oUtils.WriteText(m_strFile,"<A  HREF='#PRINT' onclick='printer_friendly()'><b>Printer Friendly</b></A>");
			frmMain.g_oUtils.WriteText(m_strFile,"</FONT><BR><BR>");
			m_intLineCount=m_intLineCount + 3;

			frmMain.g_oUtils.WriteText(m_strFile,"<!--project name--> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<!--project name-->"));

			frmMain.g_oUtils.WriteText(m_strFile,"<FONT SIZE='+1'><B> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<FONT SIZE=+1><B>"));

			frmMain.g_oUtils.WriteText(m_strFile,"Project:&nbsp&nbsp" + ProjectName + " \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("Project:&nbsp&nbsp" + ProjectName));
			m_intLineCount=m_intLineCount + 1;

			frmMain.g_oUtils.WriteText(m_strFile,"</B></FONT> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("</B></FONT>"));

			frmMain.g_oUtils.WriteText(m_strFile,"</head> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("</head>"));

			frmMain.g_oUtils.WriteText(m_strFile,"<br><br> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<br><br>"));
			m_intLineCount=m_intLineCount + 2;

			
		}
		private void EndReport()
		{
			frmMain.g_oUtils.WriteText(m_strFile,"</body> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("</body>"));

			frmMain.g_oUtils.WriteText(m_strFile,"<HEAD> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<HEAD>"));

			frmMain.g_oUtils.WriteText(m_strFile,"<META HTTP-EQUIV='PRAGMA' CONTENT='NO-CACHE'> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<META HTTP-EQUIV=PRAGMA CONTENT=NO-CACHE>"));

			frmMain.g_oUtils.WriteText(m_strFile,"</HEAD> \r\n");	
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("</HEAD>"));


			frmMain.g_oUtils.WriteText(m_strFile,"</HTML> \r\n");
			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("</HTML>"));
			frmMain.g_oUtils.WriteText(this.m_strJavaScriptFile,"window.printWindow.document.close();\r\n");

			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,"} \r\n ");

			frmMain.g_oUtils.WriteText(m_strJavaScriptFile,"</SCRIPT> \r\n");

			System.IO.StreamReader s=null;
			string strWholeFile="";
			string[] rows=null;
			string strHtmlFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"html");
			frmMain.g_oUtils.WriteText(strHtmlFile, "<HTML> \r\n");

			if (System.IO.File.Exists(this.m_strJavaScriptFile))
			{
				
				//Open the file in a stream reader.
				s = new System.IO.StreamReader(this.m_strJavaScriptFile);
				//Read the rest of the data in the file.        
				strWholeFile = s.ReadToEnd();
				rows = strWholeFile.Split("\r\n".ToCharArray());
				foreach(string r in rows)
				{
								    
					frmMain.g_oUtils.WriteText(strHtmlFile,r.Trim()+ " \n");
						
					
				}
				s.Close();
				s=null;
			}
			strWholeFile="";
			rows=null;
			if (System.IO.File.Exists(this.m_strFile))
			{
				
				//Open the file in a stream reader.
				s = new System.IO.StreamReader(this.m_strFile);
				//Read the rest of the data in the file.        
				strWholeFile = s.ReadToEnd();
				rows = strWholeFile.Split("\r\n".ToCharArray());
				foreach(string r in rows)
				{
								    
					frmMain.g_oUtils.WriteText(strHtmlFile,r.Trim() + " \n");
						
					
				}
				s.Close();
				s=null;
			}


			System.Diagnostics.Process.Start(strHtmlFile);
		}
		private void Treatments()
		{
			int x,y;
			this.m_strHtmlCurrentTable="";this.m_strHtmlPageHeader="";
			if (this.TreatmentHeader.Trim().Length > 0)
			{
				this.NewPage();
				frmMain.g_oUtils.WriteText(m_strFile,"<!--treatment header--> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<!--treatment header-->"));

				frmMain.g_oUtils.WriteText(m_strFile,"<FONT SIZE='+1'><B> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<FONT SIZE=+1><B>"));

				frmMain.g_oUtils.WriteText(m_strFile,"<CENTER>" + TreatmentHeader + "</CENTER><BR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<CENTER>" + TreatmentHeader + "</CENTER><BR>"));
				m_intLineCount=m_intLineCount + 3;
			}
			if (this.RxCollection.Count==0)
			{
				frmMain.g_oUtils.WriteText(m_strFile,"<CENTER> No Treatments Defined</CENTER><BR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<CENTER>No Treatments Defined</CENTER><BR>"));
				m_intLineCount=m_intLineCount + 1;
			}
			for (x=0;x<=this.RxCollection.Count-1;x++)
			{
				this.m_strHtmlPageHeader="";
				//
				//TREATMENT ID
				//
				if (x !=0)
				{
					m_intLineCount = LINESPERPAGE;   //new page for every treatment
					this.AddToLineCount(1);
				}
				
				this.m_strHtmlCurrentTable="<TABLE  colspan=2 width=100% align=center border=5 height=50 cellpadding=0 cellspacing=0>";
				this.m_strHtmlPageHeader="<TR>" + 
					                        "<TD align=center colspan=1  bgcolor=#cccccc width='30%' height='30%' valign=center>Item" + 
					                        "</TD>" + 
					                        "<TD align=center colspan=1  bgcolor=#cccccc width='70%' height='30%' valign=center>Description" + 
											"</TD>" + 
					                     "</TR>";			   
				frmMain.g_oUtils.WriteText(m_strFile,"<TABLE  colspan=2 width='100%' align='center' border='5', height=50 cellpadding='0', cellspacing='0'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TABLE  colspan=2 width=100% align=center border=5 height=50 cellpadding=0 cellspacing=0>"));
				this.AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=2  bgcolor='#FFBF00' width='30%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=2  bgcolor=#FFBF00 width=30% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<FONT SIZE='+1'><B> Treatment " + this.RxCollection.Item(x).RxId + " </B></FONT> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<FONT SIZE=+1><B> Treatment " + this.RxCollection.Item(x).RxId + " </B></FONT>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#cccccc' width='30%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=30% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Item \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Item"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align=cente' colspan=1  bgcolor=#cccccc width='70%' height='30%' valign=center> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=70% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Description \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Description"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));

				//
				//DESCRIPTION
				//
				AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='30%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=30% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Description \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Description"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='left' colspan=1  bgcolor='#ffffff' width='70%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=left colspan=1  bgcolor=#ffffff width=70% height=30% valign=center>"));
				if (this.RxCollection.Item(x).Description.Trim().Length == 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{
					frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(x).Description + " \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.RxCollection.Item(x).Description));
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				//
				//HARVEST METHOD
				//
				AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='30%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=30% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Harvest Method Low Slope\r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Harvest Method Low Slope"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='left' colspan=1  bgcolor='#ffffff' width='70%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=left colspan=1  bgcolor=#ffffff width=70% height=30% valign=center>"));
                if (this.RxCollection.Item(x).HarvestMethodLowSlope.Trim().Length == 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{
                    frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(x).HarvestMethodLowSlope + " \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat(this.RxCollection.Item(x).HarvestMethodLowSlope));
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				//
				//HARVEST METHOD STEEP SLOPE
				//
				AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='30%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=30% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Steep Slope Harvest Method \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Steep Slope Harvest Method"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='left' colspan=1  bgcolor='#ffffff' width='70%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=left colspan=1  bgcolor=#ffffff width=70% height=30% valign=center>"));
				if (this.RxCollection.Item(x).HarvestMethodSteepSlope.Trim().Length == 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{
					frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(x).HarvestMethodSteepSlope + " \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.RxCollection.Item(x).HarvestMethodSteepSlope));
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				AddToLineCount(5);
				//
				//PACKAGE MEMBER
				//
				AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='30%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=30% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Package Member \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Package Member"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='left' colspan=1  bgcolor='#ffffff' width='70%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=left colspan=1  bgcolor=#ffffff width=70% height=30% valign=center>"));
				if (this.RxCollection.Item(x).RxPackageMemberList.Trim().Length == 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "Not assigned to a package \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Not assigned to a package"));
				}
				else
				{
					frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(x).RxPackageMemberList + " \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.RxCollection.Item(x).RxPackageMemberList));
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
                //
				//ASSOCIATED HARVEST COST COLUMNS
				//
				AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=2  bgcolor='#F5E8BF' width='30%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=2  bgcolor=#F5E8BF width=30% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<B>Associated Harvest Cost Columns</B> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<B>Associated Harvest Cost Columns</B>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				AddToLineCount(1);
                if (this.RxCollection.Item(x).ReferenceHarvestCostColumnCollection == null ||
                      this.RxCollection.Item(x).ReferenceHarvestCostColumnCollection.Count == 0)
                {
                    frmMain.g_oUtils.WriteText(m_strFile, "<TR> \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<TR>"));
                    frmMain.g_oUtils.WriteText(m_strFile, "<TD align='center' colspan=2  bgcolor='#cccccc' width='100%' height='30%' valign='center'> \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<TD align=center colspan=2  bgcolor=#cccccc width=100% height=30% valign=center>"));
                    frmMain.g_oUtils.WriteText(m_strFile, "<B>None Defined</B> \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<B>None Defined</B>"));
                    frmMain.g_oUtils.WriteText(m_strFile, "</TD> \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("</TD>"));
                    frmMain.g_oUtils.WriteText(m_strFile, "</TR> \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("</TR>"));
                }
                else
                {
                    frmMain.g_oUtils.WriteText(m_strFile, "<TR> \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<TR>"));



                    this.m_strHtmlPageHeader = "<TR>" +
                                             "<TD align=center colspan=1 " +
                                                "bgcolor=#cccccc width='30%' height='30%' valign=center>Item" +
                                             "</TD>" +
                                             "<TD align=center colspan=1  " +
                                                "bgcolor=#cccccc width='70%' height='30%' valign=center>Description" +
                                             "</TD>" +
                                             "</TR>";


                    frmMain.g_oUtils.WriteText(m_strFile, "<TD align='center' colspan=1  bgcolor='#cccccc' width='30%' height='30%' valign='center'> \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=30% height=30% valign=center>"));
                    frmMain.g_oUtils.WriteText(m_strFile, "Item \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("Item"));
                    frmMain.g_oUtils.WriteText(m_strFile, "</TD> \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("</TD>"));
                    frmMain.g_oUtils.WriteText(m_strFile, "<TD align='center' colspan=1  bgcolor='#cccccc' width='70%' height='30%' valign='center'> \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=70% height=30% valign=center>"));
                    frmMain.g_oUtils.WriteText(m_strFile, "Description \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("Description"));
                    frmMain.g_oUtils.WriteText(m_strFile, "</TD> \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("</TD>"));
                    frmMain.g_oUtils.WriteText(m_strFile, "</TR> \r\n");
                    frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("</TR>"));

                    for (y = 0; y <= this.RxCollection.Item(x).ReferenceHarvestCostColumnCollection.Count - 1; y++)
                    {

                        if (this.RxCollection.Item(x).RxId ==
                             this.RxCollection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).RxId)
                        {
                            m_strLine = "";
                            AddToLineCount(1);
                            //
                            //HARVEST COST COLUMN
                            //
                            frmMain.g_oUtils.WriteText(m_strFile, "<TR> \r\n");
                            frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<TR>"));
                            frmMain.g_oUtils.WriteText(m_strFile, "<TD align='center' colspan=1  bgcolor='#ffffff' width='30%' height='30%' valign='center'> \r\n");
                            frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=30% height=30% valign=center>"));
                            frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).HarvestCostColumn + " \r\n");
                            frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat(this.RxCollection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).HarvestCostColumn));
                            frmMain.g_oUtils.WriteText(m_strFile, "</TD> \r\n");
                            frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("</TD>"));
                            frmMain.g_oUtils.WriteText(m_strFile, "<TD align=left' colspan=1  bgcolor='#ffffff' width='70%' height='30%' valign='center'> \r\n");
                            frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<TD align=left colspan=1  bgcolor=#ffffff width=70% height=30% valign=center>"));

                            if (RxCollection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).Description.Trim().Length == 0)
                            {
                                frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
                                frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("&nbsp"));
                            }
                            else
                            {
                                frmMain.g_oUtils.WriteText(m_strFile, RxCollection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).Description + " \r\n");
                                frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat(RxCollection.Item(x).ReferenceHarvestCostColumnCollection.Item(y).Description));
                            }
                            frmMain.g_oUtils.WriteText(m_strFile, "</TD> \r\n");
                            frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("</TD>"));
                            frmMain.g_oUtils.WriteText(m_strFile, "</TR> \r\n");
                            frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("</TR>"));


                        }
                    }
                    this.m_strHtmlPageHeader = "";
                }
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TABLE>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<BR><BR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<BR><BR>"));
				
			
			}
			
		}
		private void Packages()
		{
			LINESPERPAGE=35;
			int x,y,z,zz;
			this.m_strHtmlCurrentTable="";this.m_strHtmlPageHeader="";
			if (this.PackageHeader.Trim().Length > 0)
			{
				this.NewPage();
				frmMain.g_oUtils.WriteText(m_strFile,"<!--package header--> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<!--package header-->"));

				frmMain.g_oUtils.WriteText(m_strFile,"<FONT SIZE='+1'><B> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<FONT SIZE=+1><B>"));

				frmMain.g_oUtils.WriteText(m_strFile,"<CENTER>" + PackageHeader + "</CENTER><BR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<CENTER>" + PackageHeader + "</CENTER><BR>"));
				m_intLineCount=m_intLineCount + 3;
			}
			if (this.RxPackageCollection.Count==0)
			{
				frmMain.g_oUtils.WriteText(m_strFile,"<CENTER> No Packages Defined</CENTER><BR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,JavaScriptHtmlFormat("<CENTER>No Packages Defined</CENTER><BR>"));
				m_intLineCount=m_intLineCount + 1;
			}
			for (x=0;x<=this.RxPackageCollection.Count-1;x++)
			{
				this.m_strHtmlPageHeader="";
				//
				//PACKAGE ID
				//
				if (x !=0)
				{
					m_intLineCount = LINESPERPAGE;   //new page for every package
					this.AddToLineCount(1);
				}
				
				this.m_strHtmlCurrentTable="<TABLE  colspan=5 width=100% align=center border=5 height=50 cellpadding=0 cellspacing=0>";
				this.m_strHtmlPageHeader="<TR>" + 
					"<TD align=center colspan=1  bgcolor=#cccccc width='30%' height='30%' valign=center>Item" + 
					"</TD>" + 
					"<TD align=center colspan=4  bgcolor=#cccccc width='70%' height='30%' valign=center>Description" + 
					"</TD>" + 
					"</TR>";			   
				frmMain.g_oUtils.WriteText(m_strFile,"<TABLE  colspan=5 width='100%' align='center' border='5', height=50 cellpadding='0', cellspacing='0'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TABLE  colspan=5 width=100% align=center border=5 height=50 cellpadding=0 cellspacing=0>"));
				this.AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=5  bgcolor='#FFBF00' width='30%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=5  bgcolor=#FFBF00 width=30% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<FONT SIZE='+1'><B> Package " + this.RxPackageCollection.Item(x).RxPackageId + " </B></FONT> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<FONT SIZE=+1><B> Package " + this.RxPackageCollection.Item(x).RxPackageId + " </B></FONT>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#cccccc' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Item \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Item"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align=center colspan=4  bgcolor=#cccccc width='90%' height='30%' valign=center> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=4  bgcolor=#cccccc width=90% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Description \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Description"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				//
				//DESCRIPTION
				//
				AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Description \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Description"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align=left' colspan=4  bgcolor='#ffffff' width='90%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=left colspan=4  bgcolor=#ffffff width=90% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).Description.Trim().Length == 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{
					frmMain.g_oUtils.WriteText(m_strFile, this.RxPackageCollection.Item(x).Description + " \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.RxPackageCollection.Item(x).Description));
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				//
				//FVS CYCLE LENGTH
				//
				AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Treatment Cycle Length \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Treatment Cycle Length"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='left' colspan=4  bgcolor='#ffffff' width='90%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=left colspan=4  bgcolor=#ffffff width=90% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile, "Every " + Convert.ToString(this.RxPackageCollection.Item(x).RxCycleLength) + " years \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Every " + Convert.ToString(this.RxPackageCollection.Item(x).RxCycleLength) + " years"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				//
				//TREATMENT SCHEDULE
				//
				AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=5  bgcolor='#F5E8BF' width='30%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=5	  bgcolor=#F5E8BF width=30% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<B>Treatment Schedule</B> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<B>Treatment Schedule</B>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				AddToLineCount(1);
				this.m_strHtmlPageHeader="<TR>" + 
					"<TD align=center colspan=1 " + 
					"bgcolor=#cccccc width='10%' height='30%' valign=center>Year" + 
					"</TD>" + 
					"<TD align=center colspan=1  " + 
					"bgcolor=#cccccc width='10%' height='30%' valign=center>Rx" + 
					"</TD>" +
					"<TD align=center colspan=1  " + 
					"bgcolor=#cccccc width='10%' height='30%' valign=center>Harvest Method" + 
					"</TD>" + 
					"<TD align=center colspan=1  " + 
					"bgcolor=#cccccc width='10%' height='30%' valign=center>Steep Slope Harvest Method" + 
					"</TD>" + 
					"<TD align=center colspan=1  " + 
					"bgcolor=#cccccc width='60%' height='30%' valign=center>Description" + 
					"</TD>" + 
					"</TR>";
					
				//start row
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				//year column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#cccccc' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Year \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Year"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//rx column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#cccccc' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Rx \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Rx"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//harvest method column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#cccccc' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Harvest Method \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Harvest Method"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//steep slope harvest method column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#cccccc' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Steep Slope Harvest Method \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Steep Slope Harvest Method"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//description column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#cccccc' width='60%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=60% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Description \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Description"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//end row
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));


				//start row
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				//year 00 column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile, "00 \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("00"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//rx column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear1Rx.Trim().Length== 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{
					frmMain.g_oUtils.WriteText(m_strFile, this.RxPackageCollection.Item(x).SimulationYear1Rx + " \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(RxPackageCollection.Item(x).SimulationYear1Rx));
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//harvest method column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear1Rx.Trim().Length== 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{

					for (y=0;y<=this.RxCollection.Count-1;y++)
					{
						if (this.RxCollection.Item(y).RxId.Trim() == 
							this.RxPackageCollection.Item(x).SimulationYear1Rx.Trim())
						{
                            if (RxCollection.Item(y).HarvestMethodLowSlope.Trim().Length == 0)
							{
								frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
							}
							else
							{
                                frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(y).HarvestMethodLowSlope + "\r\n");
                                frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat(this.RxCollection.Item(y).HarvestMethodLowSlope));
							}
							break;
						}
																				
					}					
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//steep slope harvest method column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear1Rx.Trim().Length== 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{

					for (y=0;y<=this.RxCollection.Count-1;y++)
					{
						if (this.RxCollection.Item(y).RxId.Trim() == 
							this.RxPackageCollection.Item(x).SimulationYear1Rx.Trim())
						{
							if (RxCollection.Item(y).HarvestMethodSteepSlope.Trim().Length == 0)
							{
								frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
							}
							else
							{
								frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(y).HarvestMethodSteepSlope + "\r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.RxCollection.Item(y).HarvestMethodSteepSlope));
							}
							break;
						}
																				
					}					
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//rx description
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='left' colspan=1  bgcolor='#ffffff' width='60%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=left colspan=1  bgcolor=#ffffff width=60% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear1Rx.Trim().Length== 0)
					
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				//GP else if (this.RxPackageCollection.Item(x).SimulationYear1Rx.Trim()=="GP")
				//GP {
				//GP 	frmMain.g_oUtils.WriteText(m_strFile, "Growth Projection \r\n");
				//GP 	frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Growth Projection"));
				//GP }
				else
				{
					for (y=0;y<=this.RxCollection.Count-1;y++)
					{
						if (this.RxCollection.Item(y).RxId.Trim() == 
							this.RxPackageCollection.Item(x).SimulationYear1Rx.Trim())
						{
							if (RxCollection.Item(y).Description.Trim().Length == 0)
							{
								frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
							}
							else
							{
								frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(y).Description + "\r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.RxCollection.Item(y).Description));
							}
							break;
						}
																				
					}
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//end row
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				AddToLineCount(1);
				//start row
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				//year 10 column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgco=#ffffff width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile, Convert.ToString(this.RxPackageCollection.Item(x).RxCycleLength * 1).PadLeft(2,'0') + " \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(Convert.ToString(this.RxPackageCollection.Item(x).RxCycleLength * 1).PadLeft(2,'0')));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//rx column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear2Rx.Trim().Length== 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{
					frmMain.g_oUtils.WriteText(m_strFile, this.RxPackageCollection.Item(x).SimulationYear2Rx + " \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(RxPackageCollection.Item(x).SimulationYear2Rx));
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//harvest method low slope column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear2Rx.Trim().Length== 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{

					for (y=0;y<=this.RxCollection.Count-1;y++)
					{
						if (this.RxCollection.Item(y).RxId.Trim() == 
							this.RxPackageCollection.Item(x).SimulationYear2Rx.Trim())
						{
                            if (RxCollection.Item(y).HarvestMethodLowSlope.Trim().Length == 0)
							{
								frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
							}
							else
							{
                                frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(y).HarvestMethodLowSlope + "\r\n");
                                frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat(this.RxCollection.Item(y).HarvestMethodLowSlope));
							}
							break;
						}
																				
					}					
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//steep slope harvest method column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear2Rx.Trim().Length== 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{

					for (y=0;y<=this.RxCollection.Count-1;y++)
					{
						if (this.RxCollection.Item(y).RxId.Trim() == 
							this.RxPackageCollection.Item(x).SimulationYear2Rx.Trim())
						{
							if (RxCollection.Item(y).HarvestMethodSteepSlope.Trim().Length == 0)
							{
								frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
							}
							else
							{
								frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(y).HarvestMethodSteepSlope + "\r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.RxCollection.Item(y).HarvestMethodSteepSlope));
							}
							break;
						}
																				
					}					
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));

				//rx description
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='left' colspan=1  bgcolor='#ffffff' width='60%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=left colspan=1  bgcolor=#ffffff width=60% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear2Rx.Trim().Length== 0)
					
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				//GP else if (this.RxPackageCollection.Item(x).SimulationYear2Rx.Trim()=="GP")
				//GP {
				//GP	frmMain.g_oUtils.WriteText(m_strFile, "Growth Projection \r\n");
				//GP	frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Growth Projection"));
				//GP }
				else
				{
					for (y=0;y<=this.RxCollection.Count-1;y++)
					{
						if (this.RxCollection.Item(y).RxId.Trim() == 
							this.RxPackageCollection.Item(x).SimulationYear2Rx.Trim())
						{
							frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(y).Description + "\r\n");
							frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.RxCollection.Item(y).Description));
							break;
						}
																				
					}
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//end row
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));

				AddToLineCount(1);
				//start row
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				//year 20 column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile, Convert.ToString(this.RxPackageCollection.Item(x).RxCycleLength * 2) + " \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(Convert.ToString(this.RxPackageCollection.Item(x).RxCycleLength * 2)));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//rx column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear3Rx.Trim().Length== 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{
					frmMain.g_oUtils.WriteText(m_strFile, this.RxPackageCollection.Item(x).SimulationYear3Rx + " \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(RxPackageCollection.Item(x).SimulationYear3Rx));
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//harvest method column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear3Rx.Trim().Length== 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{

					for (y=0;y<=this.RxCollection.Count-1;y++)
					{
						if (this.RxCollection.Item(y).RxId.Trim() == 
							this.RxPackageCollection.Item(x).SimulationYear3Rx.Trim())
						{
							if (RxCollection.Item(y).HarvestMethodLowSlope.Trim().Length == 0)
							{
								frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
							}
							else
							{
								frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(y).HarvestMethodLowSlope + "\r\n");
                                frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat(this.RxCollection.Item(y).HarvestMethodLowSlope));
							}
							break;
						}
																				
					}					
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//steep slope harvest method column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear3Rx.Trim().Length== 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{

					for (y=0;y<=this.RxCollection.Count-1;y++)
					{
						if (this.RxCollection.Item(y).RxId.Trim() == 
							this.RxPackageCollection.Item(x).SimulationYear3Rx.Trim())
						{
							if (RxCollection.Item(y).HarvestMethodSteepSlope.Trim().Length == 0)
							{
								frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
							}
							else
							{
								frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(y).HarvestMethodSteepSlope + "\r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.RxCollection.Item(y).HarvestMethodSteepSlope));
							}
							break;
						}
																				
					}					
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//rx description
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='left' colspan=1  bgcolor='#ffffff' width='60%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=left colspan=1  bgcolor=#ffffff width=60% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear3Rx.Trim().Length== 0)
					
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				//GP else if (this.RxPackageCollection.Item(x).SimulationYear3Rx.Trim()=="GP")
				//GP {
				//GP	frmMain.g_oUtils.WriteText(m_strFile, "Growth Projection \r\n");
				//GP 	frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Growth Projection"));
				//GP }
				else
				{
					for (y=0;y<=this.RxCollection.Count-1;y++)
					{
						if (this.RxCollection.Item(y).RxId.Trim() == 
							this.RxPackageCollection.Item(x).SimulationYear3Rx.Trim())
						{
							frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(y).Description + "\r\n");
							frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.RxCollection.Item(y).Description));
							break;
						}
																				
					}
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//end row
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				
				AddToLineCount(1);
				//start row
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				//year 30 column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile, Convert.ToString(this.RxPackageCollection.Item(x).RxCycleLength * 3) + " \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(Convert.ToString(this.RxPackageCollection.Item(x).RxCycleLength * 3)));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//rx column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear4Rx.Trim().Length== 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{
					frmMain.g_oUtils.WriteText(m_strFile, this.RxPackageCollection.Item(x).SimulationYear4Rx + " \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(RxPackageCollection.Item(x).SimulationYear4Rx));
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//harvest method column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear4Rx.Trim().Length== 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{

					for (y=0;y<=this.RxCollection.Count-1;y++)
					{
						if (this.RxCollection.Item(y).RxId.Trim() == 
							this.RxPackageCollection.Item(x).SimulationYear4Rx.Trim())
						{
                            if (RxCollection.Item(y).HarvestMethodLowSlope.Trim().Length == 0)
							{
								frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
							}
							else
							{
                                frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(y).HarvestMethodLowSlope + "\r\n");
                                frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat(this.RxCollection.Item(y).HarvestMethodLowSlope));
							}
							break;
						}
																				
					}					
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//steep slope harvest method column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear4Rx.Trim().Length== 0)
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				else
				{

					for (y=0;y<=this.RxCollection.Count-1;y++)
					{
						if (this.RxCollection.Item(y).RxId.Trim() == 
							this.RxPackageCollection.Item(x).SimulationYear4Rx.Trim())
						{
							if (RxCollection.Item(y).HarvestMethodSteepSlope.Trim().Length == 0)
							{
								frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
							}
							else
							{
								frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(y).HarvestMethodSteepSlope + "\r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.RxCollection.Item(y).HarvestMethodSteepSlope));
							}
							break;
						}
																				
					}					
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//rx description
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='left' colspan=1  bgcolor='#ffffff' width='60%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=left colspan=1  bgcolor=#ffffff width=60% height=30% valign=center>"));
				if (this.RxPackageCollection.Item(x).SimulationYear4Rx.Trim().Length== 0)
					
				{
					frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
				}
				//GP else if (this.RxPackageCollection.Item(x).SimulationYear4Rx.Trim()=="GP")
				//GP {
				//GP	frmMain.g_oUtils.WriteText(m_strFile, "Growth Projection \r\n");
				//GP	frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Growth Projection"));
				//GP }
				else
				{
					for (y=0;y<=this.RxCollection.Count-1;y++)
					{
						if (this.RxCollection.Item(y).RxId.Trim() == 
							this.RxPackageCollection.Item(x).SimulationYear4Rx.Trim())
						{
							if (this.RxCollection.Item(y).Description.Trim().Length == 0)
							{
								frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("&nbsp"));
							}
							else
							{
								frmMain.g_oUtils.WriteText(m_strFile, this.RxCollection.Item(y).Description + "\r\n");
								frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.RxCollection.Item(y).Description));
							}
							break;
						}
																				
					}
				}
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//end row
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
                //
				//FVS HARVEST COST COLUMNS
				//
				AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=5  bgcolor='#F5E8BF' width='30%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=5  bgcolor=#F5E8BF width=30% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<B>Associated Harvest Cost Columns</B> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<B>Associated Harvest Cost Columns</B>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				this.m_strHtmlPageHeader="<TR>" + 
					"<TD align=center colspan=1 " + 
					"bgcolor=#cccccc width='10%' height='30%' valign=center>Rx" + 
					"</TD>" + 
					"<TD align=center colspan=1  " + 
					"bgcolor=#cccccc width='10%' height='30%' valign=center>Simulation Cycle" + 
					"</TD>" +
					"<TD align=center colspan=1 " + 
					"bgcolor=#cccccc width='10%' height='30%' valign=center>Harvest Cost Column" + 
					"</TD>" + 
					"<TD align=left colspan=2  " + 
					"bgcolor=#cccccc width='70%' height='30%' valign=center>Description" + 
					"</TD>" + 
					"</TR>"; 
				AddToLineCount(1);
				//start row
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				//rx column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#cccccc' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Rx \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Rx"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//simulation year column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#cccccc' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Simulation Cycle \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Simulation Cycle"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//Harvest Cost Column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#cccccc' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Harvest Cost Column \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Harvest Cost Column"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//description column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='left' colspan=2  bgcolor='#cccccc' width='70%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=left colspan=2  bgcolor=#cccccc width=70% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Description \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Description"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//end row
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				//see if any harvest cost columns in the package
				bool bHarvestColumnsExist=false;
				for (y=0;y<=RxCollection.Count-1;y++)
                {
			   	   if (RxCollection.Item(y).ReferenceHarvestCostColumnCollection != null && 
			   	       RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Count > 0)
			   	    {
			   	    	 
			   	   				if (this.RxCollection.Item(y).RxId.Trim() == 
												this.RxPackageCollection.Item(x).SimulationYear1Rx.Trim() ||
												this.RxCollection.Item(y).RxId.Trim() == 
												this.RxPackageCollection.Item(x).SimulationYear2Rx.Trim() ||
												this.RxCollection.Item(y).RxId.Trim() == 
												this.RxPackageCollection.Item(x).SimulationYear3Rx.Trim() ||
												this.RxCollection.Item(y).RxId.Trim() == 
												this.RxPackageCollection.Item(x).SimulationYear4Rx.Trim())
										{
											 bHarvestColumnsExist=true; break;
										}
			   	       
			   	       
			   	    }
			   	    if (bHarvestColumnsExist==true) break;
                }
				if (bHarvestColumnsExist==false)
				{
					frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
					frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=5  bgcolor='#cccccc' width='100%' height='30%' valign='center'> \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=5  bgcolor=#cccccc width=100% height=30% valign=center>"));
					frmMain.g_oUtils.WriteText(m_strFile,"<B>None Defined</B> \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<B>None Defined</B>"));
					frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
					frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
					frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				}
				else
				{
					int intCycle;
					 for (intCycle=1;intCycle<=4;intCycle++)
					 {
					 	  bHarvestColumnsExist=false;
					 	  if (intCycle==1)
					 	  {
					 	  	for (y=0;y<=RxCollection.Count-1;y++)
			   				{
			   					if (RxPackageCollection.Item(x).SimulationYear1Rx.Trim() == 
			   					    RxCollection.Item(y).RxId.Trim())
			   					{
                                    if (RxCollection.Item(y).ReferenceHarvestCostColumnCollection != null &&
                                    RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Count > 0)
                                    {
                                        bHarvestColumnsExist = true;
                                    }
                                    break;
			   					}
			   				}
					 	  }
					 	  else if (intCycle==2)
					 	  {
					 	  	for (y=0;y<=RxCollection.Count-1;y++)
			   				{
			   					if (RxPackageCollection.Item(x).SimulationYear2Rx.Trim() == 
			   					    RxCollection.Item(y).RxId.Trim())
			   					{
                                    if (RxCollection.Item(y).ReferenceHarvestCostColumnCollection != null &&
                                    RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Count > 0)
                                    {
                                        bHarvestColumnsExist = true;
                                    }
                                    break;
			   					}
			   				}
					 	  }
					 	  else if (intCycle==3)
					 	  {
					 	  	for (y=0;y<=RxCollection.Count-1;y++)
			   				{
			   					if (RxPackageCollection.Item(x).SimulationYear3Rx.Trim() == 
			   					    RxCollection.Item(y).RxId.Trim())
			   					{
                                    if (RxCollection.Item(y).ReferenceHarvestCostColumnCollection != null &&
                                    RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Count > 0)
                                    {
                                        bHarvestColumnsExist = true;
                                    }
                                    break;
			   					}
			   				}
					 	  }
					 	  else
					 	  {
					 	  	for (y=0;y<=RxCollection.Count-1;y++)
			   				{
			   					if (RxPackageCollection.Item(x).SimulationYear4Rx.Trim() == 
			   					    RxCollection.Item(y).RxId.Trim())
			   					{
                                    if (RxCollection.Item(y).ReferenceHarvestCostColumnCollection != null &&
                                        RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Count > 0)
                                    {
                                        bHarvestColumnsExist = true;
                                    }
                                    break;
			   					}
			   				}
					 	  }
					 	  if (bHarvestColumnsExist)
                          {
                              string strHarvestCostColumnList = "";
                              string strHarvestCostColumnDesc = "";
                               this.AddToLineCount(1);
                              frmMain.g_oUtils.WriteText(m_strFile, "<TR> \r\n");
                              frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<TR>"));
                              //rx
                              frmMain.g_oUtils.WriteText(m_strFile, "<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
                              frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
                              if (RxCollection.Item(y).RxId.Trim().Trim().Length == 0)
                              {
                                  frmMain.g_oUtils.WriteText(m_strFile, "NA \r\n");
                                  frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("NA"));
                              }
                              else
                              {
                                  frmMain.g_oUtils.WriteText(m_strFile, RxCollection.Item(y).RxId.Trim() + " \r\n");
                                  frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat(RxCollection.Item(y).RxId.Trim()));
                              }
                              frmMain.g_oUtils.WriteText(m_strFile, "</TD> \r\n");
                              frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("</TD>"));
                              //SimYear
                              frmMain.g_oUtils.WriteText(m_strFile, "<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
                              frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
                              if (intCycle.ToString().Trim().Length == 0)
                              {
                                  frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
                                  frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("&nbsp"));
                              }
                              else
                              {
                                  frmMain.g_oUtils.WriteText(m_strFile, intCycle.ToString().Trim() + " \r\n");
                                  frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat(intCycle.ToString().Trim()));
                              }
                              frmMain.g_oUtils.WriteText(m_strFile, "</TD> \r\n");
                              frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("</TD>"));
                              for (z = 0; z <= RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Count - 1; z++)
                              {
                                  if (RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Item(z).HarvestCostColumn.Trim().Length > 0)
                                  {
                                      if (strHarvestCostColumnList.Trim().Length > 0)
                                          strHarvestCostColumnList = strHarvestCostColumnList + "<br>" + RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Item(z).HarvestCostColumn.Trim();
                                      else
                                          strHarvestCostColumnList = RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Item(z).HarvestCostColumn.Trim();


                                      
                                      if (RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Item(z).Description.Trim().Length > 0)
                                      {
                                          if (strHarvestCostColumnDesc.Trim().Length > 0)
                                             strHarvestCostColumnDesc = strHarvestCostColumnDesc + "<br>" + 
                                                 RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Item(z).HarvestCostColumn.Trim() + "=" +
                                                 RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Item(z).Description;
                                          else
                                              strHarvestCostColumnDesc = 
                                                 RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Item(z).HarvestCostColumn.Trim() + "=" +
                                                 RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Item(z).Description;

                                      }
                                      else
                                      {
                                          if (strHarvestCostColumnDesc.Trim().Length > 0)
                                                strHarvestCostColumnDesc = strHarvestCostColumnDesc + "<br>" +
                                                RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Item(z).HarvestCostColumn.Trim() + "=No Description Given";
                                          else
                                              strHarvestCostColumnDesc = 
                                                  RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Item(z).HarvestCostColumn.Trim() + "=No Description Given";

                                      }
                                  }
                              }
                              
                              //Harvest Cost Column
                              frmMain.g_oUtils.WriteText(m_strFile, "<TD align='center' colspan=1  bgcolor='#ffffff' width='10%' height='30%' valign='center'> \r\n");
                              frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#ffffff width=10% height=30% valign=center>"));
                              if (strHarvestCostColumnList.Trim().Length == 0)
                              {
                                  frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
                                  frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("&nbsp"));
                              }
                              else
                              {
                                  frmMain.g_oUtils.WriteText(m_strFile, strHarvestCostColumnList + " \r\n");
                                  frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat(strHarvestCostColumnList));
                              }
                              frmMain.g_oUtils.WriteText(m_strFile, "</TD> \r\n");
                              frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("</TD>"));
                              //Harvest Cost Description
                              frmMain.g_oUtils.WriteText(m_strFile, "<TD align='left' colspan=2  bgcolor='#ffffff' width='70%' height='30%' valign='center'> \r\n");
                              frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("<TD align=left colspan=2  bgcolor=#ffffff width=70% height=30% valign=center>"));
                              if (strHarvestCostColumnDesc.Trim().Length == 0)
                              {
                                  frmMain.g_oUtils.WriteText(m_strFile, "&nbsp \r\n");
                                  frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("&nbsp"));
                              }
                              else
                              {
                                  frmMain.g_oUtils.WriteText(m_strFile,strHarvestCostColumnDesc + " \r\n");
                                  frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat(strHarvestCostColumnDesc));
                              }
                              frmMain.g_oUtils.WriteText(m_strFile, "</TD> \r\n");
                              frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("</TD>"));
                              frmMain.g_oUtils.WriteText(m_strFile, "</TR> \r\n");
                              frmMain.g_oUtils.WriteText(m_strJavaScriptFile, this.JavaScriptHtmlFormat("</TR>"));
                              this.AddToLineCount(1 + RxCollection.Item(y).ReferenceHarvestCostColumnCollection.Count);
						  }
					}	 
						
					this.m_strHtmlPageHeader="";
				}
				//
				//FVS COMMANDS
				//
				AddToLineCount(1);
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=5  bgcolor='#F5E8BF' width='30%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=5  bgcolor=#F5E8BF width=30% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<B>Associated FVS Commands</B> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<B>Associated FVS Commands</B>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
				this.m_strHtmlPageHeader="<TR>" + 
					"<TD align=center colspan=1 " + 
					"bgcolor=#cccccc width='10%' height='30%' valign=center>Rx" + 
					"</TD>" + 
					"<TD align=center colspan=1  " + 
					"bgcolor=#cccccc width='10%' height='30%' valign=center>Simulation Cycle" + 
					"</TD>" +
					"<TD align=center colspan=1 " + 
					"bgcolor=#cccccc width='10%' height='30%' valign=center>FVS Command" + 
					"</TD>" + 
					"<TD align=center colspan=2  " + 
					"bgcolor=#cccccc width='70%' height='30%' valign=center>Description" + 
					"</TD>" + 
					"</TR>"; 
				AddToLineCount(1);
				//start row
				frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				//rx column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#cccccc' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Rx \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Rx"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//simulation year column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#cccccc' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Simulation Cycle \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Simulation Cycle"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//FVS Command
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=1  bgcolor='#cccccc' width='10%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=1  bgcolor=#cccccc width=10% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"FVS Command \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("FVS Command"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//description column
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=2  bgcolor='#cccccc' width='70%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=2  bgcolor=#cccccc width=70% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"Description \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("Description"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				//end row
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));
			    frmMain.g_oUtils.WriteText(m_strFile,"<TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TR>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<TD align='center' colspan=5  bgcolor='#cccccc' width='100%' height='30%' valign='center'> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<TD align=center colspan=5  bgcolor=#cccccc width=100% height=30% valign=center>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<B>None Defined</B> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<B>None Defined</B>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TD> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TD>"));
				frmMain.g_oUtils.WriteText(m_strFile,"</TR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TR>"));

					
				frmMain.g_oUtils.WriteText(m_strFile,"</TABLE> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TABLE>"));
				frmMain.g_oUtils.WriteText(m_strFile,"<BR><BR> \r\n");
				frmMain.g_oUtils.WriteText(m_strJavaScriptFile,this.JavaScriptHtmlFormat("<BR><BR>"));
				
			
			}
			
		}
		private void AddToLineCount(int intLines)
		{
			m_intLineCount=m_intLineCount + intLines;
			if (m_intLineCount >= LINESPERPAGE)
			{
				NewPage();

				
			}
			else
			{
				m_bNewPage=false;
			}
			
		}
		private void NewPage()
		{
			m_bNewPage=true;
			m_intLineCount=0;
			if (this.m_strHtmlCurrentTable.Trim().Length > 0)
			{
				//terminate the table
				frmMain.g_oUtils.WriteText(this.m_strJavaScriptFile,this.JavaScriptHtmlFormat("</TABLE>"));
			}
				
			frmMain.g_oUtils.WriteText(this.m_strJavaScriptFile,this.JavaScriptHtmlFormat("<br class=pagebreak />"));
			if (this.m_strHtmlCurrentTable.Trim().Length > 0)
			{
					   
				frmMain.g_oUtils.WriteText(this.m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.m_strHtmlCurrentTable));


			}
			if (this.m_strHtmlPageHeader.Trim().Length > 0)
			{
					
				frmMain.g_oUtils.WriteText(this.m_strJavaScriptFile,this.JavaScriptHtmlFormat(this.m_strHtmlPageHeader));
				m_intLineCount=2;
					
			}
		}
		private string JavaScriptHtmlFormat(string p_strLine)
		{

		   string strLine = p_strLine;
		   strLine = strLine.Replace("\\","\\\\");
		   //strLine = "output.write(\"" + strLine + " linefeed \"); \r\n";
		   strLine = FixString(strLine,"'","&#34;");
		  // strLine = "output.write(\"" + strLine + " linefeed \"); \r\n";
			 strLine = "output.write(\"" + strLine + "\");" +  "\r\n";
			return strLine;
		   

		}
		private string JavaScriptFormat(string p_strLine)
		{

			string strLine = p_strLine;
			strLine = FixString(strLine,"'","\"");
			//strLine = "output.write(\"" + strLine + " linefeed \"); \r\n";
			return strLine;
		   

		}
		public string FixString(string SourceString , string StringToReplace, string StringReplacement)
		{
			SourceString = SourceString.Replace(StringToReplace, StringReplacement);
			return(SourceString);
		}

		public string WindowTitle
		{
			get {return this._strWindowTitle;}
			set {_strWindowTitle=value;}
		}
		public string ReportHeader
		{
			get {return this._strReportHeader;}
			set {_strReportHeader=value;}
		}
		public string TreatmentHeader
		{
			get {return this._strTreatmentHeader;}
			set {_strTreatmentHeader=value;}
		}
		public string PackageHeader
		{
			get {return _strPackageHeader;}
			set {_strPackageHeader=value;}
		}
		public string ProjectName
		{
			get {return this._strProjectName;}
			set {_strProjectName=value;}
		}
		public FIA_Biosum_Manager.RxItem_Collection RxCollection
		{
			get {return  _RxItem_Collection;}
			set { _RxItem_Collection=value;}
		}
		public RxPackageItem_Collection RxPackageCollection
		{
			get {return this._RxPackage_Collection;}
			set {_RxPackage_Collection=value;}
		}
		public bool ProcessTreatments
		{
			get {return _bTreatments;}
			set {_bTreatments=value;}
		}
		public bool ProcessPackages
		{
			get {return _bPackages;}
			set {_bPackages=value;}
		}
		

	
	}
}

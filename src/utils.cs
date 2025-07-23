


using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;


// From Winbase.h



namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for utils.
	/// </summary>
	public class utils
	{
		private  bool m_lDone=false;
		private  long m_longFindWindowLike=0;
	    public   int  m_intLevel=0;
        public static string g_strError="";
        public static int g_intError = 0;

		class User32
		{
			//window handle values
			public const uint GW_HWNDFIRST =  0;
			public const uint GW_HWNDLAST = 1;
			public const uint GW_HWNDNEXT = 2;
			public const uint GW_HWNDPREV = 3;
			public const uint GW_OWNER = 4;
			public const uint GW_CHILD = 5;
			public const uint GW_ENABLEDPOPUP = 6;
			internal const byte VK_CAPITAL = 0x14;
            public const uint SHOWWINDOW_NORMAL = 1;

			[DllImport("user32.dll")]
			public static extern IntPtr  GetWindow(IntPtr hWnd, uint uCmd);
			[DllImport("User32.dll")]
			public static extern IntPtr GetDesktopWindow();
			[DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
			public static extern long GetWindowText(IntPtr hwnd,  System.Text.StringBuilder text, int cch);
			[DllImport("User32.Dll")]
			public static extern long GetClassName(IntPtr h, System.Text.StringBuilder s, int nMaxCount);
			[DllImport("User32.Dll")]
			internal static extern ushort GetKeyState(uint nVirtKey);
            [DllImport("User32.Dll")]
            public static extern long BringWindowToTop(IntPtr h);
            [DllImport("User32.Dll")]
            public static extern long SetFocus(IntPtr h);
            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);


		}

		class kernel32
		{
			public enum DriveType : int 
			{
				Unknown = 0,	
				NoRoot = 1,
				Removable = 2,
				Localdisk = 3,
				Network = 4,
				CD = 5,
				RAMDrive = 6
			}

			[DllImport("kernel32.dll", CharSet = CharSet.Auto)]
			public static extern int GetDriveType(string lpRootPathName);
		}
		public utils()
		{
			//
			// TODO: Add constructor logic here
			//
		}
	    ~utils()
		{
          
		}
        public void SetFocus(long handle)
        {
            User32.SetFocus((System.IntPtr)handle);
        }
        public void BringWindowToTop(long handle)
        {
            User32.BringWindowToTop((System.IntPtr)handle);
        }
        public void ShowWindow(long handle, int intWindowState)
        {
            User32.ShowWindow((System.IntPtr)handle, intWindowState);
        }
		public string getFileName(string p_strFullPath)
		{
			string strFileName = "";
			for (int x = p_strFullPath.Length - 1; x >= 0; x--)
			{
				
				if (p_strFullPath.Substring(x,1) == "\\" 
					  || p_strFullPath.Substring(x,1) == "/" ) 

				{
                     break;
				}
				else 
				{
					strFileName =  p_strFullPath.Substring(x,1) + strFileName;
				}
			}
			return strFileName;

		}
		public string getFileNameUsingLastIndexOf(string p_strFullPath)
		{
			return p_strFullPath.Substring(p_strFullPath.LastIndexOf("\\")+1,p_strFullPath.Length - p_strFullPath.LastIndexOf("\\")-1);
		}
		public string getDirectory(string p_strFullPath)
		{
			string strDirectory = "";
			for (int x = p_strFullPath.Length - 1; x >= 0; x--)
			{
				
				if (p_strFullPath.Substring(x,1) == "\\" 
					|| p_strFullPath.Substring(x,1) == "/" ) 

				{
					strDirectory = p_strFullPath.Substring(0, x);
					break;
				}
			}
			return strDirectory;

		}
        public string getDriveLetter(string p_strPath)
        {
            int intAT=0;
            /**********************************************************
			 **see if there is a drive letter in the project directory
			 **********************************************************/
            intAT = p_strPath.IndexOf(":");
            if (intAT > 0)
            {
                /******************************************************
                 **load the drive and root directory into variables
                 ******************************************************/
                return p_strPath.Substring(intAT - 1, 2);
            }
            return "";
        }

		//this will copy files from src directory to destination directory
		//if recursive is true it will copy all files and subdirectories from
		//the source directory
		public static void FileCopy(string srcdir, string destdir, bool recursive)
		{
			System.IO.DirectoryInfo   dir;
			System.IO.FileInfo[]      files;
			System.IO.DirectoryInfo[] dirs;
			string          tmppath;

			//determine if the destination directory exists, if not create it
			if (! System.IO.Directory.Exists(destdir))
			{
				System.IO.Directory.CreateDirectory(destdir);
			}

			dir = new System.IO.DirectoryInfo(srcdir);
            
			//if the source dir doesn't exist, throw
			if (! dir.Exists)
			{
				throw new ArgumentException("source dir doesn't exist -> " + srcdir);
			}

			//get all files in the current dir
			files = dir.GetFiles();

			//loop through each file
			foreach(System.IO.FileInfo file in files)
			{
				//create the path to where this file should be in destdir
				tmppath=System.IO.Path.Combine(destdir, file.Name);                

				//copy file to dest dir
				file.CopyTo(tmppath, false);
			}

			//cleanup
			files = null;
            
			//if not recursive, all work is done
			if (! recursive)
			{
				return;
			}

			//otherwise, get dirs
			dirs = dir.GetDirectories();

			//loop through each sub directory in the current dir
			foreach(System.IO.DirectoryInfo subdir in dirs)
			{
				//create the path to the directory in destdir
				tmppath = System.IO.Path.Combine(destdir, subdir.Name);

				//recursively call this function over and over again
				//with each new dir.
				FileCopy(subdir.FullName, tmppath, recursive);
			}
            
			//cleanup
			dirs = null;
            
			dir = null;
		}

		public string[] getLocalHardDriveList()
		{
			string[] s = System.Environment.GetLogicalDrives();
			string[] strLocalDrives = new string[25];
			//MessageBox.Show(s.Length.ToString());
			int x=0;
			try
			{
				
				for(x=0;x<=24;x++) strLocalDrives[x]="";
				//s = System.Environment.GetLocalDrives();
				int y=0;
				for (x=0; x<=s.Length-1; x++)
				{
					if (kernel32.GetDriveType(s[x]) == 3)
					{
						strLocalDrives[y] = s[x];
						y++;
					}
				}
			}
			catch 
			{
				
			}
			return strLocalDrives;


		}

		public long FindWindowLike(IntPtr hWndStart, string WindowText, string Classname, bool lExact, bool lInit)
		{
			IntPtr hwnd;
			string sWindowText="";
			long longnumber;
			string sClassname="";
			int x=0;
            int CurrentWindowTextLength=0;
			int WindowTextLength = WindowText.Trim().Length;
			char[] charParamWindowText = new char[WindowTextLength];
			charParamWindowText = WindowText.Trim().ToUpper().ToCharArray(0,WindowTextLength);

			if (lInit == true)
			{
				this.m_intLevel = 0;
			}
			if (this.m_intLevel == 0)
			{
				 hWndStart = User32.GetDesktopWindow();
			}
		
			//Increase recursion counter
		    this.m_intLevel++;		
			
            hwnd = User32.GetWindow(hWndStart, User32.GW_CHILD);
			while ((Int32)hwnd != 0 && this.m_longFindWindowLike == 0)
			{
				//Search children by recursion
				longnumber = this.FindWindowLike(hwnd, WindowText, Classname, lExact, false);

				//Get the window text and class name
				System.Text.StringBuilder sb = new System.Text.StringBuilder(255);
				
				//return the length of the name of the window string
				longnumber = User32.GetWindowText(hwnd, sb, sb.Capacity );
				if (sb.ToString().Trim().Length > 0)
				{
				}
				System.Text.StringBuilder sb2 = new System.Text.StringBuilder(255);
				sb2.Append(sClassname);
	            //return the length of the name of the class string
				longnumber = User32.GetClassName(hwnd, sb2,sb2.Capacity);

				
				if (lExact == true)
				{
				    //strings must be an exact match	
					if (sb.ToString().Trim().ToUpper() == WindowText.Trim().ToUpper()) 
					{
						this.m_intLevel = 0;
						this.m_longFindWindowLike = (long)hwnd;
						return this.m_longFindWindowLike;
					}
				
				}
				else 
				{
					if (WindowText.Trim().Length >= sb.ToString().Trim().Length &&
						sb.ToString().Trim().Length > 0)
					{
						
						sWindowText = "";
						sWindowText = sb.ToString().Trim().ToUpper() ;//sb.ToString().Substring(0,(int)WindowText.Trim().Length).ToUpper();
						CurrentWindowTextLength = sWindowText.Trim().Length;
						char[] charCurrentWindowText = new char[CurrentWindowTextLength];
						charCurrentWindowText = sWindowText.Trim().ToCharArray(0,CurrentWindowTextLength); //sb.ToString().ToCharArray(0, x);
                        
						//do a character by character comparison because of 
						//problems with the Substring method of the string class
						for (x=0;x<=WindowTextLength-1;x++)
						{
							if (x > CurrentWindowTextLength-1)
							{
								this.m_lDone=true;
								break;
							}
							
						    if (charCurrentWindowText[x] != charParamWindowText[x]) break;
						   
						 
						}
                        //test to see if the current window is an exact match up to the end 
						//of the passed window string
						if (x > CurrentWindowTextLength || this.m_lDone==true)
						{
							//exact match
							this.m_lDone=true;
							this.m_intLevel = 0;
							this.m_longFindWindowLike = (long)hwnd;
							return this.m_longFindWindowLike;
						}
					}

				}
				if (this.m_longFindWindowLike == 0)
				{
					//Get next child window
					hwnd = User32.GetWindow(hwnd, User32.GW_HWNDNEXT);
				}
			}
			if (m_longFindWindowLike > 0)
			{
				return this.m_longFindWindowLike;
			}
			else
			{
				//Reduce the recursion counter
				this.m_intLevel = this.m_intLevel - 1;
				return 0;
			}
		}
		public string getRandomFile(string strDir, string strExt)
		{
			//int intRandom;
			string strFile;
			string strRandomFile="";
			//int x = 0;
			while (true)
			{
				strRandomFile = "";
				strFile = "fia_biosum_";
				//for (x=1;x<=8;x++)
				//{
					strFile = strFile + Convert.ToString(Convert.ToInt32(this.RandomNumber(1,9999)));
				//}
				if (strDir.Trim().Length > 0)
				{
					if (strDir.Substring(strDir.Length-1,1)=="\\")
					{
						strRandomFile = strDir + strFile;
					}
					else
					{
						strRandomFile = strDir + "\\" + strFile;
					}
				}
				else
				{
					strRandomFile = strFile;
				}
         
				if (strExt.Trim().Length > 0)
				{
					strFile = strFile + "." + strExt;
					strRandomFile = strRandomFile + "." + strExt;
				}

				if (System.IO.File.Exists(strRandomFile) == false)
					break;
     
     
			}
			return strRandomFile;
		}

		public int RandomNumber(int min, int max)
		{
			Random random = new Random();
			return random.Next(min, max); 
		}
		/// <summary>
		/// convert a delimited list to an array
		/// </summary>
		/// <param name="p_strList"></param>
		/// <param name="p_strDelimiter"></param>
		/// <returns></returns>
		public string[] ConvertListToArray(string p_strList,string p_strDelimiter)
		{
			if (p_strList.Trim().Length == 0) return null;
			string[] strArray = p_strList.Split(p_strDelimiter.ToCharArray());
			return strArray;

		}
		/// <summary>
		/// convert an array to a delimited list
		/// </summary>
		/// <param name="p_strArray"></param>
		/// <param name="p_strDelimiter"></param>
		/// <returns></returns>
		public string ConvertArrayToList(string[] p_strArray,string p_strDelimiter)
		{
			if (p_strArray.Length == 0) return "";
			string strList="";
			for (int x=0;x<=p_strArray.Length -1;x++)
			{
				strList = strList + p_strArray[x] + p_strDelimiter;
			}
			if (strList.Trim().Length > 0) strList = strList.Substring(0,strList.Length - 1);
			return strList;
		}
		public void LoadStringToCharArray(string str, ref char[] p_Array, bool bInit,int intBegin)
		{
			int x=0;
			if (bInit == true)
			{
				for (x=0; x <= p_Array.Length-1; x++) p_Array[x] = Convert.ToChar(" ");//Convert.ToChar("\0"); //Convert.ToChar(" ");
			}
			//	p_Array.Initialize();

			int intCount=0;
			for (x=intBegin; x <= str.Length-1;x++)
			{
				p_Array[x] = Convert.ToChar(str.Substring(intCount,1));
				intCount++;
			}
			
		
		}
		public void LoadCharArrayToString(ref string str, char[] p_Array, bool bInit)
		{
			if (bInit==true)
				str="";
			
			for (int x=0; x <= p_Array.Length-1;x++)
			{
				str += p_Array[x];
			}
		}
		//check to see if the string is capable of 
		//math functions by examining each char in the string
		public bool IsTheStringNumeric(string str)
		{
			int x;
			int intDecCnt=0;
			string strNumeric = "-+1234567890.";
			for (x=0;x<=str.Trim().Length-1;x++)
			{
				if (strNumeric.IndexOf(str.Substring(x,1),0) < 0)
				{ 
					return false;
				}
				else
				{
					switch (str.Substring(x,1))
					{
						case "-":
						   if (x > 0) return false;
						   break;
						case "+":
							if (x > 0) return false;
							break;
						case ".":
							if (intDecCnt == 1) return false;
							intDecCnt++;
							break;
					}
				}
			}
			return true;

		}
		public void ShellExecute(string strFileName)
		{
			System.Diagnostics.Process proc = new System.Diagnostics.Process();
			proc.StartInfo.UseShellExecute = true;
			try
			{
				proc.StartInfo.FileName = strFileName;
			}
			catch
			{
			}
			try
			{
				proc.Start();
			}
			catch (Exception caught)
			{
				MessageBox.Show(caught.Message);
			}

		}
		public bool IsCapsLockOn()
		{
			if (0 != (User32.GetKeyState(User32.VK_CAPITAL) & 1))
			{
				return true;
			}
            return false;															 
															
		}
		public void WriteText(string p_strTextFile,string p_strText)
		{
			System.IO.FileStream oTextFileStream;
			System.IO.StreamWriter oTextStreamWriter;

			if (!System.IO.File.Exists(p_strTextFile))
			{
				oTextFileStream = new System.IO.FileStream(p_strTextFile, System.IO.FileMode.Create, 
					System.IO.FileAccess.Write);
			}
			else
			{
				oTextFileStream = new System.IO.FileStream(p_strTextFile, System.IO.FileMode.Append, 
					System.IO.FileAccess.Write);
			}
			
			oTextStreamWriter = new System.IO.StreamWriter(oTextFileStream);
			oTextStreamWriter.Write(p_strText);
			oTextStreamWriter.Close();
			oTextFileStream.Close();
		}
        public void RunProcess(string p_strWorkingFolder, string p_strFile, string p_strFileType)
        {
           

            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            try
            {
                proc.StartInfo.RedirectStandardOutput = false;
                proc.StartInfo.RedirectStandardInput = false;
                proc.StartInfo.RedirectStandardError = false;
                proc.StartInfo.CreateNoWindow = false;
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
               
                proc.StartInfo.WorkingDirectory = p_strWorkingFolder;

                proc.StartInfo.ErrorDialog = false;
                proc.EnableRaisingEvents = false;
                try
                {
                    if (p_strFileType == "BAT" || p_strFileType == "TEMPBAT")
                    {
                        proc.StartInfo.FileName = "cmd.exe";
                        proc.StartInfo.Arguments = "/c START /B /WAIT " + p_strFile;
                    }
                    else if (p_strFileType == "VBS")
                    {
                        proc.StartInfo.FileName = "wscript";
                        proc.StartInfo.Arguments = p_strFile;
                    }
                    else if (p_strFileType == "MSI")
                    {
                        proc.StartInfo.FileName = "msiexec.exe";
                        proc.StartInfo.Arguments = "/I " + p_strFile + " /qn";

                    }
                    else
                    {
                        proc.StartInfo.FileName = p_strFile;
                    }
                    proc.Start();
                    ShowWindow((long)proc.MainWindowHandle, (int)utils.User32.SHOWWINDOW_NORMAL);
                    proc.WaitForExit();
                }
                catch (Exception e)
                {
                    g_intError = -1;
                    g_strError = "Installer.RunProcess e:" + e.Message;
                }
            }
            catch (System.ComponentModel.Win32Exception w32e)
            {
                g_intError = -1;
                g_strError = "Installer.RunProcess w32e:" + w32e.Message;

            }

        }

        public static bool FS_NETWORK_IS_AVAILABLE => FS_NETWORK == FS_NETWORK_STATUS.Available;
        public static bool FS_NETWORK_IS_NOT_AVAILABLE => FS_NETWORK == FS_NETWORK_STATUS.NotAvailable;

        public static FS_NETWORK_STATUS FS_NETWORK = FS_NETWORK_STATUS.NotAvailable;
        public enum FS_NETWORK_STATUS
        {
            Available,
            NotAvailable
        }
        public static FS_NETWORK_STATUS FS_NETWORK_CHECK()
        {
            //FIADB01P: PRODUCTION ORACLE SERVER
            return Ping("fsxrndx0178.fdc.fs.usda.gov");
        }
        public static FS_NETWORK_STATUS Ping(string host)
        {
            FS_NETWORK_STATUS oStatus = FS_NETWORK_STATUS.NotAvailable;
            string strError = "";
            if (host != null && host.Trim().Length > 0)
            {
                try
                {
                    System.Net.NetworkInformation.Ping p = new System.Net.NetworkInformation.Ping();
                    
                    if (p.Send(host, 500).Status == System.Net.NetworkInformation.IPStatus.Success)
                    {
                        oStatus = FS_NETWORK_STATUS.Available;
                    }
                    p.Dispose();
                    p = null;

                }
                catch (Exception e)
                {
                    strError = e.Message;
                }
            }
            return oStatus;
        }

        /// <summary>
        /// This enum is used to categorize column datatypes for parsing values from DbDataReaders.
        /// </summary>
        public enum DataType
        {
            INTEGER,
            DOUBLE,
            BYTE,
            STRING
        }

        private static readonly Dictionary<DataType, Func<object, string>> _typeConverterDict =
            new Dictionary<DataType, Func<object, string>>()
            {
                {DataType.INTEGER, (input) => Convert.ToInt32(input).ToString().Trim()},
                {DataType.DOUBLE, (input) => Convert.ToDouble(input).ToString().Trim()},
                {DataType.BYTE, (input) => Convert.ToByte(input).ToString().Trim()},
                {DataType.STRING, (input) => "'" + input.ToString().Trim() + "'"},
            };

        public static string GetParsedInsertValues(DbDataReader reader, List<Tuple<string, DataType>> columnsAndDataTypes)
        {
            if (columnsAndDataTypes == null || columnsAndDataTypes.Count < 1) return "";
            return string.Join(",", 
                columnsAndDataTypes.Select(pair => reader[pair.Item1] == DBNull.Value ? "null" 
                        : _typeConverterDict[pair.Item2](reader[pair.Item1]))
                .ToList());
        }

        public static string GetParsedInsertValues(DbDataReader reader, List<Tuple<string,string, DataType>> columnsAndDataTypes)
        {
            if (columnsAndDataTypes == null || columnsAndDataTypes.Count < 1) return "";
            return string.Join(",",
                columnsAndDataTypes.Select(pair => reader[pair.Item1] == DBNull.Value ? "null"
                        : _typeConverterDict[pair.Item3](reader[pair.Item1]))
                .ToList());
        }

        /// <summary>
        /// Wraps the supplied string in ` characters. 
        /// This is necessary for certain Access column names.
        /// </summary>
        /// <param name="str">The string to wrap in `.</param>
        /// <returns></returns>
        public static string WrapInBacktick(string str)
        {
            return $@"`{str}`";
        }

        /// <summary>Converts the string representation of an SQLite data type to the equivalent Access data type strings.
		/// NOTE: String values are set to length 255 regardless of their length in the original database</summary>
        /// <param name="dataTypeFromDB">The string representation of an SQLite data type.</param>
        public static string DataTypeConvert(string dataTypeFromDB, bool bTargetSqlite)
        {
            var convertedType = "";
            switch (dataTypeFromDB)
            {
                case "SYSTEM.INT32":
                    convertedType = "INTEGER";
                    break;
                case "SYSTEM.DOUBLE":
                    convertedType = "DOUBLE";
                    break;
                case "SYSTEM.STRING":
                    convertedType = "CHAR(255)";
                    break;
				case "SYSTEM.BYTE":
					convertedType = "CHAR(1)";
					break;
				case "SYSTEM.SINGLE":
					convertedType = "INTEGER";
					break;
				case "SYSTEM.INT64":
					convertedType = "INTEGER";
					break;
                default:
                    convertedType = "UNRECOGNIZED";
                    break;
            }
            // MS Access ODBC driver interprets LONG INTEGER as LONG TEXT; If target is an SQLite table
            // that may be read by MS ACCESS, INT32 needs to be INTEGER
            if (convertedType.Equals("LONG") && bTargetSqlite)
            {
                convertedType = "INTEGER";
            }
            return convertedType;
        }

        /// <summary>Translates "FVSOut.db" column naming conventions into BioSum ones, e.g. SPECIESFIA -> SPECIES.</summary>
        /// <param name="strToCheck">String to check. If it matches, map it. Otherwise return the same string.</param>
        public static string TranslateColumn(string strToCheck)
        {
            // Mapping of FVSOut column names to biosum column names. Add new entries here.
            Dictionary<string, string> sqliteToAccessColNames = new Dictionary<string, string>
            {
                { "SPECIESFIA", "SPECIES" }
            };
            var translatedStr = strToCheck.ToUpper();
            // Map SPECIESFIA to SPECIES. In the future, add other column mappings (e.g. stuff that's different in FVSOUT.db from the target Access Mdbs) here.
            if (sqliteToAccessColNames.Keys.Contains(translatedStr))
            {
                translatedStr = sqliteToAccessColNames[translatedStr];                
            }
            return translatedStr;
        }
    }

}

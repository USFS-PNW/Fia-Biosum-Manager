using System;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for ValidateNumericValues.
	/// </summary>
	public class ValidateNumericValues
	{
		 public int m_intError=0;
		 public string m_strError="";
		 private int _intRoundDecimalLength=18;
		 private string _strReturnValue="";
		 private double _dblMaxValue=-1;
		 private double _dblMinValue=-1;
		 private bool  _bTestForMax=false;
		 private bool  _bTestForMaxMin=false;
		 private bool  _bTestForMin=false;
		 private bool  _bNullsAllowed=true;
		 private bool _bMoney=false;
		const int SHORT_LOW = -32768;				//signed 16 bit
		const int SHORT_HIGH = 32767;				//signed 16 bit
		const int INTEGER_LOW = -2147483648;        //signed 32 bit
		const int INTEGER_HIGH = 2147483647;        //signed 32 bit
		const long LONG_LOW =  (long)-9223372036854775808; //signed 64 bit
		const long LONG_HIGH = 9223372036854775807; //signed 64 bit
		const byte  BYTE_LOW = 0;						//unsigned 8 bit integer
		const byte BYTE_HIGH = 255;					//unsigned 8 bit integer
		const sbyte SBYTE_LOW = -128;					//signed 8 bit integer
		const byte SBYTE_HIGH = 127;					//signed 8 bit integer
		const byte UINTEGER_LOW = 0;                  //unsigned 32 bit
		const uint UINTEGER_HIGH = 4294967295;        //unsigned 32 bit
		const byte ULONG_LOW =  0;                    //unsigned 64 bit
		const ulong ULONG_HIGH = 18446744073709551615; //signed 64 bit
		const byte USHORT_LOW = 0;					   //unsigned 16 bit
		const ushort USHORT_HIGH=65535;				   //unsigned 16 bit




		public ValidateNumericValues()
		{
			//
			// TODO: Add constructor logic here
			//
		}
		 public void ValidateDouble(string p_strValue)
		{
			ReturnValue="";
			m_intError=0;
			m_strError="";
			if (IsValueNull(p_strValue)) return;

			//strip out commas
			string strTestValue=p_strValue.Replace(",","");
			if (IsValueNull(strTestValue)) return;

			if (strTestValue.Substring(0,1)=="$") 
			{
				if (strTestValue.Trim().Length == 1)
				{
				   if (IsValueNull("")) return;
				}
				else
				   strTestValue=strTestValue.Substring(1,strTestValue.Length - 1);
			}
			ValidateNumeric(strTestValue);
			
		}
		 public void ValidateSingle(string p_strValue)
		{
			ValidateFloat(p_strValue);
		}
		 public void ValidateFloat(string p_strValue)
		{
			ReturnValue="";
			m_intError=0;
			m_strError="";
			if (IsValueNull(p_strValue)) return;

			//strip out commas
			string strTestValue=p_strValue.Replace(",","");
			if (IsValueNull(strTestValue)) return;

			if (strTestValue.Substring(0,1)=="$") 
			{
				if (strTestValue.Trim().Length == 1)
				{
					if (IsValueNull("")) return;
				}
				else
					strTestValue=strTestValue.Substring(1,strTestValue.Length - 1);
			}
			ValidateNumeric(strTestValue);
			if (m_intError==0)
			{
				try
				{
					float intValue = Convert.ToSingle(strTestValue);
					if (TestForMax)
					{
						if (intValue > MaxValue)
						{
							m_intError=-1;
							m_strError="Value exceeds the allowed maximum value of " + MaxValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					if (m_intError==0 && TestForMin==true)
					{
						if (intValue < MinValue)
						{
							m_intError=-1;
							m_strError="Value is less than the allowed minimum value of " + MinValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					

					
				}
				catch (Exception e)
				{
					m_intError=-1;
					m_strError = e.Message;
					MessageBox.Show(e.Message);
				}

			}
		}
		 public void ValidateShort(string p_strValue)
		{
			ReturnValue="";
			m_intError=0;
			m_strError="";
			if (IsValueNull(p_strValue)) return;

			//strip out commas
			string strTestValue=p_strValue.Replace(",","");
			if (IsValueNull(strTestValue)) return;

			if (strTestValue.Substring(0,1)=="$") 
			{
				if (strTestValue.Trim().Length == 1)
				{
					if (IsValueNull("")) return;
				}
				else
					strTestValue=strTestValue.Substring(1,strTestValue.Length - 1);
			}
			ValidateNumeric(strTestValue);
			if (m_intError==0)
			{
				try
				{
					short intShort = Convert.ToInt16(strTestValue);
					
				}
				catch (Exception e)
				{
					m_intError=-1;
					m_strError = e.Message;
					MessageBox.Show(e.Message);
				}

			}
		}
		 public void ValidateInteger(string p_strValue)
		{
			ReturnValue="";
			m_intError=0;
			m_strError="";
			if (IsValueNull(p_strValue)) return;

			//strip out commas
			string strTestValue=p_strValue.Replace(",","");
			if (IsValueNull(strTestValue)) return;

			if (strTestValue.Substring(0,1)=="$") 
			{
				if (strTestValue.Trim().Length == 1)
				{
					if (IsValueNull("")) return;
				}
				else
					strTestValue=strTestValue.Substring(1,strTestValue.Length - 1);
			}

			
			ValidateNumeric(strTestValue);
			if (m_intError==0)
			{
				try
				{
					int intInteger = Convert.ToInt32(strTestValue);

                    if (TestForMin == true)
                    {
                        if ((double)intInteger < (double)MinValue)
                        {
                            m_intError = -1;
                            m_strError = "Value is less than the allowed minimum value of " + MinValue.ToString();
                            MessageBox.Show(m_strError, "FIA Biosum");
                        }
                    }

                }
				catch (Exception e)
				{
					m_intError=-1;
					m_strError = e.Message;
					MessageBox.Show(e.Message);
				}

			}
		}
		 public void ValidateLong(string p_strValue)
		{
			ReturnValue="";
			m_intError=0;
			m_strError="";
			if (IsValueNull(p_strValue)) return;

			//strip out commas
			string strTestValue=p_strValue.Replace(",","");
			if (IsValueNull(strTestValue)) return;

			if (strTestValue.Substring(0,1)=="$") 
			{
				if (strTestValue.Trim().Length == 1)
				{
					if (IsValueNull("")) return;
				}
				else
					strTestValue=strTestValue.Substring(1,strTestValue.Length - 1);
			}
			ValidateNumeric(strTestValue);
			if (m_intError==0)
			{
				try
				{
					long intLong = Convert.ToInt64(strTestValue);
					
				}
				catch (Exception e)
				{
					m_intError=-1;
					m_strError = e.Message;
					MessageBox.Show(e.Message);
				}

			}
		}
		 public void ValidateDecimal(string p_strValue)
		{
			ReturnValue="";
			m_intError=0;
			m_strError="";
			if (IsValueNull(p_strValue)) return;

			//strip out commas
			string strTestValue=p_strValue.Replace(",","");
			if (IsValueNull(strTestValue)) return;

			if (strTestValue.Substring(0,1)=="$") 
			{
				if (strTestValue.Trim().Length == 1)
				{
					if (IsValueNull("")) return;
				}
				else
					strTestValue=strTestValue.Substring(1,strTestValue.Length - 1);
			}
			ValidateNumeric(strTestValue);
			if (m_intError==0)
			{
				try
				{
					decimal intValue = Convert.ToDecimal(strTestValue);
					
					if (TestForMax)
					{
						if ((double)intValue > (double)MaxValue)
						{
							m_intError=-1;
							m_strError="Value exceeds the allowed maximum value of " + MaxValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					if (m_intError==0 && TestForMin==true)
					{
						if ((double)intValue < (double)MinValue)
						{
							m_intError=-1;
							m_strError="Value is less than the allowed minimum value of " + MinValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					
				}
				catch (Exception e)
				{
					m_intError=-1;
					m_strError = e.Message;
					MessageBox.Show(e.Message);
				}

			}
		}
		 public void ValidateSignedByte(string p_strValue)
		{
			ReturnValue="";
			m_intError=0;
			m_strError="";
			if (IsValueNull(p_strValue)) return;

			//strip out commas
			string strTestValue=p_strValue.Replace(",","");
			if (IsValueNull(strTestValue)) return;

			if (strTestValue.Substring(0,1)=="$") 
			{
				if (strTestValue.Trim().Length == 1)
				{
					if (IsValueNull("")) return;
				}
				else
					strTestValue=strTestValue.Substring(1,strTestValue.Length - 1);
			}
			ValidateNumeric(strTestValue);
			if (m_intError==0)
			{
				try
				{
					sbyte intValue = Convert.ToSByte(strTestValue);
					if (TestForMax)
					{
						if (intValue > MaxValue)
						{
							m_intError=-1;
							m_strError="Value exceeds the allowed maximum value of " + MaxValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					if (m_intError==0 && TestForMin==true)
					{
						if (intValue < MinValue)
						{
							m_intError=-1;
							m_strError="Value is less than the allowed minimum value of " + MinValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					
					
				}
				catch (Exception e)
				{
					m_intError=-1;
					m_strError = e.Message;
					MessageBox.Show(e.Message);
				}

			}
		}
		 public void ValidateUnsignedInteger(string p_strValue)
		{
			ReturnValue="";
			m_intError=0;
			m_strError="";
			if (IsValueNull(p_strValue)) return;

			//strip out commas
			string strTestValue=p_strValue.Replace(",","");
			if (IsValueNull(strTestValue)) return;

			if (strTestValue.Substring(0,1)=="$") 
			{
				if (strTestValue.Trim().Length == 1)
				{
					if (IsValueNull("")) return;
				}
				else
					strTestValue=strTestValue.Substring(1,strTestValue.Length - 1);
			}
			ValidateNumeric(strTestValue);
			if (m_intError==0)
			{
				try
				{
					uint intValue = Convert.ToUInt32(strTestValue);
					if (TestForMax)
					{
						if (intValue > MaxValue)
						{
							m_intError=-1;
							m_strError="Value exceeds the allowed maximum value of " + MaxValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					if (m_intError==0 && TestForMin==true)
					{
						if (intValue < MinValue)
						{
							m_intError=-1;
							m_strError="Value is less than the allowed minimum value of " + MinValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					
					
				}
				catch (Exception e)
				{
					m_intError=-1;
					m_strError = e.Message;
					MessageBox.Show(e.Message);
				}

			}
		}
		 public void ValidateUnsignedLong(string p_strValue)
		{
			ReturnValue="";
			m_intError=0;
			m_strError="";
			if (IsValueNull(p_strValue)) return;

			//strip out commas
			string strTestValue=p_strValue.Replace(",","");
			if (IsValueNull(strTestValue)) return;

			if (strTestValue.Substring(0,1)=="$") 
			{
				if (strTestValue.Trim().Length == 1)
				{
					if (IsValueNull("")) return;
				}
				else
					strTestValue=strTestValue.Substring(1,strTestValue.Length - 1);
			}
			ValidateNumeric(strTestValue);
			if (m_intError==0)
			{
				try
				{
					ulong intValue = Convert.ToUInt64(strTestValue);
					if (TestForMax)
					{
						if (intValue > MaxValue)
						{
							m_intError=-1;
							m_strError="Value exceeds the allowed maximum value of " + MaxValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					if (m_intError==0 && TestForMin==true)
					{
						if (intValue < MinValue)
						{
							m_intError=-1;
							m_strError="Value is less than the allowed minimum value of " + MinValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					
					
				}
				catch (Exception e)
				{
					m_intError=-1;
					m_strError = e.Message;
					MessageBox.Show(e.Message);
				}

			}
		}
		 public void ValidateUnsignedShort(string p_strValue)
		{
			ReturnValue="";
			m_intError=0;
			m_strError="";
			if (IsValueNull(p_strValue)) return;

			//strip out commas
			string strTestValue=p_strValue.Replace(",","");
			if (IsValueNull(strTestValue)) return;

			if (strTestValue.Substring(0,1)=="$") 
			{
				if (strTestValue.Trim().Length == 1)
				{
					if (IsValueNull("")) return;
				}
				else
					strTestValue=strTestValue.Substring(1,strTestValue.Length - 1);
			}
			ValidateNumeric(strTestValue);
			if (m_intError==0)
			{
				try
				{
					ushort intValue = Convert.ToUInt16(strTestValue);
					if (TestForMax)
					{
						if (intValue > MaxValue)
						{
							m_intError=-1;
							m_strError="Value exceeds the allowed maximum value of " + MaxValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					if (m_intError==0 && TestForMin==true)
					{
						if (intValue < MinValue)
						{
							m_intError=-1;
							m_strError="Value is less than the allowed minimum value of " + MinValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					
					
				}
				catch (Exception e)
				{
					m_intError=-1;
					m_strError = e.Message;
					MessageBox.Show(e.Message);
				}

			}
		}
		 public void ValidateByte(string p_strValue)
		{
			ReturnValue="";
			m_intError=0;
			m_strError="";
			if (IsValueNull(p_strValue)) return;

			//strip out commas
			string strTestValue=p_strValue.Replace(",","");
			if (IsValueNull(strTestValue)) return;

			if (strTestValue.Substring(0,1)=="$") 
			{
				if (strTestValue.Trim().Length == 1)
				{
					if (IsValueNull("")) return;
				}
				else
					strTestValue=strTestValue.Substring(1,strTestValue.Length - 1);
			}
			ValidateNumeric(strTestValue);
			if (m_intError==0)
			{
				try
				{
					short intValue = Convert.ToByte(strTestValue);
					
					if (TestForMax)
					{
						if (intValue > MaxValue)
						{
							m_intError=-1;
							m_strError="Value exceeds the allowed maximum value of " + MaxValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					if (m_intError==0 && TestForMin==true)
					{
						if (intValue < MinValue)
						{
							m_intError=-1;
							m_strError="Value is less than the allowed minimum value of " + MinValue.ToString();
							MessageBox.Show(m_strError,"FIA Biosum");
						}
					}
					
				}
				catch (Exception e)
				{
					m_intError=-1;
					m_strError = e.Message;
					MessageBox.Show(e.Message);
				}

			}
		}
		 private void ValidateNumeric(string p_strValue)
		{
			m_intError=0;
			m_strError="";
			try
			{
				

				double dbl = Convert.ToDouble(p_strValue);

				Round(p_strValue);


			}
			catch (Exception e)
			{
				m_intError=-1;
				m_strError = e.Message;
				MessageBox.Show(e.Message);
			}
		}
		 private void Round(string p_strValue)
		{
            int x;

			//get the position of decimal period
			int intDecimalPosition = p_strValue.IndexOf(".",0);
            if (intDecimalPosition < 0)
            {
                _strReturnValue = p_strValue;
                if (_bMoney) _strReturnValue = "$" + _strReturnValue + ".00";
                return;
            }
            else if (intDecimalPosition == 0) p_strValue = "0" + p_strValue;
			//get the decimal value
			string strDecimalValue = p_strValue.Substring(intDecimalPosition+1,p_strValue.Length - intDecimalPosition-1);
            //while (strDecimalValue.Length < this._intRoundDecimalLength)
            //    strDecimalValue = strDecimalValue + "0";
            
			
			//check if decimal length is greater than the rounded decimal length
			if (p_strValue.Trim().Length - intDecimalPosition > _intRoundDecimalLength)
			{
				//round the value
				_strReturnValue = Convert.ToString(System.Math.Round(Convert.ToDecimal(p_strValue),_intRoundDecimalLength));
				if (Convert.ToDouble(strDecimalValue)==0 && _strReturnValue.IndexOf(".",0)<0)
				{
					_strReturnValue = _strReturnValue + "." + strDecimalValue.Substring(0,_intRoundDecimalLength);
				}
			}
			else
				_strReturnValue=p_strValue;

            string strWholeValue = "";   
            intDecimalPosition = _strReturnValue.IndexOf(".",0);
            
			strWholeValue = _strReturnValue.Substring(0,_strReturnValue.Length - intDecimalPosition-1);
			if (_bMoney) 
			{  
				
				
				if (intDecimalPosition < 0) 
				{ 
					_strReturnValue="$" + _strReturnValue + ".00";
				}
				else 
				{
					strDecimalValue = _strReturnValue.Substring(intDecimalPosition+1,_strReturnValue.Length - intDecimalPosition-1);
				   
					if (strDecimalValue.Length == 0)
					{
						_strReturnValue=_strReturnValue + "00";
					}
					else if (strDecimalValue.Length==1)
					{
						_strReturnValue=_strReturnValue + "0";

					}
					
					if (strWholeValue.Trim().Length==0)
					{
						_strReturnValue="0" + _strReturnValue;
					}
					_strReturnValue="$" + _strReturnValue;
					
				}
			}
		}
		 private bool IsValueNull(string p_strValue)
		{
			if (p_strValue.Trim().Length ==0) 
			{
				if (NullsAllowed)
				{
				}
				else 
				{
					m_intError=-1;
					m_strError="Enter a value";
					MessageBox.Show(m_strError,"FIA Biosum");
            	}
				return true;
			}
			return false;
		}

		 public int RoundDecimalLength
		{
			set {_intRoundDecimalLength = value;}
		}
		 public string ReturnValue
		{
			get {return _strReturnValue;}
			set {_strReturnValue=value;}
		}
		 public double MaxValue
		{
			get {return _dblMaxValue;}
			set {_dblMaxValue=value;}
		}
		 public double MinValue
		{
			get {return _dblMinValue;}
			set {_dblMinValue=value;}
		}
		 public bool TestForMaxMin
		{
			get {return _bTestForMaxMin;}
			set {TestForMax=value;TestForMin=value;_bTestForMaxMin=value;}
		}
		 public bool TestForMax
		{
			get {return _bTestForMax;}
			set {_bTestForMax=value;}
		}
		 public bool TestForMin
		{
			get {return _bTestForMin;}
			set {_bTestForMin=value;}
		}
		 public bool NullsAllowed
		{
			get {return _bNullsAllowed;}
			set {_bNullsAllowed=value;}
		}
		 public bool Money
		{
			get {return _bMoney;}
			set {_bMoney=value;}
		}


	}
}

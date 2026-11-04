using System;
//using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace FIA_Biosum_Manager
{
    public delegate bool DelegateListViewItemCheckStatus(int row);
    public delegate void DelegateThreadFinished();
    


    public class DelegateTools
    {
        public System.Threading.ManualResetEvent m_oEventStopThread;
        public System.Threading.ManualResetEvent m_oEventThreadStopped;
        public DelegateThreadFinished m_oDelegateThreadFinished;
        public string m_strPropertyValue = "";
        //
        //generic delegate 
        //
        public delegate void SetControlValueCallback(Control oControl, string strPropName, object oPropValue);
        public delegate object GetControlValueCallback(Control oControl, string strPropName,bool bCallback);
        public delegate object GetControlValueWithParentCallback(Control oParentControl,Control oControl, string strPropName, bool bCallback);
        public delegate void ExecuteControlMethodCallback(Control oControl,string strMethodName);
        public delegate void AddControlCallback(Control oParentControl, Control oChildControl);
        public delegate void RemoveControlCallback(Control oParentControl, Control oChildControl);
        public delegate void ExecuteControlMethodWithParamCallback(Control oControl, string strMethodName, object[] oParam);
        public delegate object GetValueExecuteControlMethodWithParamCallback(Control oControl, string strMethodName, object[] oParam, bool bCallback);

        //
        //toolbar delegates
        //
        public delegate object GetToolBarButtonPropertyValueCallback(System.Windows.Forms.ToolBar oToolBar, string strButtonName, string strPropName,bool bCallback);
        public delegate void SetToolBarButtonPropertyValueCallback(System.Windows.Forms.ToolBar oToolBar, string strButtonName, string strPropName, object oPropValue);
        //
        //status bar delegates
        //
        public delegate void SetStatusBarPanelTextValueCallback(System.Windows.Forms.StatusBar oStatusBar, int strStatusBarPanelIndex, string strText);
        public delegate void ExecuteStatusBarPanelMethodCallback(System.Windows.Forms.StatusBar oStatusBar, int strStatusBarPanelIndex, string strMethodName);
        //
        //listview delegates
        //
        public delegate object GetListViewCallback(System.Windows.Forms.ListView oListView, bool bCallback);
        public delegate object GetListViewItemCallback(System.Windows.Forms.ListView oListView, int intItem, bool bCallback);
        public delegate string GetListViewTextValueCallback(System.Windows.Forms.ListView oListView, int intRow, int intCol, bool bCallback);
        public delegate void SetListViewTextValueCallback(System.Windows.Forms.ListView oListView, int intRow, int intCol, string strValue);
        public delegate void SetListViewItemPropertyValueCallback(System.Windows.Forms.ListView oListView, int intItem, string strPropName, object oPropValue);
        public delegate void SetListViewSubItemPropertyValueCallback(System.Windows.Forms.ListView oListView, int intItem, int intSubItem, string strPropName, object oPropValue);
        public delegate void SetListViewColumnsPropertyValueCallback(System.Windows.Forms.ListView oListView, string strPropName, object oPropValue);
        public delegate void SetListViewColumnsItemPropertyValueCallback(System.Windows.Forms.ListView oListView,int intItem,string strPropName, object oPropValue);
        public delegate object GetListViewItemPropertyValueCallback(System.Windows.Forms.ListView oListView, int intItem, string strPropName,bool bCallback);
        public delegate object GetListViewItemsPropertyValueCallback(System.Windows.Forms.ListView oListView, string strPropName, bool bCallback);
        public delegate object GetListViewSubItemPropertyValueCallback(System.Windows.Forms.ListView oListView, int intItem, int intSubItem, string strPropName, bool bCallback);
		public delegate object GetListViewItemSubItemsPropertyValueCallback(System.Windows.Forms.ListViewItem oListViewItem,string strPropName,bool bCallback);
        private delegate int GetListViewCheckedItemsCountCallback(System.Windows.Forms.ListView oListView, bool bCallback);
        public delegate void UpdateListViewItemCallback(System.Windows.Forms.ListView oListView, int intRow, int intCol, string p_strText,
                                                        System.Drawing.Color p_oBackgroundColor,
                                                        System.Drawing.Color p_oForegroundColor);
        public delegate void EnsureListViewItemVisibleCallback(System.Windows.Forms.ListView oListView, int intItem);
        public delegate void ExecuteListViewItemsMethodCallback(System.Windows.Forms.ListView oListView, string strMethodName);
        public delegate void ExecuteListViewColumnsMethodCallback(System.Windows.Forms.ListView oListView, string strMethodName);
        public delegate void ExecuteListViewColumnsMethodWithParamCallback(System.Windows.Forms.ListView oListView, string strMethodName, object[] oParam);
        public delegate void ExecuteListViewItemsMethodWithParamCallback(System.Windows.Forms.ListView oListView, string strMethodName, object[] oParam);
        public delegate void ExecuteListViewSubItemsMethodWithParamCallback(System.Windows.Forms.ListView oListView, int intItem, string strMethodName, object[] oParam);
        public delegate void AddListViewItemCallback(System.Windows.Forms.ListView oListView, string strValue);
        public delegate void AddListViewSubItemCallback(System.Windows.Forms.ListView oListView, int intItem, string strValue);

        //
        //extended listview delegates
        //
        public delegate object GetListViewExItemPropertyValueCallback(ListViewEmbeddedControls.ListViewEx oListViewEx, int intItem, string strPropName, bool bCallback);
        public delegate void SetListViewExItemPropertyValueCallback(ListViewEmbeddedControls.ListViewEx oListViewEx, int intItem, string strPropName, object oPropValue);
        public delegate void EnsureListViewExItemVisibleCallback(ListViewEmbeddedControls.ListViewEx oListViewEx, int intItem);
        public delegate string GetListViewExTextValueCallback(ListViewEmbeddedControls.ListViewEx oListViewEx, int intRow, int intCol, bool bCallback);
        private delegate int GetListViewExCheckedItemsCountCallback(ListViewEmbeddedControls.ListViewEx oListViewEx, bool bCallback);
        public delegate void SetListViewExTextValueCallback(ListViewEmbeddedControls.ListViewEx oListViewEx, int intRow, int intCol, string strValue);
        //
        //combo box delegates
        //
        public delegate object GetComboItemCallback(System.Windows.Forms.ComboBox oCombo, int intItem, bool bCallback);
        public delegate void AddComboItemCallback(System.Windows.Forms.ComboBox oCombo, string strValue);
        public delegate void ExecuteComboItemsMethodCallback(System.Windows.Forms.ComboBox oCombo, string strMethodName);
        public delegate void ExecuteComboItemsMethodWithParamCallback(System.Windows.Forms.ComboBox oCombo, string strMethodName, object[] oParam);
        //
        //form size delegate
        //
        public delegate void SetControlSizeCallback(Control oControl, int intHeight,int intWidth);

            
        public System.Threading.Thread m_oThread;
        private string _strCurrentThreadProcess;
        private bool _bCurrentThreadProcessDone=false;
        private bool _bCurrentThreadProcessStarted = false;
        private bool _bCurrentThreadProcessAborted=false;
        private bool _bCurrentThreadProcessSuspended = false;
		private bool _bCurrentThreadProcessIdle=true;
        public string m_strError = "";
        public int m_intError=0;
        object m_oReturnValue = null;
        int m_intReturnValue = -1;

        public DelegateTools()
        {
            
        }
        public void InitializeThreadEvents()
        {
            this.m_oDelegateThreadFinished = new DelegateThreadFinished(this.ThreadFinished);
            //initialize thread events
            this.m_oEventThreadStopped = new ManualResetEvent(false);
            this.m_oEventStopThread = new ManualResetEvent(false);

        }
        public string GetListViewTextValue(System.Windows.Forms.ListView p_oListView, int p_intRow, int p_intCol,bool p_bCallback)
        {
            if (p_bCallback == false) m_strPropertyValue = "?";

            if (p_oListView.InvokeRequired)
            {
                GetListViewTextValueCallback d = new GetListViewTextValueCallback(GetListViewTextValue);
                p_oListView.Invoke(d, new object[] { p_oListView, p_intRow,  p_intCol, true });

            }
            else
            {
                m_strPropertyValue = p_oListView.Items[p_intRow].SubItems[p_intCol].Text.Trim();
            }

            if (m_strPropertyValue == "?" && p_bCallback == false) m_strPropertyValue = "";
            return m_strPropertyValue;


        }
        public void SetListViewTextValue(System.Windows.Forms.ListView p_oListView, int p_intRow,int p_intCol, string p_strValue)
        {
            if (p_oListView.InvokeRequired)
            {
                SetListViewTextValueCallback d = new SetListViewTextValueCallback(SetListViewTextValue);
                p_oListView.Invoke(d, new object[] { p_oListView, p_intRow, p_intCol,p_strValue});
            }
            else
            {
                p_oListView.Items[p_intRow].SubItems[p_intCol].Text = p_strValue;
            }
        }
        public void SetListViewItemPropertyValue(System.Windows.Forms.ListView p_oListView, int p_intItem, string p_strPropName,object p_oPropValue)
        {
            if (p_oListView.InvokeRequired)
            {
                SetListViewItemPropertyValueCallback d = new SetListViewItemPropertyValueCallback(SetListViewItemPropertyValue);
                p_oListView.Invoke(d, new object[] { p_oListView, p_intItem, p_strPropName, p_oPropValue });
            }
            else
            {
                Type t = p_oListView.Items[p_intItem].GetType();
                System.Reflection.PropertyInfo[] props = t.GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {
                    if (p.Name.Trim().ToUpper() == p_strPropName.Trim().ToUpper())
                    {
                        p.SetValue(p_oListView.Items[p_intItem], p_oPropValue, null);
                        break;
                    }
                }
            }

        }
        public void SetListViewColumnsPropertyValue(System.Windows.Forms.ListView p_oListView, string p_strPropName, object p_oPropValue)
        {
            if (p_oListView.InvokeRequired)
            {
                SetListViewColumnsPropertyValueCallback d = new SetListViewColumnsPropertyValueCallback(SetListViewColumnsPropertyValue);
                p_oListView.Invoke(d, new object[] { p_oListView, p_strPropName, p_oPropValue });
            }
            else
            {
                Type t = p_oListView.Columns.GetType();
                System.Reflection.PropertyInfo[] props = t.GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {
                    if (p.Name.Trim().ToUpper() == p_strPropName.Trim().ToUpper())
                    {
                        p.SetValue(p_oListView.Columns, p_oPropValue, null);
                        break;
                    }
                }
            }

        }
        public void SetListViewColumnsItemPropertyValue(System.Windows.Forms.ListView p_oListView,int p_intItem,string p_strPropName, object p_oPropValue)
        {
            if (p_oListView.InvokeRequired)
            {
                SetListViewColumnsItemPropertyValueCallback d = new SetListViewColumnsItemPropertyValueCallback(SetListViewColumnsItemPropertyValue);
                p_oListView.Invoke(d, new object[] { p_oListView,p_intItem, p_strPropName, p_oPropValue });
            }
            else
            {
                Type t = p_oListView.Columns[p_intItem].GetType();
                System.Reflection.PropertyInfo[] props = t.GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {
                    if (p.Name.Trim().ToUpper() == p_strPropName.Trim().ToUpper())
                    {
                        p.SetValue(p_oListView.Columns[p_intItem], p_oPropValue, null);
                        break;
                    }
                }
            }
        }

        public object GetListView(System.Windows.Forms.ListView p_oListView, bool p_bCallback)
        {
            if (p_bCallback == false) m_oReturnValue = null;

            if (p_oListView.InvokeRequired)
            {
                GetListViewCallback d = new GetListViewCallback(GetListView);
                p_oListView.Invoke(d, new object[] { p_oListView, true });
            }
            else
            {
                m_oReturnValue = p_oListView;
            }
            if (m_oReturnValue == null && p_bCallback == false) m_oReturnValue = false;
            return m_oReturnValue;


        }
        public object GetListViewItem(System.Windows.Forms.ListView p_oListView, int p_intItem, bool p_bCallback)
        {
            if (p_bCallback == false) m_oReturnValue = null;

            if (p_oListView.InvokeRequired)
            {
                GetListViewItemCallback d = new GetListViewItemCallback(GetListViewItem);
                p_oListView.Invoke(d, new object[] { p_oListView, p_intItem, true });
            }
            else
            {
                m_oReturnValue = p_oListView.Items[p_intItem];
            }
            if (m_oReturnValue == null && p_bCallback == false) m_oReturnValue = false;
            return m_oReturnValue;


        }
        public int GetListViewCheckedItemsCount(System.Windows.Forms.ListView p_oListView, bool p_bCallback)
        {
            if (p_bCallback == false) m_intReturnValue = -1;

            if (p_oListView.InvokeRequired)
            {
                GetListViewCheckedItemsCountCallback d = new GetListViewCheckedItemsCountCallback(GetListViewCheckedItemsCount);
                p_oListView.Invoke(d, new object[] { p_oListView, true });
            }
            else
            {
                m_intReturnValue = p_oListView.CheckedItems.Count;
            }
            if (m_intReturnValue == -1 && p_bCallback == false) m_intReturnValue = 0;
            return m_intReturnValue;


        }
        public object GetListViewItemPropertyValue(System.Windows.Forms.ListView p_oListView, int p_intItem, string p_strPropName,bool p_bCallback)
        {
            if (p_bCallback == false) m_oReturnValue = null;

            if (p_oListView.InvokeRequired)
            {
                GetListViewItemPropertyValueCallback d = new GetListViewItemPropertyValueCallback(GetListViewItemPropertyValue);
                p_oListView.Invoke(d, new object[] { p_oListView, p_intItem, p_strPropName,true});
            }
            else
            {
                Type t = p_oListView.Items[p_intItem].GetType();
                System.Reflection.PropertyInfo[] props = t.GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {
                    if (p.Name.Trim().ToUpper() == p_strPropName.Trim().ToUpper())
                    {
                        m_oReturnValue = p.GetValue(p_oListView.Items[p_intItem], null);
                        break;
                    }
                }
            }
            if (m_oReturnValue == null && p_bCallback == false) m_oReturnValue = false;
            return m_oReturnValue;


        }

        public object GetListViewItemsPropertyValue(System.Windows.Forms.ListView p_oListView, string p_strPropName,  bool p_bCallback)
        {
            if (p_bCallback == false) m_oReturnValue = null;

            if (p_oListView.InvokeRequired)
            {
                GetListViewItemsPropertyValueCallback d = new GetListViewItemsPropertyValueCallback(GetListViewItemsPropertyValue);
                p_oListView.Invoke(d, new object[] { p_oListView, p_strPropName, true });
            }
            else
            {
                Type t = p_oListView.Items.GetType();
                System.Reflection.PropertyInfo[] props = t.GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {
                    if (p.Name.Trim().ToUpper() == p_strPropName.Trim().ToUpper())
                    {
                        m_oReturnValue = p.GetValue(p_oListView.Items, null);
                        break;
                    }
                }
            }
            if (m_oReturnValue == null && p_bCallback == false) m_oReturnValue = false;
            return m_oReturnValue;


        }
		public object GetListViewItemSubItemsPropertyValue(System.Windows.Forms.ListViewItem p_oListViewItem, string p_strPropName,  bool p_bCallback)
		{
			if (p_bCallback == false) m_oReturnValue = null;

			if (p_oListViewItem.ListView.InvokeRequired)
			{
				GetListViewItemSubItemsPropertyValueCallback d = new GetListViewItemSubItemsPropertyValueCallback(GetListViewItemSubItemsPropertyValue);
				p_oListViewItem.ListView.Invoke(d, new object[] {p_oListViewItem, p_strPropName, true });
			}
			else
			{
				Type t = p_oListViewItem.SubItems.GetType();
				System.Reflection.PropertyInfo[] props = t.GetProperties();
				foreach (System.Reflection.PropertyInfo p in props)
				{
					if (p.Name.Trim().ToUpper() == p_strPropName.Trim().ToUpper())
					{
						m_oReturnValue = p.GetValue(p_oListViewItem.SubItems, null);
						break;
					}
				}
			}
			if (m_oReturnValue == null && p_bCallback == false) m_oReturnValue = false;
			return m_oReturnValue;


		}

        public object GetListViewSubItemPropertyValue(System.Windows.Forms.ListView p_oListView, int p_intItem, int p_intSubItem, string p_strPropName, bool p_bCallback)
        {
            if (p_bCallback == false) m_oReturnValue = null;

            if (p_oListView.InvokeRequired)
            {
                GetListViewSubItemPropertyValueCallback d = new GetListViewSubItemPropertyValueCallback(GetListViewSubItemPropertyValue);
                p_oListView.Invoke(d, new object[] { p_oListView, p_intItem, p_intSubItem, p_strPropName,true});
            }
            else
            {
                Type t = p_oListView.Items[p_intItem].SubItems[p_intSubItem].GetType();
                System.Reflection.PropertyInfo[] props = t.GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {
                    if (p.Name.Trim().ToUpper() == p_strPropName.Trim().ToUpper())
                    {
                        m_oReturnValue = p.GetValue(p_oListView.Items[p_intItem].SubItems[p_intSubItem], null);
                        break;
                    }
                }
            }
            if (m_oReturnValue == null && p_bCallback == false) m_oReturnValue = false;
            return m_oReturnValue;


        }

        public void SetListViewSubItemPropertyValue(System.Windows.Forms.ListView p_oListView, int p_intItem,int p_intSubItem, string p_strPropName, object p_oPropValue)
        {
            if (p_oListView.InvokeRequired)
            {
                SetListViewSubItemPropertyValueCallback d = new SetListViewSubItemPropertyValueCallback(SetListViewSubItemPropertyValue);
                p_oListView.Invoke(d, new object[] { p_oListView, p_intItem, p_intSubItem,p_strPropName, p_oPropValue });
            }
            else
            {
                Type t = p_oListView.Items[p_intItem].SubItems[p_intSubItem].GetType();
                System.Reflection.PropertyInfo[] props = t.GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {
                    if (p.Name.Trim().ToUpper() == p_strPropName.Trim().ToUpper())
                    {
                        p.SetValue(p_oListView.Items[p_intItem].SubItems[p_intSubItem], p_oPropValue, null);
                        break;
                    }
                }
            }

        }



        public void UpdateListViewItem(
                    System.Windows.Forms.ListView p_oListView,
                    int p_intRow,
                    int p_intCol,
                    string p_strText,
                    System.Drawing.Color p_oBackgroundColor,
                    System.Drawing.Color p_oForegroundColor)
        {
            if (p_oListView.InvokeRequired)
            {
                UpdateListViewItemCallback d = new UpdateListViewItemCallback(UpdateListViewItem);
                p_oListView.Invoke(
                    d,
                    new object[] { p_oListView, p_intRow, p_intCol, p_strText, p_oBackgroundColor, p_oForegroundColor });
            }
            else
            {
                p_oListView.Items[p_intRow].Selected = true;
                System.Windows.Forms.ListViewItem.ListViewSubItem oSubItem =
                    p_oListView.Items[p_intRow].SubItems[p_intCol];
                    oSubItem.BackColor = p_oBackgroundColor;
                    oSubItem.ForeColor = p_oForegroundColor;
                p_oListView.Items[p_intRow].SubItems[p_intCol].Text = p_strText;
            }

        }
        public void EnsureListViewItemVisible(
            System.Windows.Forms.ListView p_oListView,
            int p_intItem)
        {
            if (p_oListView.InvokeRequired)
            {
                EnsureListViewItemVisibleCallback d = new EnsureListViewItemVisibleCallback(EnsureListViewItemVisible);
                p_oListView.Invoke(
                    d,
                    new object[] { p_oListView, p_intItem});
            }
            else
            {
                p_oListView.EnsureVisible(p_intItem);
            }

        }
        public void ExecuteListViewItemsMethod(System.Windows.Forms.ListView p_oListView, string p_strMethodName)
        {
            if (p_oListView.InvokeRequired)
            {
                ExecuteListViewItemsMethodCallback d = new ExecuteListViewItemsMethodCallback(ExecuteListViewItemsMethod);
                p_oListView.Invoke(d, new object[] { p_oListView, p_strMethodName });
            }
            else
            {
                Type t = p_oListView.Items.GetType();
                System.Reflection.MethodInfo[] methods = t.GetMethods();
                foreach (System.Reflection.MethodInfo m in methods)
                {
                    if (m.Name.Trim().ToUpper() == p_strMethodName.Trim().ToUpper())
                    {
                        m.Invoke(p_oListView.Items, null);
                        break;
                    }
                }
            }

        }
        public void ExecuteListViewItemsMethodWithParam(System.Windows.Forms.ListView p_oListView, string p_strMethodName,object[] p_oParam)
        {
            if (p_oListView.InvokeRequired)
            {
                ExecuteListViewItemsMethodWithParamCallback d = new ExecuteListViewItemsMethodWithParamCallback(ExecuteListViewItemsMethodWithParam);
                p_oListView.Invoke(d, new object[] { p_oListView, p_strMethodName,p_oParam });
            }
            else
            {
                Type t = p_oListView.Items.GetType();
                System.Reflection.MethodInfo[] methods = t.GetMethods();
                foreach (System.Reflection.MethodInfo m in methods)
                {
                    if (m.Name.Trim().ToUpper() == p_strMethodName.Trim().ToUpper())
                    {
                        m.Invoke(p_oListView.Items, p_oParam);
                        break;
                    }
                }
            }

        }
        public void ExecuteListViewSubItemsMethodWithParam(System.Windows.Forms.ListView p_oListView, int p_intItem, string p_strMethodName, object[] p_oParam)
        {
            if (p_oListView.InvokeRequired)
            {
                ExecuteListViewSubItemsMethodWithParamCallback d = new ExecuteListViewSubItemsMethodWithParamCallback(ExecuteListViewSubItemsMethodWithParam);
                p_oListView.Invoke(d, new object[] { p_oListView, p_intItem, p_strMethodName, p_oParam });
            }
            else
            {
                Type t = p_oListView.Items[p_intItem].SubItems.GetType();
                System.Reflection.MethodInfo[] methods = t.GetMethods();
                foreach (System.Reflection.MethodInfo m in methods)
                {
                    if (m.Name.Trim().ToUpper() == p_strMethodName.Trim().ToUpper())
                    {
                        m.Invoke(p_oListView.Items[p_intItem].SubItems, p_oParam);
                        break;
                    }
                }
            }

        }


        public void ExecuteListViewColumnsMethod(System.Windows.Forms.ListView p_oListView, string p_strMethodName)
        {
            if (p_oListView.InvokeRequired)
            {
                ExecuteListViewColumnsMethodCallback d = new ExecuteListViewColumnsMethodCallback(ExecuteListViewItemsMethod);
                p_oListView.Invoke(d, new object[] { p_oListView, p_strMethodName });
            }
            else
            {
                Type t = p_oListView.Items.GetType();
                System.Reflection.MethodInfo[] methods = t.GetMethods();
                foreach (System.Reflection.MethodInfo m in methods)
                {
                    if (m.Name.Trim().ToUpper() == p_strMethodName.Trim().ToUpper())
                    {
                        m.Invoke(p_oListView.Columns, null);
                        break;
                    }
                }
            }

        }
        public void ExecuteListViewColumnsMethodWithParam(System.Windows.Forms.ListView p_oListView, string p_strMethodName,object[] p_oParam)
        {
            if (p_oListView.InvokeRequired)
            {
                ExecuteListViewColumnsMethodWithParamCallback d = new ExecuteListViewColumnsMethodWithParamCallback(ExecuteListViewColumnsMethodWithParam);
                p_oListView.Invoke(d, new object[] { p_oListView, p_strMethodName,p_oParam });
            }
            else
            {
                Type t = p_oListView.Columns.GetType();
                System.Reflection.MethodInfo[] methods = t.GetMethods();
                foreach (System.Reflection.MethodInfo m in methods)
                {
                    if (m.Name.Trim().ToUpper() == p_strMethodName.Trim().ToUpper())
                    {
                        m.Invoke(p_oListView.Columns, p_oParam);
                        break;
                    }
                }
            }

        }
        public void AddListViewItem(System.Windows.Forms.ListView p_oListView, string p_strValue)
        {
            if (p_oListView.InvokeRequired)
            {
                AddListViewItemCallback d = new AddListViewItemCallback(AddListViewItem);
                p_oListView.Invoke(
                    d,
                    new object[] { p_oListView, p_strValue });
            }
            else
            {
                p_oListView.Items.Add(p_strValue);
            }
        }
        public void AddListViewSubItem(System.Windows.Forms.ListView p_oListView,int p_intItem, string p_strValue)
        {
            if (p_oListView.InvokeRequired)
            {
                AddListViewSubItemCallback d = new AddListViewSubItemCallback(AddListViewSubItem);
                p_oListView.Invoke(
                    d,
                    new object[] { p_oListView,p_intItem,p_strValue });
            }
            else
            {
                p_oListView.Items[p_intItem].SubItems.Add(p_strValue);
            }
        }



        public string GetListViewExTextValue(ListViewEmbeddedControls.ListViewEx p_oListViewEx, int p_intRow, int p_intCol, bool p_bCallback)
        {
            if (p_bCallback == false) m_strPropertyValue = "?";

            if (p_oListViewEx.InvokeRequired)
            {
                GetListViewExTextValueCallback d = new GetListViewExTextValueCallback(GetListViewExTextValue);
                p_oListViewEx.Invoke(d, new object[] { p_oListViewEx, p_intRow, p_intCol, true });

            }
            else
            {
                m_strPropertyValue = p_oListViewEx.Items[p_intRow].SubItems[p_intCol].Text.Trim();
            }

            if (m_strPropertyValue == "?" && p_bCallback == false) m_strPropertyValue = "";
            return m_strPropertyValue;


        }

        public object GetListViewExItemPropertyValue(ListViewEmbeddedControls.ListViewEx p_oListViewEx, int p_intItem, string p_strPropName, bool p_bCallback)
        {
            if (p_bCallback == false) m_oReturnValue = null;

            if (p_oListViewEx.InvokeRequired)
            {
                GetListViewExItemPropertyValueCallback d = new GetListViewExItemPropertyValueCallback(GetListViewExItemPropertyValue);
                p_oListViewEx.Invoke(d, new object[] { p_oListViewEx, p_intItem, p_strPropName,true });
            }
            else
            {
                Type t = p_oListViewEx.Items[p_intItem].GetType();
                System.Reflection.PropertyInfo[] props = t.GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {
                    if (p.Name.Trim().ToUpper() == p_strPropName.Trim().ToUpper())
                    {
                        m_oReturnValue = p.GetValue(p_oListViewEx.Items[p_intItem], null);
                        break;
                    }
                }
            }
            if (m_oReturnValue == null && p_bCallback == false) m_oReturnValue = false;
            return m_oReturnValue;


        }


        public void SetListViewExItemPropertyValue(ListViewEmbeddedControls.ListViewEx p_oListViewEx, int p_intItem, string p_strPropName, object p_oPropValue)
        {
            if (p_oListViewEx.InvokeRequired)
            {
                SetListViewExItemPropertyValueCallback d = new SetListViewExItemPropertyValueCallback(SetListViewExItemPropertyValue);
                p_oListViewEx.Invoke(d, new object[] { p_oListViewEx, p_intItem, p_strPropName, p_oPropValue });
            }
            else
            {
                Type t = p_oListViewEx.Items[p_intItem].GetType();
                System.Reflection.PropertyInfo[] props = t.GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {
                    if (p.Name.Trim().ToUpper() == p_strPropName.Trim().ToUpper())
                    {
                        p.SetValue(p_oListViewEx.Items[p_intItem], p_oPropValue, null);
                    }
                }
            }

        }
        public void SetListViewExTextValue(ListViewEmbeddedControls.ListViewEx p_oListViewEx, int p_intRow, int p_intCol, string p_strValue)
        {
            if (p_oListViewEx.InvokeRequired)
            {
                SetListViewExTextValueCallback d = new SetListViewExTextValueCallback(SetListViewExTextValue);
                p_oListViewEx.Invoke(d, new object[] { p_oListViewEx, p_intRow, p_intCol, p_strValue });
            }
            else
            {
                p_oListViewEx.Items[p_intRow].SubItems[p_intCol].Text = p_strValue;
            }
        }


        public void EnsureListViewExItemVisible(
            ListViewEmbeddedControls.ListViewEx p_oListViewEx, int p_intItem)
        {
            if (p_oListViewEx.InvokeRequired)
            {
                EnsureListViewExItemVisibleCallback d = new EnsureListViewExItemVisibleCallback(EnsureListViewExItemVisible);
                p_oListViewEx.Invoke(
                    d,
                    new object[] { p_oListViewEx, p_intItem });
            }
            else
            {
                p_oListViewEx.EnsureVisible(p_intItem);
            }
        }
        public int GetListViewExCheckedItemsCount(ListViewEmbeddedControls.ListViewEx p_oListViewEx, bool p_bCallback)
        {
            if (p_bCallback == false) m_intReturnValue = -1;

            if (p_oListViewEx.InvokeRequired)
            {
                GetListViewExCheckedItemsCountCallback d = new GetListViewExCheckedItemsCountCallback(GetListViewExCheckedItemsCount);
                p_oListViewEx.Invoke(d, new object[] { p_oListViewEx, true });
            }
            else
            {
                m_intReturnValue = p_oListViewEx.CheckedItems.Count;
            }
            if (m_intReturnValue == -1 && p_bCallback == false) m_intReturnValue = 0;
            return m_intReturnValue;


        }

        public object GetComboItem(System.Windows.Forms.ComboBox p_oCombo, int p_intItem, bool p_bCallback)
        {
            if (p_bCallback == false) m_oReturnValue = null;

            if (p_oCombo.InvokeRequired)
            {
                GetComboItemCallback d = new GetComboItemCallback(GetComboItem);
                p_oCombo.Invoke(d, new object[] { p_oCombo, p_intItem, true });
            }
            else
            {
                m_oReturnValue = p_oCombo.Items[p_intItem];
            }
            if (m_oReturnValue == null && p_bCallback == false) m_oReturnValue = "";
            return m_oReturnValue;
        }
        public void ExecuteComboItemsMethod(System.Windows.Forms.ComboBox p_oCombo, string p_strMethodName)
        {
            if (p_oCombo.InvokeRequired)
            {
                ExecuteComboItemsMethodCallback d = new ExecuteComboItemsMethodCallback(ExecuteComboItemsMethod);
                p_oCombo.Invoke(d, new object[] { p_oCombo, p_strMethodName });
            }
            else
            {
                Type t = p_oCombo.Items.GetType();
                System.Reflection.MethodInfo[] methods = t.GetMethods();
                foreach (System.Reflection.MethodInfo m in methods)
                {
                    if (m.Name.Trim().ToUpper() == p_strMethodName.Trim().ToUpper())
                    {
                        m.Invoke(p_oCombo.Items, null);
                        break;
                    }
                }
            }

        }
        public void ExecuteComboItemsMethodWithParam(System.Windows.Forms.ComboBox p_oCombo, string p_strMethodName, object[] p_oParam)
        {
            if (p_oCombo.InvokeRequired)
            {
                ExecuteComboItemsMethodWithParamCallback d = new ExecuteComboItemsMethodWithParamCallback(ExecuteComboItemsMethodWithParam);
                p_oCombo.Invoke(d, new object[] { p_oCombo, p_strMethodName, p_oParam });
            }
            else
            {
                Type t = p_oCombo.Items.GetType();
                System.Reflection.MethodInfo[] methods = t.GetMethods();
                foreach (System.Reflection.MethodInfo m in methods)
                {
                    if (m.Name.Trim().ToUpper() == p_strMethodName.Trim().ToUpper())
                    {
                        m.Invoke(p_oCombo.Items, p_oParam);
                        break;
                    }
                }
            }

        }


        public void AddComboItem(System.Windows.Forms.ComboBox p_oCombo, string p_strValue)
        {
            if (p_oCombo.InvokeRequired)
            {
                AddComboItemCallback d = new AddComboItemCallback(AddComboItem);
                p_oCombo.Invoke(
                    d,
                    new object[] { p_oCombo, p_strValue });
            }
            else
            {
                p_oCombo.Items.Add(p_strValue);
            }
        }



		public void SetStatusBarPanelTextValue(System.Windows.Forms.StatusBar p_oStatusBar,int p_intStatusBarPanelIndex, string p_strValue)
		{
			if (p_oStatusBar.InvokeRequired)
			{
				SetStatusBarPanelTextValueCallback d = new SetStatusBarPanelTextValueCallback(this.SetStatusBarPanelTextValue);
				p_oStatusBar.Invoke(d, new object[] { p_oStatusBar, p_intStatusBarPanelIndex, p_strValue });
			}
			else
			{
				p_oStatusBar.Panels[p_intStatusBarPanelIndex].Text = p_strValue;
			}
		}
		public void ExecuteStatusBarPanelMethod(System.Windows.Forms.StatusBar p_oStatusBar, int p_intStatusBarPanelIndex, string p_strMethodName)
		{
			if (p_oStatusBar.InvokeRequired)
			{
				ExecuteStatusBarPanelMethodCallback d = new ExecuteStatusBarPanelMethodCallback(ExecuteStatusBarPanelMethod);
				p_oStatusBar.Invoke(d, new object[] { p_oStatusBar, p_intStatusBarPanelIndex,p_strMethodName });
			}
			else
			{
			    Type t = p_oStatusBar.Panels[p_intStatusBarPanelIndex].GetType();
				System.Reflection.MethodInfo[] methods = t.GetMethods();
				foreach (System.Reflection.MethodInfo m in methods)
				{
					if (m.Name.Trim().ToUpper() == p_strMethodName.Trim().ToUpper())
					{
						m.Invoke(p_oStatusBar.Panels[p_intStatusBarPanelIndex],null);
						return;
					}
				}
			}
		}

		
        public void SetControlPropertyValue(Control p_oControl, string p_strPropName, object p_oPropValue)
        {
            if (p_oControl.InvokeRequired)
            {
                SetControlValueCallback d = new SetControlValueCallback(SetControlPropertyValue);
                p_oControl.Invoke(d, new object[] { p_oControl, p_strPropName, p_oPropValue });
            }
            else
            {
                Type t = p_oControl.GetType();
                System.Reflection.PropertyInfo[] props = t.GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {
                    if (p.Name.Trim().ToUpper() == p_strPropName.Trim().ToUpper())
                    {
                        p.SetValue(p_oControl, p_oPropValue, null);
                    }
                }
            }

        }

        public object GetControlPropertyValue(Control p_oControl, string p_strPropName,bool p_bCallback)
        {
            if (p_bCallback == false) m_oReturnValue = null;
            if (p_oControl.InvokeRequired)
            {
                GetControlValueCallback d = new GetControlValueCallback(GetControlPropertyValue);
                p_oControl.Invoke(d, new object[] { p_oControl, p_strPropName,true });
            }
            else
            {
                Type t = p_oControl.GetType();
                System.Reflection.PropertyInfo[] props = t.GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {
                    if (p.Name.Trim().ToUpper() == p_strPropName.Trim().ToUpper())
                    {
                        m_oReturnValue = p.GetValue(p_oControl, null);
                        break;
                    }
                }
            }
            if (m_oReturnValue == null && p_bCallback == false) m_oReturnValue = false;
            return m_oReturnValue;

        }
        public object GetControlPropertyValue(Control p_oParentControl,Control p_oControl, string p_strPropName, bool p_bCallback)
        {
            if (p_bCallback == false) m_oReturnValue = null;
            if (p_oParentControl.InvokeRequired)
            {
                GetControlValueWithParentCallback d = new GetControlValueWithParentCallback(GetControlPropertyValue);
                p_oParentControl.Invoke(d, new object[] {p_oParentControl, p_oControl, p_strPropName, true });
            }
            else
            {
                Type t = p_oControl.GetType();
                System.Reflection.PropertyInfo[] props = t.GetProperties();
                foreach (System.Reflection.PropertyInfo p in props)
                {
                    if (p.Name.Trim().ToUpper() == p_strPropName.Trim().ToUpper())
                    {
                        m_oReturnValue = p.GetValue(p_oControl, null);
                        break;
                    }
                }
            }
            if (m_oReturnValue == null && p_bCallback == false) m_oReturnValue = false;
            return m_oReturnValue;

        }
        public void ExecuteControlMethod(Control p_oControl,string p_strMethodName)
        {
            
                if (p_oControl.InvokeRequired)
                {
                    ExecuteControlMethodCallback d = new ExecuteControlMethodCallback(ExecuteControlMethod);
                    try
                    {
                        p_oControl.Invoke(d, new object[] { p_oControl, p_strMethodName });
                    }
                    catch
                    {
                    }
                }
                else
                {
                    Type t = p_oControl.GetType();
                    System.Reflection.MethodInfo[] methods = t.GetMethods();
                    foreach (System.Reflection.MethodInfo m in methods)
                    {
                        if (m.Name.Trim().ToUpper() == p_strMethodName.Trim().ToUpper())
                        {
                            m.Invoke(p_oControl, null);
                            break;
                        }
                    }
                }
            

        }
        public void ExecuteControlMethodWithParam(Control p_oControl, string p_strMethodName, object[] p_oParam)
        {
            if (p_oControl.InvokeRequired)
            {
                ExecuteControlMethodWithParamCallback d = new ExecuteControlMethodWithParamCallback(ExecuteControlMethodWithParam);
                p_oControl.Invoke(d, new object[] { p_oControl, p_strMethodName, p_oParam });
            }
            else
            {
                Type t = p_oControl.GetType();
                System.Reflection.MethodInfo[] methods = t.GetMethods();
                foreach (System.Reflection.MethodInfo m in methods)
                {
                    if (m.Name.Trim().ToUpper() == p_strMethodName.Trim().ToUpper())
                    {
                        m.Invoke(p_oControl, p_oParam);
                        break;
                    }
                }
            }

        }

        public object GetValueExecuteControlMethodWithParam(Control p_oControl, string p_strMethodName, object[] p_oParam, 
            bool p_bCallback)
        {
            if (p_bCallback == false) m_oReturnValue = null;
            if (p_oControl.InvokeRequired)
            {
                GetValueExecuteControlMethodWithParamCallback d = new GetValueExecuteControlMethodWithParamCallback(GetValueExecuteControlMethodWithParam);
                p_oControl.Invoke(d, new object[] { p_oControl, p_strMethodName, p_oParam, true });
            }
            else
            {
                Type t = p_oControl.GetType();
                System.Reflection.MethodInfo[] methods = t.GetMethods();
                foreach (System.Reflection.MethodInfo m in methods)
                {
                    if (m.Name.Trim().ToUpper() == p_strMethodName.Trim().ToUpper())
                    {
                        m_oReturnValue = m.Invoke(p_oControl, p_oParam);
                        break;
                    }
                }
            }
            if (m_oReturnValue == null && p_bCallback == false) m_oReturnValue = false;
            return m_oReturnValue;
        }

        public void SetControlSize(Control p_oControl, int p_intHeight,int p_intWidth)
        {
            if (p_oControl.InvokeRequired)
            {
                SetControlSizeCallback d = new SetControlSizeCallback(SetControlSize);
                p_oControl.Invoke(d, new object[] { p_oControl,p_intHeight,p_intWidth});
            }
            else
            {
                p_oControl.Size = new System.Drawing.Size(p_intWidth, p_intHeight);
            }

        }


        public void AddControl(Control p_oParentControl, Control p_oChildControl)
        {
            if (p_oParentControl.InvokeRequired)
            {
                AddControlCallback d = new AddControlCallback(AddControl);
                p_oParentControl.Invoke(
                    d,
                    new object[] { p_oParentControl, p_oChildControl });
            }
            else
            {
                p_oParentControl.Controls.Add(p_oChildControl);
            }
        }
        public void RemoveControl(Control p_oParentControl, Control p_oChildControl)
        {
            if (p_oParentControl.InvokeRequired)
            {
                RemoveControlCallback d = new RemoveControlCallback(RemoveControl);
                p_oParentControl.Invoke(
                    d,
                    new object[] { p_oParentControl, p_oChildControl });
            }
            else
            {
                p_oParentControl.Controls.Remove(p_oChildControl);
            }
        }





        public void StopThread()
        {
            if (this.m_oThread != null && m_oThread.IsAlive)
            {
                //set event stop
                this.m_oEventStopThread.Set();

                //wait until thread finishes or stops
                while (m_oThread.IsAlive)
                {
                    if (System.Threading.WaitHandle.WaitAll((new ManualResetEvent[] { this.m_oEventThreadStopped }),
                        100, true))
                    {
                        break;
                    }
                    Application.DoEvents();
                }
            }
            this.CurrentThreadProcessDone = true;
			this.CurrentThreadProcessIdle=true;
            MessageBox.Show("Thread " + this.m_oThread.Name + " Aborted");
        }
        public bool AbortProcessing(string p_strTitle, string p_strQuestion)
        {
            try
            {
                string strMsg = p_strQuestion;
                DialogResult result = MessageBox.Show(strMsg, p_strTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                switch (result)
                {
                    case DialogResult.Yes:
                        if (this.CurrentThreadProcessSuspended)
                        {
                            m_oThread.Resume();
                            this.CurrentThreadProcessSuspended = false;
                        }
                        this.m_oThread.Abort();
                        this.CurrentThreadProcessAborted = true;
						this.CurrentThreadProcessIdle=true;
                        return true;
                }
            }
            catch (Exception err)
            {
                
            }
            return false;

        }
        public void ShowMessageBox(System.Windows.Forms.UserControl p_oUserControl, string p_strMessage, string p_strCaption)
        {
            if (p_oUserControl.InvokeRequired)
            {
                p_oUserControl.Invoke(new Action(() =>
                {
                    MessageBox.Show(p_oUserControl, p_strMessage, p_strCaption);
                }));
            }
            else
            {
                MessageBox.Show(p_oUserControl, p_strMessage, p_strCaption);
            }
        }
        public DialogResult ShowMessageBoxWithResult (System.Windows.Forms.UserControl p_oUserControl, string p_strMessage, string p_strCaption)
        {
            DialogResult res = new DialogResult();
            if (p_oUserControl.InvokeRequired)
            {
                p_oUserControl.Invoke(new Action(() =>
                {
                    res = MessageBox.Show(p_oUserControl, p_strMessage, p_strCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                }));
            }
            else
            {
                res = MessageBox.Show(p_oUserControl, p_strMessage, p_strCaption, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }
            return res;
        }

        private void ThreadFinished()
        {
            this.CurrentThreadProcessDone = true;
			this.CurrentThreadProcessIdle=true;
            //MessageBox.Show("DelegateTools.ThreadFinished: Thread " + this.m_oThread.Name + " Finished");
        }
        public string CurrentThreadProcessName
        {
            set { _strCurrentThreadProcess = value; }
            get { return _strCurrentThreadProcess; }
        }
        public bool CurrentThreadProcessDone
        {
            set { _bCurrentThreadProcessDone = value; }
            get { return _bCurrentThreadProcessDone; }
        }
        public bool CurrentThreadProcessAborted
        {
            set { _bCurrentThreadProcessAborted = value; }
            get { return _bCurrentThreadProcessAborted; }
        }
        public bool CurrentThreadProcessStarted
        {
            set { _bCurrentThreadProcessStarted = value; }
            get { return _bCurrentThreadProcessStarted; }
        }
        public bool CurrentThreadProcessSuspended
        {
            set { _bCurrentThreadProcessSuspended = value; }
            get { return _bCurrentThreadProcessSuspended; }
        }
		public bool CurrentThreadProcessIdle
		{
			set { _bCurrentThreadProcessIdle = value; }
			get { return _bCurrentThreadProcessIdle; }
		}


    }

}

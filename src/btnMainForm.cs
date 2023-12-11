using System;
using System.Windows.Forms;
using System.Drawing;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for btnMainForm.
	/// </summary>
	public class btnMainForm:System.Windows.Forms.Button
	{
		private FIA_Biosum_Manager.frmMain m_frmMain;
		private System.Drawing.Color _DisabledColor = System.Drawing.Color.Red;
		private System.Windows.Forms.ToolTip m_tooltip;
		public btnMainForm(FIA_Biosum_Manager.frmMain p_frmMain)
		{
			this.m_frmMain =p_frmMain;
			//
			// TODO: Add constructor logic here
			//
			this.BackColor = System.Drawing.Color.FromArgb(169,169,169);
			this.Click += new System.EventHandler(this.Custom_Click);
			this.MouseEnter += new System.EventHandler(this.Custom_MouseEnter);
			this.MouseLeave += new System.EventHandler(this.Custom_MouseLeave);
			this.m_tooltip = new ToolTip();
			this.m_tooltip.ShowAlways = true;
			this.m_tooltip.AutoPopDelay = 30000;
			

			
		}
		private void Custom_Click(object sender, System.EventArgs e)
		{
			//MessageBox.Show(this.Text);
			this.m_frmMain.button_click(this.Text);
		}
		private void Custom_MouseEnter(object sender, System.EventArgs e)
		{
         	this.BackColor = System.Drawing.Color.FromArgb(200,200,200);
		}
		private void Custom_MouseLeave(object sender, System.EventArgs e)
		{
			this.BackColor = System.Drawing.Color.FromArgb(169,169,169);
		}

		public System.Drawing.Color DisabledColor
		{
            get
			{
                 return _DisabledColor;
			}
			set 
			{
                 _DisabledColor = value;
			}
			  
		}
		public FIA_Biosum_Manager.frmMain p_frmMain
		{
			set 
			{
				this.m_frmMain = value;
			}
		}
		public string strToolTip
		{
			set { this.m_tooltip.SetToolTip(this,value);}
		}




	}
	public class btnMainFormCollection:System.Collections.CollectionBase
	{
		public void Add(btnMainForm btnMainForm1)
		{
			// vérify if object is not already in
			if (this.List.Contains(btnMainForm1))
				throw new InvalidOperationException();
 
			// adding it
			this.List.Add(btnMainForm1);
 
			// return collection
			//return this;
		}
		public void Remove(int index)
		{
			// Check to see if there is a widget at the supplied index.
			if (index > Count - 1 || index < 0)
				// If no widget exists, a messagebox is shown and the operation 
				// is cancelled.
			{
				System.Windows.Forms.MessageBox.Show("Index not valid!");
			}
			else
			{
				List.RemoveAt(index); 
			}
		}
		public btnMainForm Item(int Index)
		{
			// The appropriate item is retrieved from the List object and
			// explicitly cast to the Widget type, then returned to the 
			// caller.
			return (btnMainForm) List[Index];
		}

	}
}

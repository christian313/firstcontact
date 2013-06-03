/*
 * Created by SharpDevelop.
 * User: christian
 * Date: 03.06.2013
 * Time: 17:36
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.InteropServices;
using Outlook = NetOffice.OutlookApi;
using Microsoft.Vbe.Interop.Forms;
using System.Windows.Forms
	
namespace FirstContact
{
	/// <summary>
	/// Description of Class1.
	/// </summary>
    [ComVisible(true),
    Guid("807FB124-45D5-4E47-94D1-979FB4FB85F5"),
    ProgId("FirstContact.FormRegionHookup"),
    ClassInterface(ClassInterfaceType.AutoDual)]
    class FormRegionHookup : Outlook.FormRegionStartu
	{
		public FormRegionHookup()
		{
		}
		
        public object GetFormRegionStorage(string FormRegionName, object Item, int LCID,
        Outlook.OlFormRegionMode FormRegionMode, Outlook.OlFormRegionSize FormRegionSize)
        {
            Application.DoEvents();
            System.Diagnostics.Debug.Write("GetFormRegionStorage, FormRegionName: " + FormRegionName);
            switch (FormRegionName)
            {
                case "FormRegionsVS":
                    System.Diagnostics.Debug.Write("case: " + FormRegionName);
                    byte[] ofsBytes = Properties.Resources.firstcontact;
                    return ofsBytes;
                default:
                    return null;
            }
        }

        public void BeforeFormRegionShow(Outlook.FormRegion FormRegion)
        {
            this.FormRegion = FormRegion;
            this.UserForm = FormRegion.Form as UserForm;

            System.Diagnostics.Debug.Write("BeforeFormRegionShow");

            try
            {
                //System.Diagnostics.Debug.Write("BeforeFormRegionShow 1");
                CommandButton1 = UserForm.Controls.Item("CommandButton1") as Outlook.OlkCommandButton;
                //System.Diagnostics.Debug.Write("BeforeFormRegionShow 2");
                CommandButton1.Click += new Outlook.OlkCommandButtonEvents_ClickEventHandler(CommandButton1_Click);
                CommandButton2 = UserForm.Controls.Item("CommandButton2") as Outlook.OlkCommandButton;
                CommandButton2.Click += new Outlook.OlkCommandButtonEvents_ClickEventHandler(CommandButton2_Click);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public object GetFormRegionManifest(string formRegionName, int LCID)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public object GetFormRegionIcon(string formRegionName, int LCID, Outlook.OlFormRegionIcon icon)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        private void CommandButton1_Click()
        {
            System.Diagnostics.Debug.Write("CommandButton1_Click has been called from Button 1");
        }

        private void CommandButton2_Click()
        {
            System.Diagnostics.Debug.Write("CommandButton1_Click has been called from Button 2");
        		
	}
}

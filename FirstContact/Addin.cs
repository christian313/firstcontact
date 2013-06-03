using System;
using System.Reflection;
using System.Windows.Forms; 
using Microsoft.Win32;
using System.Runtime.InteropServices;
using Extensibility;
using NetOffice;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;

namespace FirstContact
{
    [GuidAttribute("807FB124-45D5-4E47-94D1-979FB4FB85F5"), ProgId("FirstContact.Addin"), ComVisible(true)]
    public class Addin :IDTExtensibility2
    {
		Outlook.Application _application;

        public Addin()
        {
        }
        
        #region IDTExtensibility2 Members

 		void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
			CreateUserInterface();            
        }

        void IDTExtensibility2.OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
 			_application = new Outlook.Application(null, Application);          
        }

        void IDTExtensibility2.OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {

			RemoveUserInterface();

			if(null != _application)
			{
				_application.Dispose();
				_application = null;
			}
        }

        void IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        {
           
        }

        void IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {
             
        }

        #endregion
		
		#region Classic UI

		private void CreateUserInterface()
		{
			// TODO: create UI items
		}

		private void RemoveUserInterface()
		{
			
		}

		#endregion

        #region COM Register Functions

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            try
            {
                // add codebase value
                Assembly thisAssembly = Assembly.GetAssembly(typeof(Addin));
                RegistryKey key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32\\1.0.0.0");
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();

                key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32");
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();

                // add bypass key
                // http://support.microsoft.com/kb/948461
                key = Registry.ClassesRoot.CreateSubKey("Interface\\{000C0601-0000-0000-C000-000000000046}");
                string defaultValue = key.GetValue("") as string;
                if (null == defaultValue)
                    key.SetValue("", "Office .NET Framework Lockback Bypass Key");
                key.Close();
                
				// register addin in Outlook
				Registry.LocalMachine.CreateSubKey(@"Software\Microsoft\Office\Outlook\Addins\FirstContact.Addin");
				RegistryKey regKeyOutlook = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Office\Outlook\Addins\FirstContact.Addin", true);
				regKeyOutlook.SetValue("LoadBehavior", Convert.ToInt32(3));
				regKeyOutlook.SetValue("FriendlyName", "FirstContact");
				regKeyOutlook.SetValue("Description", "add FormRegion to Outlook Contact");
				regKeyOutlook.Close();

            }
            catch (Exception ex)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Register Addin", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            try
            { 
                // unregister addin
                Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable", false);
                
                // unregister addin in office
				Registry.LocalMachine.DeleteSubKey(@"Software\Microsoft\Office\Outlook\Addins\FirstContact.Addin", false);

            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Unregister Addin", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

    }
}

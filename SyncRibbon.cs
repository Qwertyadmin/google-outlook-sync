using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;


namespace googleSync
{
	[ComVisible(true)]
	public class SyncRibbon : Office.IRibbonExtensibility
	{
		private Office.IRibbonUI ribbon;

		public SyncRibbon()
		{
		}

		#region Membri IRibbonExtensibility

		public string GetCustomUI(string ribbonID)
		{
			return GetResourceText("googleSync.SyncRibbon.xml");
		}

		#endregion

		#region Callback della barra multifunzione
		//Creare qui i metodi callback. Per altre informazioni sull'aggiunta di metodi callback, vedere https://go.microsoft.com/fwlink/?LinkID=271226

		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			this.ribbon = ribbonUI;
		}

        public void CreateOEventButton(Office.IRibbonControl control)
        {
            SyncClass syncClass = new SyncClass();
            try
            {
                syncClass.CreateOEvent(ThisAddIn.calendarService, ThisAddIn.oStore, ThisAddIn.calendars, "primary");
            }
            catch (Exception error)
            {
                syncClass.WriteLog("ERROR: " + error);
            }
        }

		public void CreateOEventMenuButton(Office.IRibbonControl control)
		{
			SyncForm syncForm = new SyncForm();
			syncForm.Show();
		}

		public void DeleteOEventButton(Office.IRibbonControl control)
        {
            SyncClass syncClass = new SyncClass();
            try
            {
                syncClass.DeleteOEvent(ThisAddIn.calendarService, ThisAddIn.oStore, ThisAddIn.calendars);
            }
            catch (Exception error)
            {
                syncClass.WriteLog("ERROR: " + error);
            }
        }

        public void CreateOContactButton(Office.IRibbonControl control)
		{
			SyncClass syncClass = new SyncClass();
            try
            {
                syncClass.CreateOContact(ThisAddIn.addressBookService, ThisAddIn.oStore);
            }
			catch (Exception error)
			{
				syncClass.WriteLog("ERROR: " + error);
			}
		}

		public void DeleteOContactButton(Office.IRibbonControl control)
		{
			SyncClass syncClass = new SyncClass();
            try
            {
                syncClass.DeleteOContact(ThisAddIn.addressBookService, ThisAddIn.oStore);
            }
			catch (Exception error)
			{
				syncClass.WriteLog("ERROR: " + error);
			}
		}

		public void CalendarSyncButton(Office.IRibbonControl control)
        {
			SyncClass syncClass = new SyncClass();
            try
            {
                syncClass.CalendarSync(ThisAddIn.calendarService, ThisAddIn.oStore, ThisAddIn.calendars);
            }
			catch (Exception error)
			{
				syncClass.WriteLog("ERROR: " + error);
			}
		}

		public void AddressBookSyncButton(Office.IRibbonControl control)
		{
			SyncClass syncClass = new SyncClass();
            try
            {
                syncClass.AddressBookSync(ThisAddIn.addressBookService, ThisAddIn.oStore);
            }
			catch (Exception error)
			{
				syncClass.WriteLog("ERROR: " + error);
			}
		}

		public void CalendarResetButton(Office.IRibbonControl control)
		{
			SyncClass syncClass = new SyncClass();
            try
            {
                syncClass.CalendarReset(ThisAddIn.calendarService, ThisAddIn.oStore, ThisAddIn.calendars);
            }
			catch (Exception error)
			{
				syncClass.WriteLog("ERROR: " + error);
			}
		}

		public void AddressBookResetButton(Office.IRibbonControl control)
		{
			SyncClass syncClass = new SyncClass();
            try
            {
				DialogResult result = MessageBox.Show("Vuoi rimuovere tutti i contatti ed eseguire una sincronizzazione completa? Tutti i contatti non sincronizzati con Google andranno persi.", "Attenzione", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
				if (result == DialogResult.OK)
				{
					syncClass.AddressBookReset(ThisAddIn.addressBookService, ThisAddIn.oStore);
					syncClass.AddressBookSync(ThisAddIn.addressBookService, ThisAddIn.oStore);
					MessageBox.Show("Processo completato.", "Processo completato", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			catch (Exception error)
			{
				syncClass.WriteLog("ERROR: " + error);
			}
		}

		public void ShowLogButton(Office.IRibbonControl control)
		{
			SyncClass syncClass = new SyncClass();
            try
            {
                syncClass.ShowLog();
            }
			catch (Exception error)
			{
				syncClass.WriteLog("ERROR: " + error);
			}
		}
		#endregion

		#region Supporti

		private static string GetResourceText(string resourceName)
		{
			Assembly asm = Assembly.GetExecutingAssembly();
			string[] resourceNames = asm.GetManifestResourceNames();
			for (int i = 0; i < resourceNames.Length; ++i)
			{
				if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
				{
					using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
					{
						if (resourceReader != null)
						{
							return resourceReader.ReadToEnd();
						}
					}
				}
			}
			return null;
		}

		#endregion
	}
}
